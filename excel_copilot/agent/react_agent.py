import json
import inspect
import re
import threading
from typing import Generator, List, Dict, Any, Optional, Tuple, Callable

from excel_copilot.config import MAX_ITERATIONS, HISTORY_MAX_MESSAGES
from excel_copilot.core.exceptions import LLMResponseError, ToolExecutionError, UserStopRequested
from excel_copilot.agent.prompts import CopilotMode, build_system_prompt
from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.excel_manager import ExcelManager, ExcelConnectionError
from excel_copilot.tools.actions import ExcelActions

# 翻訳などのタスクで、AIが一度に処理しようとするテキストの最大文字数。
# これを超えると、AgentはAIに分割処理を促すフィードバックを返す。
MAX_PROCESSING_TEXT_LENGTH = 4000

class ReActAgent:
    """
    ReAct (Reasoning and Acting) フレームワークに基づいたAIエージェント。
    UIからの停止要求(stop_event)に対応し、構造化された辞書をyieldします。
    """
    def __init__(self, tools: List[callable], tool_schemas: List[Dict], browser_manager: BrowserCopilotManager, sheet_name: Optional[str] = None, mode: CopilotMode = CopilotMode.TRANSLATION, progress_callback: Optional[Callable[[str], None]] = None):
        self.browser_manager = browser_manager
        self.sheet_name = sheet_name
        self.mode = mode
        self.tools = {tool.__name__: tool for tool in tools}
        self.tool_schemas_str = json.dumps(tool_schemas, indent=2, ensure_ascii=False)
        self.system_prompt = build_system_prompt(self.mode, self.tool_schemas_str)
        self.messages: List[Dict[str, str]] = []
        self.progress_callback = progress_callback

    def reset(self):
        """Reset the conversation history for a fresh run."""

        self.messages = []

    def set_mode(self, mode: CopilotMode):
        """Switch the agent to a new mode and rebuild the system prompt."""

        if mode == self.mode:
            return
        self.mode = mode
        self.system_prompt = build_system_prompt(self.mode, self.tool_schemas_str)
        self.reset()

    def _initialize_messages(self, user_query: str):
        """会話履歴を初期化する"""
        self.messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_query}
        ]

    def _build_prompt(self) -> str:
        """LLMに渡すプロンプトを構築する"""
        # 古い履歴を省略
        if len(self.messages) > HISTORY_MAX_MESSAGES * 2 + 1:
            recent_history = self.messages[-(HISTORY_MAX_MESSAGES * 2):]
            prompt_messages = [self.messages[0]] + recent_history
        else:
            prompt_messages = self.messages

        prompt = "\n".join([f"{msg['role']}: {msg['content']}" for msg in prompt_messages])
        prompt += "\n\nassistant:"
        return prompt

    def _parse_llm_output(self, response: str) -> Tuple[str, Optional[str], Optional[str]]:
        """LLMの出力を Thought, Action, Final Answer に分割する。"""

        def _extract_json_payload(text: str) -> Optional[str]:
            """Return the first JSON payload in text, or None if not found."""
            if not text:
                return None
            leading_trimmed = text.lstrip()
            if not leading_trimmed or leading_trimmed[0] not in "[{":
                return None
            try:
                decoder = json.JSONDecoder()
                _, end_idx = decoder.raw_decode(leading_trimmed)
            except json.JSONDecodeError:
                return None
            return leading_trimmed[:end_idx]

        response = (response or "").strip()
        if not response:
            raise LLMResponseError("LLMから空の応答が返されました。")

        # JSONのみで構成されたレスポンス（Thought/Actionラベル省略）に対応
        json_only_payload = _extract_json_payload(response)
        if json_only_payload and json_only_payload == response:
            return "", json_only_payload, None

        thought = ""
        action_str: Optional[str] = None
        final_answer: Optional[str] = None

        colon_pattern = r"\s*[:：]"
        final_answer_match = re.search(rf"Final Answer{colon_pattern}", response, re.IGNORECASE)
        if final_answer_match:
            thought_match = re.search(
                rf"Thought{colon_pattern}",
                response[:final_answer_match.start()],
                re.IGNORECASE,
            )
            if thought_match:
                thought = response[thought_match.end():final_answer_match.start()].strip()
            final_answer = response[final_answer_match.end():].strip()
            return thought, None, final_answer

        action_match = re.search(rf"Action{colon_pattern}", response, re.IGNORECASE)
        thought_match = re.search(rf"Thought{colon_pattern}", response, re.IGNORECASE)

        if thought_match:
            thought_end = action_match.start() if action_match else len(response)
            thought = response[thought_match.end():thought_end].strip()
        elif action_match:
            # Thoughtラベルが無くActionのみのケースを許容する
            thought = response[:action_match.start()].strip()
        else:
            raise LLMResponseError("応答形式が不正です。'Thought:' または 'Final Answer:' が見つかりません。")

        if action_match:
            action_str_raw = response[action_match.end():].strip()
            json_payload = _extract_json_payload(action_str_raw)
            if not json_payload:
                raise LLMResponseError("Actionブロック内にJSONが見つかりませんでした。")
            action_str = json_payload

        return thought, action_str, final_answer

    def _execute_tool(self, action_json_str: str, excel_actions: ExcelActions) -> Any:
        """ツールを実行する"""
        try:
            action_data = json.loads(action_json_str)
            tool_name = action_data.get("tool_name") or action_data.get("toolname")
            arguments = action_data.get("arguments", {})
        except json.JSONDecodeError as e:
            raise ToolExecutionError(f"ActionのJSON形式が不正です: {e}")

        if not tool_name or tool_name not in self.tools:
            raise ToolExecutionError(f"ツール '{tool_name}' は存在しません。")

        tool_function = self.tools[tool_name]
        sig = inspect.signature(tool_function)

        # 必要な引数を自動的に注入
        if 'actions' in sig.parameters:
            arguments['actions'] = excel_actions
        if 'browser_manager' in sig.parameters:
            arguments['browser_manager'] = self.browser_manager
        if 'sheetname' in sig.parameters and 'sheetname' not in arguments and self.sheet_name:
            arguments['sheetname'] = self.sheet_name

        try:
            result = tool_function(**arguments)
            if hasattr(excel_actions, 'consume_progress_messages'):
                excel_actions.consume_progress_messages()
            return result
        except Exception as e:
            # エラーのスタックトレースも表示するとデバッグに役立つ
            import traceback
            print(f"Tool execution error: {traceback.format_exc()}")
            raise ToolExecutionError(f"ツール '{tool_name}' の実行に失敗しました: {e}")

    def run(self, user_query: str, stop_event: threading.Event) -> Generator[Dict[str, Any], None, None]:
        """エージェントのメイン実行ループ"""
        self._initialize_messages(user_query)

        try:
            with ExcelManager() as manager:
                excel_actions = ExcelActions(manager, progress_callback=self.progress_callback)
                for i in range(MAX_ITERATIONS):
                    if stop_event.is_set():
                        yield {"type": "info", "content": "処理が中断されました。"}
                        return

                    yield {"type": "status", "content": f"思考サイクル {i + 1}/{MAX_ITERATIONS}..."}
                    prompt = self._build_prompt()

                    try:
                        response_content = self.browser_manager.ask(prompt, stop_event=stop_event)
                        if response_content.startswith("エラー:"):
                            raise LLMResponseError(response_content)
                    except UserStopRequested:
                        yield {"type": "info", "content": "ユーザーの操作で処理が中断されました。"}
                        return
                    except Exception as e:
                        yield {"type": "error", "content": f"Copilotとの通信に失敗しました: {e}"}
                        return

                    if stop_event.is_set():
                        yield {"type": "info", "content": "処理が中断されました。"}
                        return

                    try:
                        thought, action_json_str, final_answer = self._parse_llm_output(response_content)

                        yield {"type": "thought", "content": thought}
                        self.messages.append({"role": "assistant", "content": f"Thought: {thought}"})

                        if action_json_str:
                            yield {"type": "action", "content": action_json_str}
                            self.messages[-1]["content"] += f"\nAction: {action_json_str}"

                            observation = self._execute_tool(action_json_str, excel_actions)
                            
                            # 読み取り結果が長すぎる場合のフィードバック
                            is_read_tool = "read" in json.loads(action_json_str).get("tool_name", "")
                            is_large_output = isinstance(observation, str) and len(observation) > MAX_PROCESSING_TEXT_LENGTH
                            if is_read_tool and is_large_output:
                                feedback = (
                                    f"エラー: 読み込んだデータが長すぎます({len(observation)}文字)。"
                                    "一度に処理するには大きすぎるため、操作は中断されました。"
                                    "思考を修正し、より小さな範囲に分割して処理を続けてください。"
                                )
                                yield {"type": "error", "content": feedback}
                                self.messages.append({"role": "user", "content": f"Observation: {feedback}"})
                                continue

                            yield {"type": "observation", "content": str(observation)}
                            self.messages.append({"role": "user", "content": f"Observation: {observation}"})

                        elif final_answer:
                            yield {"type": "final_answer", "content": final_answer}
                            self.messages.append({"role": "assistant", "content": f"Final Answer: {final_answer}"})
                            return

                        else:
                            raise LLMResponseError("LLMの応答にActionまたはFinal Answerが含まれていません。")

                    except (LLMResponseError, ToolExecutionError) as e:
                        error_feedback = f"エラーが発生しました: {e}。思考を修正し、別のアプローチを試してください。"
                        yield {"type": "error", "content": error_feedback}
                        self.messages.append({"role": "user", "content": f"Error: {error_feedback}"})

        except ExcelConnectionError as e:
            yield {"type": "error", "content": f"Excelに接続できませんでした: {e}"}
            return
        except Exception as e:
            import traceback
            yield {"type": "error", "content": f"予期せぬエラーが発生しました: {e}\n{traceback.format_exc()}"}
            return

        yield {"type": "info", "content": f"最大反復回数 ({MAX_ITERATIONS}回) に到達しました。"}
