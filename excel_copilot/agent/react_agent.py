import json
import inspect
import re
import threading
from pathlib import Path
from urllib.parse import urlparse
from typing import Generator, List, Dict, Any, Optional, Tuple, Callable, Set

from excel_copilot.config import MAX_ITERATIONS, HISTORY_MAX_MESSAGES
from excel_copilot.core.exceptions import LLMResponseError, ToolExecutionError, UserStopRequested
from excel_copilot.agent.prompts import CopilotMode, build_system_prompt
from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.excel_manager import ExcelManager, ExcelConnectionError
from excel_copilot.tools.actions import ExcelActions

# 翻訳などのタスクで、AIが一度に処理しようとするテキストの最大文字数。
# これを超えると、AgentはAIに分割処理を促すフィードバックを返す。
MAX_PROCESSING_TEXT_LENGTH = 4000

_COMPLETION_TOKENS_JA = (
    "完了",
    "出力",
    "書き込み",
    "反映",
    "更新",
    "記入",
)

_COMPLETION_TOKENS_EN = (
    "translation",
    "translated",
    "written",
    "wrote",
    "output",
    "applied",
    "completed",
    "done",
    "updated",
)

_TRAILING_ERROR_PREFIXES = ("Error:", "エラー:")


def _interpret_completion_response(response: str) -> Optional[str]:
    """Return a cleaned completion message when tags are missing."""

    sanitized = (response or "").strip()
    if not sanitized:
        return None

    lowered = sanitized.lower()
    has_completion_marker = any(token in sanitized for token in _COMPLETION_TOKENS_JA) or any(
        token in lowered for token in _COMPLETION_TOKENS_EN
    )
    mentions_translation = "翻訳" in sanitized or "translate" in lowered or "translation" in lowered
    mentions_cells = bool(re.search(r"[A-Za-z]{1,3}\d+", sanitized))

    if not has_completion_marker and not (mentions_translation and mentions_cells):
        return None

    cleaned = sanitized
    for prefix in _TRAILING_ERROR_PREFIXES:
        marker_index = cleaned.find(prefix)
        if marker_index != -1:
            candidate = cleaned[:marker_index].strip()
            if candidate:
                cleaned = candidate
            break

    return cleaned or None

class ReActAgent:
    """
    ReAct (Reasoning and Acting) フレームワークに基づいたAIエージェント。
    UIからの停止要求(stop_event)に対応し、構造化された辞書をyieldします。
    """
    def __init__(
        self,
        tools: List[callable],
        tool_schemas: List[Dict],
        browser_manager: BrowserCopilotManager,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        mode: CopilotMode = CopilotMode.TRANSLATION,
        progress_callback: Optional[Callable[[str], None]] = None,
    ):
        self.browser_manager = browser_manager
        self.sheet_name = sheet_name
        self.workbook_name = workbook_name
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

    def set_workbook(self, workbook_name: Optional[str]):
        self.workbook_name = workbook_name

    def _initialize_messages(self, user_query: str):
        """会話履歴を初期化する"""
        self.messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_query}
        ]

    def _build_prompt(self, stage_instruction: Optional[str] = None) -> str:
        """LLMに渡すプロンプトを組み立てる"""
        # 過去ログを絞る処理
        if len(self.messages) > HISTORY_MAX_MESSAGES * 2 + 1:
            recent_history = self.messages[-(HISTORY_MAX_MESSAGES * 2):]
            prompt_messages = [self.messages[0]] + recent_history
        else:
            prompt_messages = self.messages

        prompt_lines = [f"{msg['role']}: {msg['content']}" for msg in prompt_messages]
        if stage_instruction:
            prompt_lines.append(f"user: {stage_instruction}")
        prompt = "\n".join(prompt_lines)
        prompt += "\n\nassistant:"
        return prompt

    def _build_stage_instruction(self) -> Optional[str]:
        tool_names = ", ".join([f"`{name}`" for name in self.tools.keys()]) or "the available tool"
        last_user_message = next((msg for msg in reversed(self.messages) if msg.get("role") == "user"), None)
        last_content = last_user_message.get("content", "") if last_user_message else ""
        awaiting_followup = last_content.startswith("Observation:")
        expecting_action = (not awaiting_followup) or last_content.startswith("Error:")

        if expecting_action:
            return (
                "Respond with `Thought:` explaining your next step, then emit exactly one `Action:` that calls "
                f"{tool_names} using JSON arguments. Do not provide `Final Answer:` yet."
            )

        return (
            "Review the latest Observation in `Thought:`. If the task is complete, return a concise `Final Answer:`. "
            "Otherwise, call a tool again with a fresh `Action`."
        )

    def _parse_llm_output(self, response: str) -> Tuple[str, Optional[str], Optional[str]]:
        """LLMの出力を Thought, Action, Final Answer に分割する。"""

        def _extract_json_payload(text: str) -> Optional[str]:
            """Return the first JSON object/array found anywhere in the text."""

            if not text:
                return None

            decoder = json.JSONDecoder()
            for match in re.finditer(r"[\[{]", text):
                start_idx = match.start()
                candidate = text[start_idx:]
                try:
                    _, relative_end = decoder.raw_decode(candidate)
                except json.JSONDecodeError:
                    continue
                end_idx = start_idx + relative_end
                return text[start_idx:end_idx]
            return None

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
        action_match = re.search(rf"Action{colon_pattern}", response, re.IGNORECASE)
        thought_match = re.search(rf"Thought{colon_pattern}", response, re.IGNORECASE)
        final_answer_match = re.search(rf"Final Answer{colon_pattern}", response, re.IGNORECASE)

        if action_match:
            if thought_match and thought_match.start() < action_match.start():
                thought = response[thought_match.end():action_match.start()].strip()
            elif thought_match:
                thought = response[thought_match.end():action_match.start()].strip()
            else:
                thought = response[:action_match.start()].strip()

            action_str_raw = response[action_match.end():].strip()
            json_payload = _extract_json_payload(action_str_raw)
            if not json_payload:
                raise LLMResponseError("ActionブロックからJSONが検出できませんでした。")
            return thought, json_payload, None

        if final_answer_match:
            if thought_match and thought_match.start() < final_answer_match.start():
                thought = response[thought_match.end():final_answer_match.start()].strip()
            elif thought_match:
                thought = response[thought_match.end():final_answer_match.start()].strip()
            else:
                thought = response[:final_answer_match.start()].strip()
            final_answer = response[final_answer_match.end():].strip()
            return thought, None, final_answer

        if thought_match:
            thought = response[thought_match.end():].strip()
            return thought, None, None

        final_answer_candidate = _interpret_completion_response(response)
        if final_answer_candidate is not None:
            return "", None, final_answer_candidate

        preview = response.replace("\n", " ").strip()
        print(
            f"LLM parse error raw response: {preview[:300]}{'…' if len(preview) > 300 else ''}",
            flush=True,
        )
        raise LLMResponseError("応答解析に失敗しました。'Thought:' または 'Final Answer:' が見つかりません。")

    def _ensure_reference_support(self, arguments: Dict[str, Any], excel_actions: ExcelActions) -> None:
        """Ensure translate_range_with_references receives usable supporting material."""
        def _normalize(value: Any) -> List[str]:
            if not value:
                return []
            if isinstance(value, str):
                candidates = [value]
            elif isinstance(value, (tuple, set, list)):
                candidates = list(value)
            else:
                candidates = [value]
            normalized: List[str] = []
            for candidate in candidates:
                if isinstance(candidate, str):
                    stripped = candidate.strip()
                    if stripped:
                        normalized.append(stripped)
            return normalized

        def _dedupe(urls: List[str]) -> List[str]:
            seen: Set[str] = set()
            result: List[str] = []
            for url in urls:
                if url not in seen:
                    seen.add(url)
                    result.append(url)
            return result

        source_urls = _normalize(arguments.get("source_reference_urls"))
        target_urls = _normalize(arguments.get("target_reference_urls"))
        legacy_urls = _normalize(arguments.get("reference_urls"))

        if legacy_urls:
            if not source_urls:
                source_urls = list(legacy_urls)
            else:
                for url in legacy_urls:
                    if url not in source_urls:
                        source_urls.append(url)

        if not source_urls and not target_urls:
            tokens = self._extract_reference_tokens()
            inferred_urls = self._resolve_reference_hints(tokens, excel_actions)
            if inferred_urls:
                source_urls.extend(inferred_urls)

        source_urls = _dedupe(source_urls)
        target_urls = _dedupe(target_urls)

        if source_urls:
            arguments["source_reference_urls"] = source_urls
        else:
            arguments.pop("source_reference_urls", None)

        if target_urls:
            arguments["target_reference_urls"] = target_urls
        else:
            arguments.pop("target_reference_urls", None)

        # Legacy key should no longer be used once source URLs are populated.
        if "reference_urls" in arguments:
            if source_urls:
                arguments.pop("reference_urls", None)
            elif legacy_urls:
                arguments["reference_urls"] = legacy_urls
            else:
                arguments.pop("reference_urls", None)

    def _extract_reference_tokens(self) -> List[str]:
        """Collect potential reference hints (URLs / filenames) from conversation history."""
        url_pattern = re.compile(r"https?://[^\s\u3000<>\"'）)]+", re.IGNORECASE)
        file_pattern = re.compile(
            r"[^\s\u3000<>\"']+\.(?:pdf|docx|doc|xlsx|xlsm|xls|csv|txt|md)",
            re.IGNORECASE,
        )
        tokens: List[str] = []
        seen: Set[str] = set()

        for message in reversed(self.messages):
            if message.get("role") != "user":
                continue
            content = message.get("content") or ""
            if not content:
                continue
            normalized = content.replace("　", " ")
            for match in url_pattern.finditer(normalized):
                candidate = self._trim_trailing_punctuation(match.group(0))
                if candidate and candidate not in seen:
                    seen.add(candidate)
                    tokens.append(candidate)
            for match in file_pattern.finditer(normalized):
                candidate = self._trim_trailing_punctuation(match.group(0))
                if candidate and candidate not in seen:
                    seen.add(candidate)
                    tokens.append(candidate)

        return list(reversed(tokens))

    @staticmethod
    def _trim_trailing_punctuation(value: str) -> str:
        trimmed = value.strip()
        trailing_chars = ".,;:)]}>】〕）｣\"'"
        while trimmed and trimmed[-1] in trailing_chars:
            trimmed = trimmed[:-1].rstrip()
        return trimmed

    @staticmethod
    def _strip_enclosing_quotes(value: str) -> str:
        trimmed = value.strip()
        if len(trimmed) >= 2 and trimmed[0] == trimmed[-1] and trimmed[0] in {"'", '"'}:
            inner = trimmed[1:-1]
            if trimmed[0] == "'":
                return inner.replace("''", "'")
            return inner
        return trimmed

    def _resolve_reference_hints(self, tokens: List[str], excel_actions: ExcelActions) -> List[str]:
        """Convert extracted tokens into usable reference URLs (HTTP/file)."""
        if not tokens:
            return []

        workbook_dir: Optional[Path] = None
        try:
            workbook_path = getattr(excel_actions.book, "fullname", "") or getattr(
                excel_actions.book, "full_name", ""
            )
            if workbook_path:
                workbook_dir = Path(workbook_path).resolve().parent
        except Exception:
            workbook_dir = None

        resolved_urls: List[str] = []
        seen_urls: Set[str] = set()

        search_roots: List[Path] = []
        try:
            search_roots.append(Path.cwd())
            search_roots.append(Path.cwd().parent)
        except Exception:
            pass

        if workbook_dir:
            search_roots.append(workbook_dir)

        for token in tokens:
            cleaned = self._strip_enclosing_quotes(token)
            if not cleaned:
                continue

            parsed = urlparse(cleaned)
            if parsed.scheme and (parsed.netloc or (parsed.scheme.lower() == "file" and parsed.path)):
                normalized = cleaned.strip()
                if normalized not in seen_urls:
                    seen_urls.add(normalized)
                    resolved_urls.append(normalized)
                continue

            candidate_path = Path(cleaned)
            candidate_paths: List[Path] = []
            if candidate_path.is_absolute():
                candidate_paths.append(candidate_path)
            else:
                for root in search_roots:
                    try:
                        candidate_paths.append((root / candidate_path).expanduser())
                    except Exception:
                        continue

            for path_candidate in candidate_paths:
                try:
                    if path_candidate.exists():
                        uri = path_candidate.resolve().as_uri()
                        if uri not in seen_urls:
                            seen_urls.add(uri)
                            resolved_urls.append(uri)
                        break
                except Exception:
                    continue

        return resolved_urls

    def _execute_tool(self, action_json_str: str, excel_actions: ExcelActions, stop_event: threading.Event) -> Any:
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
        if 'stop_event' in sig.parameters:
            arguments['stop_event'] = stop_event

        if tool_name == "translate_range_with_references":
            self._ensure_reference_support(arguments, excel_actions)

        try:
            result = tool_function(**arguments)
            if hasattr(excel_actions, 'consume_progress_messages'):
                excel_actions.consume_progress_messages()
            return result
        except UserStopRequested:
            raise
        except Exception as e:
            # エラーのスタックトレースも表示するとデバッグに役立つ
            import traceback
            print(f"Tool execution error: {traceback.format_exc()}")
            raise ToolExecutionError(f"ツール '{tool_name}' の実行に失敗しました: {e}")

    def run(self, user_query: str, stop_event: threading.Event) -> Generator[Dict[str, Any], None, None]:
        """エージェントのメイン実行ループ"""
        self._initialize_messages(user_query)

        try:
            with ExcelManager(self.workbook_name) as manager:
                excel_actions = ExcelActions(manager, progress_callback=self.progress_callback)
                for i in range(MAX_ITERATIONS):
                    if stop_event.is_set():
                        yield {"type": "info", "content": "処理が中断されました。"}
                        return

                    yield {"type": "status", "content": f"思考サイクル {i + 1}/{MAX_ITERATIONS}..."}
                    stage_instruction = self._build_stage_instruction()
                    prompt = self._build_prompt(stage_instruction)

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

                            try:
                                observation = self._execute_tool(action_json_str, excel_actions, stop_event)
                            except UserStopRequested:
                                yield {"type": "info", "content": "ユーザーの操作で処理が中断されました。"}
                                return

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
