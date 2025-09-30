# excel_copilot/core/browser_copilot_manager.py

from playwright.sync_api import (
    sync_playwright,
    Page,
    BrowserContext,
    Playwright,
    TimeoutError as PlaywrightTimeoutError,
    Locator,
)
import time
import pyperclip
import sys
import re
from typing import Optional, Callable, List, Tuple, Union

from ..config import COPILOT_USER_DATA_DIR

class BrowserCopilotManager:
    """
    Playwrightを使い、M365 Copilotのチャット画面を操作するクラス。
    初期化、プロンプトの送信、応答の取得を責務に持つ。
    """
    def __init__(self, user_data_dir: str, headless: bool = False):
        self.user_data_dir = user_data_dir
        self.headless = headless
        self.playwright: Optional[Playwright] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def start(self):
        """Playwrightを起動し、Copilotページに接続・初期化する"""
        try:
            self.playwright = sync_playwright().start()
            self.context = self.playwright.chromium.launch_persistent_context(
                self.user_data_dir,
                headless=self.headless,
                slow_mo=50,
                channel="msedge" # Edgeを指定
            )
            self.page = self.context.new_page()
            print("Copilotページに移動します...")
            self.page.goto("https://m365.cloud.microsoft/chat/", timeout=90000)
            print("ページに接続しました。初期化を開始します...")
            self._initialize_copilot_mode()

        except PlaywrightTimeoutError as e:
            print(f"エラー: Copilotページへの接続がタイムアウトしました。URLやネットワークを確認してください。: {e}")
            self.close()
            raise
        except Exception as e:
            print(f"エラー: ブラウザの起動中に予期せぬエラーが発生しました。: {e}")
            self.close()
            raise

    def _wait_for_first_visible(self, description: str, locator_factories: List[Union[Callable[[], Locator], Tuple[str, Callable[[], Locator]]]], timeout: float) -> Locator:
        """指定されたロケーター群の中から最初に可視状態になった要素を取得する"""
        last_exception: Optional[Exception] = None

        for index, factory_entry in enumerate(locator_factories, start=1):
            if isinstance(factory_entry, tuple):
                label, factory = factory_entry
            else:
                label, factory = None, factory_entry

            label_text = label or f"候補 #{index}"
            print(f"{description}: {label_text} を探索しています...")

            try:
                locator = factory()
            except Exception as factory_error:
                last_exception = factory_error
                print(f"{description}: {label_text} のロケーター作成でエラーが発生しました: {factory_error}")
                continue

            timeout_ms = int(timeout) if timeout is not None else 0
            if timeout_ms <= 0:
                timeout_ms = 5000

            deadline = time.monotonic() + timeout_ms / 1000.0
            zero_logged = False
            zero_attempts = 0

            while True:
                remaining_sec = deadline - time.monotonic()
                if remaining_sec <= 0:
                    break

                try:
                    candidate_count = locator.count()
                except Exception as count_error:
                    last_exception = count_error
                    print(f"{description}: {label_text} の要素数取得に失敗しました: {count_error}")
                    break

                if candidate_count == 0:
                    if not zero_logged:
                        print(f"{description}: {label_text} はまだ見つかりません (0 件)。")
                        zero_logged = True
                    zero_attempts += 1
                    limit = max(3, min(10, int(timeout_ms / 2000)))
                    if zero_attempts >= limit:
                        print(f"{description}: {label_text} は引き続き 0 件のため次の候補へ移ります。")
                        break
                    time.sleep(min(0.2, max(0.05, remaining_sec)))
                    continue

                print(f"{description}: {label_text} で {candidate_count} 件の候補を検出しました。")

                for position in range(candidate_count):
                    now_remaining_sec = deadline - time.monotonic()
                    if now_remaining_sec <= 0:
                        break

                    per_attempt_timeout = int(max(1000, min(5000, now_remaining_sec * 1000)))
                    candidate = locator.nth(position)

                    try:
                        if candidate.is_visible():
                            print(f"{description}: {label_text} の候補 #{position + 1} が即座に可視化されました。")
                            return candidate
                    except Exception:
                        pass

                    try:
                        candidate.wait_for(state="visible", timeout=per_attempt_timeout)
                        print(f"{description}: {label_text} の候補 #{position + 1} が可視状態になりました。")
                        return candidate
                    except PlaywrightTimeoutError as wait_error:
                        last_exception = wait_error
                        print(f"{description}: {label_text} の候補 #{position + 1} 可視化待機でタイムアウトしました。")
                        continue
                    except Exception as candidate_error:
                        last_exception = candidate_error
                        print(f"{description}: {label_text} の候補 #{position + 1} でエラー: {candidate_error}")
                        continue

                time.sleep(min(0.2, max(0.05, deadline - time.monotonic())))

        raise RuntimeError(f"{description}が見つかりません。UI が変更された可能性があります。") from last_exception

    def _fill_chat_input(self, chat_input: Locator, prompt: str):
        """Simulate a human paste into the chat editor so Copilot treats URLs normally."""
        if not self.page:
            raise RuntimeError("Page is not initialized.")

        try:
            chat_input.scroll_into_view_if_needed()
        except Exception:
            pass

        chat_input.click()
        modifier = "Meta" if sys.platform == "darwin" else "Control"

        # Clear any existing value
        try:
            chat_input.press(f"{modifier}+A")
        except Exception:
            try:
                self.page.keyboard.press(f"{modifier}+A")
            except Exception:
                pass
        try:
            chat_input.press("Backspace")
        except Exception:
            try:
                self.page.keyboard.press("Backspace")
            except Exception:
                pass

        clipboard_value = prompt.replace("\n", "\r\n")
        pasted = False
        try:
            pyperclip.copy(clipboard_value)
            try:
                chat_input.press(f"{modifier}+V")
            except Exception:
                self.page.keyboard.press(f"{modifier}+V")
        except Exception:
            pasted = False
        else:
            pasted = True

        self.page.wait_for_timeout(400)
        try:
            current_text = chat_input.inner_text().strip()
        except Exception:
            current_text = ""

        if not current_text:
            try:
                chat_input.type(prompt, delay=15)
            except Exception:
                try:
                    self.page.keyboard.type(prompt, delay=15)
                except Exception:
                    pass
            self.page.wait_for_timeout(200)
            try:
                current_text = chat_input.inner_text().strip()
            except Exception:
                current_text = ""

        if not current_text:
            try:
                injected = self.page.evaluate(
                    """
                    (target, value) => {
                        if (!target) return false;
                        target.innerHTML = '';
                        value.split('\n').forEach((line) => {
                            if (line) {
                                const p = document.createElement('p');
                                p.textContent = line;
                                target.appendChild(p);
                            } else {
                                const p = document.createElement('p');
                                p.innerHTML = '<br>';
                                target.appendChild(p);
                            }
                        });
                        const inputEvt = new InputEvent('input', { bubbles: true, data: value, inputType: 'insertText' });
                        target.dispatchEvent(inputEvt);
                        const changeEvt = new Event('change', { bubbles: true });
                        target.dispatchEvent(changeEvt);
                        return true;
                    }
                    """,
                    chat_input,
                    prompt,
                )
            except Exception:
                injected = False
            if injected:
                self.page.wait_for_timeout(200)
                try:
                    current_text = chat_input.inner_text().strip()
                except Exception:
                    current_text = ""

        if not current_text:
            raise RuntimeError("Failed to populate the chat input with the prompt.")

    def _chat_input_locator_factories(self) -> List[Tuple[str, Callable[[], Locator]]]:
        if not self.page:
            raise RuntimeError("ページが初期化されていません。")
        return [
            ("id=m365-chat-editor-target-element", lambda: self.page.locator('#m365-chat-editor-target-element')),
            ("role=combobox aria-label=チャット入力", lambda: self.page.get_by_role('combobox', name='チャット入力')),
            ("aria-describedby^=chat-input-placeholder", lambda: self.page.locator('[aria-describedby^="chat-input-placeholder" i]')),
            ("class=fai-EditorInput__input", lambda: self.page.locator('.fai-EditorInput__input')),
            ("contenteditable role textbox", lambda: self.page.locator('[contenteditable="true"][role="textbox"]')),
            ("role=textbox & contenteditable", lambda: self.page.locator('[role="textbox"][contenteditable="true"]')),
            ("id=prompt-textarea", lambda: self.page.locator('#prompt-textarea')),
            ("class=ProseMirror", lambda: self.page.locator('.ProseMirror')),
            ("ProseMirror contenteditable", lambda: self.page.locator('.ProseMirror[contenteditable="true"]')),
            ("ProseMirror paragraph", lambda: self.page.locator('.ProseMirror').get_by_role('paragraph')),
            ("contenteditable only", lambda: self.page.locator('[contenteditable="true"]')),
            ("role=textbox", lambda: self.page.locator('[role="textbox"]')),
            ("data-testid=chatInput", lambda: self.page.locator('[data-testid="chatInput"]')),
            ("data-testid=chat-input", lambda: self.page.locator('[data-testid="chat-input"]')),
            ("data-testid=threadComposerRichText", lambda: self.page.locator('[data-testid="threadComposerRichText"]')),
            ("threadComposerRichText > contenteditable", lambda: self.page.locator('[data-testid="threadComposerRichText"] [contenteditable="true"]')),
            ("threadComposerRichText paragraph", lambda: self.page.locator('[data-testid="threadComposerRichText"]').get_by_role('paragraph')),
            ("chatInput paragraph", lambda: self.page.locator('[data-testid="chatInput"]').get_by_role('paragraph')),
            ("aria-label contains チャット", lambda: self.page.locator('[aria-label*="チャット"]')),
            ("aria-label contains メッセージ", lambda: self.page.locator('[aria-label*="メッセージ"]')),
            ("aria-label contains message", lambda: self.page.locator('[aria-label*="message"]')),
            ("aria-label contains Copilot", lambda: self.page.locator('[aria-label*="Copilot"]')),
            ("placeholder=質問してみましょう", lambda: self.page.get_by_placeholder('質問してみましょう')),
            ("placeholder=Type a message", lambda: self.page.get_by_placeholder('Type a message')),
            ("placeholder=Ask Copilot", lambda: self.page.get_by_placeholder('Ask Copilot')),
            ("placeholder=Ask your question", lambda: self.page.get_by_placeholder('Ask your question')),
            ("placeholder=Ask me anything", lambda: self.page.get_by_placeholder('Ask me anything')),
            ("placeholder contains prompt", lambda: self.page.get_by_placeholder(re.compile('prompt', re.IGNORECASE))),
            ("aria-label contains prompt", lambda: self.page.locator('[aria-label*="prompt" i]')),
            ("textbox name contains prompt", lambda: self.page.get_by_role('textbox', name=re.compile('prompt', re.IGNORECASE))),
            ("data-testid=promptTextArea", lambda: self.page.locator('[data-testid="promptTextArea"]')),
            ("data-testid=prompt-text-area", lambda: self.page.locator('[data-testid="prompt-text-area"]')),
            ("data-testid=promptInput", lambda: self.page.locator('[data-testid="promptInput"]')),
            ("textbox name=チャット入力欄", lambda: self.page.get_by_role('textbox', name='チャット入力欄')),
            ("textbox name contains チャット", lambda: self.page.get_by_role('textbox', name=re.compile('チャット', re.IGNORECASE))),
            ("textbox name contains message", lambda: self.page.get_by_role('textbox', name=re.compile('message', re.IGNORECASE))),
            ("textbox name contains Copilot", lambda: self.page.get_by_role('textbox', name=re.compile('Copilot', re.IGNORECASE))),
            ("placeholder contains Copilot", lambda: self.page.get_by_placeholder(re.compile('Copilot', re.IGNORECASE))),
            ("placeholder contains メッセージ", lambda: self.page.get_by_placeholder(re.compile('メッセージ', re.IGNORECASE))),
            ("paragraph last", lambda: self.page.get_by_role('paragraph').last),
            ("paragraph first", lambda: self.page.get_by_role('paragraph').first),
            ("paragraph role", lambda: self.page.get_by_role('paragraph')),
        ]

    def _chat_input_additional_locator_factories(self) -> List[Tuple[str, Callable[[], Locator]]]:
        factories: List[Tuple[str, Callable[[], Locator]]] = []
        if not self.page:
            return factories

        factories.extend([
            ("textarea element", lambda: self.page.locator("textarea")),
            ("role=combobox (any)", lambda: self.page.locator("[role=\"combobox\"]")),
            ("aria-label contains message (ci)", lambda: self.page.locator("[aria-label*=\"message\" i]")),
            ("aria-label contains compose", lambda: self.page.locator("[aria-label*=\"compose\" i]")),
            ("aria-label contains prompt (ci)", lambda: self.page.locator("[aria-label*=\"prompt\" i]")),
            ("role=textbox generic", lambda: self.page.locator("[role=\"textbox\"]")),
            ("contenteditable div", lambda: self.page.locator("div[contenteditable='true']")),
            ("contenteditable section", lambda: self.page.locator("section[contenteditable='true']")),
            ("contenteditable span", lambda: self.page.locator("span[contenteditable='true']")),
            ("data-testid contains composer", lambda: self.page.locator("[data-testid*=\"composer\" i]")),
            ("data-testid contains prompt", lambda: self.page.locator("[data-testid*=\"prompt\" i]")),
            ("data-automationid prompt-text-area", lambda: self.page.locator("[data-automationid=\"prompt-text-area\"]")),
            ("data-automationid promptTextArea", lambda: self.page.locator("[data-automationid=\"promptTextArea\"]")),
        ])

        placeholders = [
            "Type a message",
            "Ask Copilot",
            "Ask your question",
            "Ask me anything",
            "Send a message",
            "How can Copilot help you?",
        ]
        for placeholder in placeholders:
            factories.append((f"placeholder={placeholder}", lambda placeholder=placeholder: self.page.get_by_placeholder(placeholder)))

        textbox_names = [
            "Chat input",
            "Copilot prompt",
            "Prompt",
            "Message compose box",
            "Write your message",
            "Write a prompt",
        ]
        for name in textbox_names:
            factories.append((f"textbox name={name}", lambda name=name: self.page.get_by_role('textbox', name=name)))
            factories.append((f"textbox name contains {name}", lambda name=name: self.page.get_by_role('textbox', name=re.compile(re.escape(name), re.IGNORECASE))))

        return factories

    def _iframe_chat_input_locator_factories(self) -> List[Tuple[str, Callable[[], Locator]]]:
        factories: List[Tuple[str, Callable[[], Locator]]] = []
        if not self.page:
            return factories

        try:
            iframe_locator = self.page.locator("iframe")
            iframe_count = iframe_locator.count()
        except Exception as iframe_error:
            print(f"チャット入力欄: iframe探索で警告: {iframe_error}")
            return factories

        placeholders = [
            "Type a message",
            "Ask Copilot",
            "Ask your question",
            "Ask me anything",
            "Send a message",
            "How can Copilot help you?",
        ]

        for idx in range(iframe_count):
            try:
                frame_locator = iframe_locator.nth(idx)
                frame = frame_locator.content_frame()
            except Exception as frame_error:
                print(f"チャット入力欄: iframe #{idx + 1} の content_frame 取得に失敗しました: {frame_error}")
                continue

            if not frame:
                continue

            factories.extend([
                (f"iframe#{idx + 1} contenteditable role textbox", lambda frame=frame: frame.locator('[contenteditable=\"true\"][role=\"textbox\"]')),
                (f"iframe#{idx + 1} contenteditable", lambda frame=frame: frame.locator('[contenteditable=\"true\"]')),
                (f"iframe#{idx + 1} textbox", lambda frame=frame: frame.locator('[role=\"textbox\"]')),
                (f"iframe#{idx + 1} textarea", lambda frame=frame: frame.locator('textarea')),
                (f"iframe#{idx + 1} paragraph", lambda frame=frame: frame.get_by_role('paragraph')),
                (f"iframe#{idx + 1} data-testid prompt", lambda frame=frame: frame.locator('[data-testid*=\"prompt\" i]')),
            ])

            for placeholder in placeholders:
                factories.append((f"iframe#{idx + 1} placeholder={placeholder}", lambda frame=frame, placeholder=placeholder: frame.get_by_placeholder(placeholder)))

        return factories

    def _resolve_chat_input_target(self, locator: Locator) -> Locator:
        """チャット欄を操作できる contenteditable コンテナを指すロケーターに正規化する"""
        try:
            enriched = locator.locator("xpath=ancestor-or-self::*[@contenteditable='true'][1]")
            if enriched.count() > 0:
                print("チャット入力欄: contenteditable な親要素にフォーカスを切り替えます。")
                return enriched.first
        except Exception as resolve_error:
            print(f"チャット入力欄: contenteditable 親要素の特定に失敗しました: {resolve_error}")
        return locator

    def _initialize_copilot_mode(self):
        """GPT-5 チャットモードを有効化し、入力欄を準備する"""
        try:
            print("Copilot ページが表示されるのを待機しています...")
            try:
                self.page.wait_for_function(
                    "() => location.href.startsWith('https://m365.cloud.microsoft/chat') && location.search.includes('auth=2')",
                    timeout=180000,
                )
                print("Copilot ページのロードを確認しました。")
            except PlaywrightTimeoutError as wait_error:
                current_url = self.page.url if self.page else ""
                print(f"警告: Copilot ページのURLが期待値に到達していません (現在: {current_url})。引き続き初期化を試みます。詳細: {wait_error}")
            self.page.wait_for_load_state('domcontentloaded')
            self.page.wait_for_timeout(500)

            gpt5_button_factories = [
                ("button.fui-ToggleButton:has-text('GPT-5 を試す')", lambda: self.page.locator("button.fui-ToggleButton:has-text('GPT-5 を試す')")),
                ("button#GPT-5 を試す", lambda: self.page.get_by_role("button", name="GPT-5 を試す")),
                ("button#GPT-5 で質問", lambda: self.page.get_by_role("button", name="GPT-5 で質問")),
                ("button#GPT-5 を使用", lambda: self.page.get_by_role("button", name="GPT-5 を使用")),
                ("button#GPT-5", lambda: self.page.get_by_role("button", name="GPT-5")),
                ("button#Copilot GPT-5", lambda: self.page.get_by_role("button", name="Copilot GPT-5")),
                ("button role regex GPT-5", lambda: self.page.get_by_role("button", name=re.compile(r"GPT[-\s]?5", re.IGNORECASE))),
                ("menuitem regex GPT-5", lambda: self.page.get_by_role("menuitem", name=re.compile(r"GPT[-\s]?5", re.IGNORECASE))),
                ("button:has-text GPT-5", lambda: self.page.locator("button", has_text=re.compile(r"GPT[-\s]?5", re.IGNORECASE))),
                ("aria-label=GPT-5 を試す", lambda: self.page.locator('[aria-label="GPT-5 を試す" i]')),
                ("aria-label contains GPT-5", lambda: self.page.locator('[aria-label*="GPT-5"]')),
                ("data-testid=gpt5Button", lambda: self.page.locator('[data-testid="gpt5Button"]')),
                ("data-testid=GPT5Button", lambda: self.page.locator('[data-testid="GPT5Button"]')),
            ]
            try:
                gpt5_button = self._wait_for_first_visible(
                    "GPT-5 モード切り替えボタン",
                    gpt5_button_factories,
                    timeout=20000,
                )
            except RuntimeError as gpt5_error:
                print("GPT-5 モードボタンが見つからなかったため、既定のモードで続行します。")
                print(gpt5_error)
            else:
                print("GPT-5 モードボタンをクリックします...")
                try:
                    gpt5_button.click()
                    self.page.wait_for_timeout(800)
                except Exception as click_error:
                    print(f"GPT-5 モードボタンのクリックに失敗しました: {click_error}")

            print("チャット入力欄の読み込みを待機しています。サインインが求められる場合はブラウザで完了してください。")
            try:
                chat_input = self._wait_for_first_visible(
                    "チャット入力欄",
                    self._chat_input_locator_factories(),
                    timeout=180000,
                )
            except RuntimeError as chat_error:
                print("チャット入力欄が既定のパターンで見つからなかったため、追加のフォールバック探索を実行します。")
                print(chat_error)
                fallback_factories: List[Tuple[str, Callable[[], Locator]]] = [("paragraph role", lambda: self.page.get_by_role("paragraph"))]
                fallback_factories.extend(self._chat_input_additional_locator_factories())
                fallback_factories.extend(self._iframe_chat_input_locator_factories())

                if not fallback_factories:
                    raise RuntimeError("チャット入力欄に利用可能なフォールバックがありません。")

                try:
                    chat_input = self._wait_for_first_visible(
                        "チャット入力欄 (フォールバック)",
                        fallback_factories,
                        timeout=20000,
                    )
                except RuntimeError as fallback_error:
                    raise RuntimeError("チャット入力欄のフォールバックにも失敗しました。") from fallback_error

            chat_input = self._resolve_chat_input_target(chat_input)

            try:
                chat_input.scroll_into_view_if_needed()
            except Exception:
                pass

            print("チャット入力欄にフォーカスします...")
            try:
                chat_input.click(timeout=5000)
            except TypeError:
                chat_input.click()
            except Exception as click_error:
                print(f"チャット入力欄のクリックに失敗しましたが、続行します: {click_error}")



            self.page.wait_for_timeout(300)
            print("準備完了です。GPT-5 での対話を開始できます。")

        except PlaywrightTimeoutError as e:
            print(f"エラー: チャット入力欄の検出待機でタイムアウトしました。UIの表示が変更された可能性があります: {e}")
            raise
        except Exception as e:
            print(f"エラー: GPT-5 への切り替え処理中に予期せぬエラーが発生しました: {e}")
            raise


    def ask(self, prompt: str) -> str:
        """プロンプトを送信し、Copilotからの応答をクリップボード経由で取得する"""
        if not self.page:
            raise ConnectionError("ブラウザが初期化されていません。start()を呼び出してください。")

        try:
            # 応答のコピーボタンの現在の数を数える
            copy_button_selector = '[data-testid="CopyButtonTestId"]'
            initial_copy_button_count = self.page.locator(copy_button_selector).count()

            # チャット入力欄にプロンプトを入力して送信
            chat_input = self._wait_for_first_visible(
                "チャット入力欄",
                self._chat_input_locator_factories(),
                timeout=45000,
            )

            chat_input = self._resolve_chat_input_target(chat_input)

            try:
                chat_input.scroll_into_view_if_needed()
            except Exception:
                pass

            try:
                chat_input.click(timeout=5000)
            except TypeError:
                chat_input.click()
            except Exception as click_error:
                print(f"チャット入力欄へのフォーカス時に警告: {click_error}")


            self._fill_chat_input(chat_input, prompt)
            # Allow the pasted content to settle before sending
            try:
                self.page.wait_for_timeout(350)
            except Exception:
                time.sleep(0.35)
            send_button = self._wait_for_first_visible(
                "送信ボタン",
                [
                    ("label=送信", lambda: self.page.get_by_label("送信")),
                    ("label=送信 (Ctrl+Enter)", lambda: self.page.get_by_label("送信 (Ctrl+Enter)")),
                    ("button name=送信", lambda: self.page.get_by_role("button", name="送信")),
                    ("button name=送信 (Ctrl+Enter)", lambda: self.page.get_by_role("button", name="送信 (Ctrl+Enter)")),
                    ("button name=Send", lambda: self.page.get_by_role("button", name="Send")),
                    ("label=Send", lambda: self.page.get_by_label("Send")),
                    ("test-id=SendButtonTestId", lambda: self.page.get_by_test_id("SendButtonTestId")),
                ],
                timeout=15000,
            )
            send_button.click()

            # 新しいコピーボタン（＝新しい応答）が出現するのを待つ
            print("Copilotの応答を待っています...")
            new_copy_button_locator = self.page.locator(copy_button_selector).nth(initial_copy_button_count)
            new_copy_button_locator.wait_for(state="visible", timeout=180000) # タイムアウトを3分に延長
            print("応答が完了したと判断しました。")

            # 最新のコピーボタンをクリックして内容を取得
            copy_buttons = self.page.locator(copy_button_selector)
            if copy_buttons.count() > initial_copy_button_count:
                print("応答をクリップボードにコピーします...")
                copy_buttons.last.click()
                time.sleep(0.5)  # クリップボードへの反映を待つ
                response_text = pyperclip.paste().strip()
                # "Thought:" が含まれている場合、そこから後を抽出する
                thought_pos = response_text.find("Thought:")
                if thought_pos != -1:
                    return response_text[thought_pos:]
                return response_text
            else:
                return "エラー: 新しい応答のコピーボタンが見つかりませんでした。"

        except PlaywrightTimeoutError:
            return "エラー: Copilotからの応答がタイムアウトしました。処理が複雑すぎるか、Copilotが応答不能になっている可能性があります。"
        except Exception as e:
            return f"エラー: ブラウザ操作中に予期せぬエラーが発生しました: {e}"

    def close(self):
        """ブラウザとPlaywrightセッションを安全に閉じる"""
        if self.context:
            try:
                self.context.close()
                print("ブラウザコンテキストを閉じました。")
            except Exception as e:
                print(f"ブラウザコンテキストを閉じる際にエラーが発生しました: {e}")
        if self.playwright:
            try:
                self.playwright.stop()
                print("Playwrightを停止しました。")
            except Exception as e:
                print(f"Playwrightを停止する際にエラーが発生しました: {e}")
        self.page = None
        self.context = None
        self.playwright = None
