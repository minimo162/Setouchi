# excel_copilot/core/browser_copilot_manager.py

from playwright.sync_api import (
    sync_playwright,
    Page,
    BrowserContext,
    Playwright,
    TimeoutError as PlaywrightTimeoutError,
    Locator,
)
import logging
import time
import pyperclip
import sys
import re
from pathlib import Path
from typing import Optional, Callable, List, Tuple, Union
from threading import Event

from ..config import (
    COPILOT_BROWSER_CHANNELS,
    COPILOT_SLOW_MO_MS,
    COPILOT_PAGE_GOTO_TIMEOUT_MS,
    COPILOT_SUPPRESS_BROWSER_FOCUS,
    COPILOT_SUPPRESS_BROWSER_FOCUS_LEFT,
    COPILOT_SUPPRESS_BROWSER_FOCUS_TOP,
)
from .exceptions import UserStopRequested

class BrowserCopilotManager:
    """
    Playwrightを使い、M365 Copilotのチャット画面を操作するクラス。
    初期化、プロンプトの送信、応答の取得を責務に持つ。
    """
    def __init__(
        self,
        user_data_dir: str,
        headless: bool = False,
        browser_channels: Optional[List[str]] = None,
        goto_timeout_ms: Optional[int] = None,
        slow_mo_ms: Optional[int] = None,
    ):
        self.user_data_dir = user_data_dir
        self.headless = headless
        self.browser_channels = list(dict.fromkeys(browser_channels or COPILOT_BROWSER_CHANNELS))
        self.goto_timeout_ms = goto_timeout_ms or COPILOT_PAGE_GOTO_TIMEOUT_MS
        self.slow_mo_ms = slow_mo_ms if slow_mo_ms is not None else COPILOT_SLOW_MO_MS
        self.playwright: Optional[Playwright] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self._logger = logging.getLogger(__name__)
        self._focus_suppressed_once = False

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def start(self):
        """Playwrightを起動し、Copilotページに接続・初期化する"""
        try:
            user_data_path = Path(self.user_data_dir).expanduser()
            user_data_path.mkdir(parents=True, exist_ok=True)
            self._logger.debug("Using browser profile directory at %s", user_data_path)

            self.playwright = sync_playwright().start()
            self.context = self._launch_persistent_context(user_data_path)

            if self.context.pages:
                self.page = self.context.pages[0]
            else:
                self.page = self.context.new_page()

            if not self.page:
                raise RuntimeError("ブラウザページの初期化に失敗しました。")

            self._focus_suppressed_once = False
            if COPILOT_SUPPRESS_BROWSER_FOCUS:
                self._suppress_browser_focus()
            else:
                self._ensure_browser_visible()
            try:
                self.page.set_default_timeout(self.goto_timeout_ms)
            except Exception:
                pass

            self._logger.info("Copilotページに移動します...")
            self.page.goto("https://m365.cloud.microsoft/chat/", timeout=self.goto_timeout_ms)
            self._logger.info("ページに接続しました。初期化を開始します...")
            self._initialize_copilot_mode()
            if COPILOT_SUPPRESS_BROWSER_FOCUS:
                self._suppress_browser_focus()
            else:
                self._ensure_browser_visible()

        except PlaywrightTimeoutError as e:
            message = (
                f"エラー: Copilotページへの接続がタイムアウトしました。URLやネットワークを確認してください。: {e}"
            )
            self._logger.error(message)
            print(message)
            self.close()
            raise
        except Exception as e:
            message = f"エラー: ブラウザの起動中に予期せぬエラーが発生しました。: {e}"
            self._logger.error(message)
            print(message)
            self.close()
            raise

    def _launch_persistent_context(self, user_data_path: Path) -> BrowserContext:
        if not self.playwright:
            raise RuntimeError("Playwrightが初期化されていません。")

        attempted_channels: List[Optional[str]] = []
        for channel in self.browser_channels:
            if channel not in attempted_channels:
                attempted_channels.append(channel)
        if None not in attempted_channels:
            attempted_channels.append(None)

        last_exception: Optional[Exception] = None
        for channel in attempted_channels:
            launch_kwargs = {
                "headless": self.headless,
                "slow_mo": max(0, int(self.slow_mo_ms)),
            }
            if channel:
                launch_kwargs["channel"] = channel
                self._logger.info("Chromium persistent context を起動します (channel=%s, headless=%s)", channel, self.headless)
            else:
                self._logger.info("Chromium persistent context を起動します (channel=auto, headless=%s)", self.headless)
            try:
                context = self.playwright.chromium.launch_persistent_context(
                    str(user_data_path),
                    **launch_kwargs,
                )
                self._logger.info("Chromium の起動に成功しました (channel=%s)", channel or "auto")
                return context
            except Exception as exc:
                last_exception = exc
                self._logger.warning("Channel=%s でのブラウザ起動に失敗しました: %s", channel or "auto", exc)
                continue

        raise RuntimeError("利用可能なブラウザチャネルでの起動に失敗しました。") from last_exception

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

    def _read_chat_input_text(self, chat_input: Locator) -> str:
        """Return current text content from the chat input for emptiness checks."""
        if not self.page:
            return ""
        try:
            content = self.page.evaluate(
                """
                (target) => {
                    if (!target) {
                        return '';
                    }

                    const readString = (value) => (typeof value === 'string' ? value : '');

                    const directValue = readString(target.value);
                    if (directValue.trim()) {
                        return directValue;
                    }

                    const inner = readString(target.innerText);
                    if (inner.trim()) {
                        return inner;
                    }

                    const textContent = readString(target.textContent);
                    if (textContent.trim()) {
                        return textContent;
                    }

                    const nestedEditable = target.querySelector('[contenteditable=\"true\"], textarea, [role=\"textbox\"]');
                    if (nestedEditable) {
                        const nestedValue =
                            readString(nestedEditable.value) ||
                            readString(nestedEditable.innerText) ||
                            readString(nestedEditable.textContent);
                        if (nestedValue.trim()) {
                            return nestedValue;
                        }
                    }

                    const doc = target.ownerDocument || document;
                    if (!doc || !doc.createTreeWalker) {
                        return '';
                    }

                    const walker = doc.createTreeWalker(target, NodeFilter.SHOW_TEXT, null);
                    let buffer = '';
                    while (walker.nextNode()) {
                        const node = walker.currentNode;
                        if (!node || !node.nodeValue) {
                            continue;
                        }
                        const parentEl = node.parentElement;
                        if (parentEl) {
                            const placeholder = parentEl.getAttribute && parentEl.getAttribute('data-placeholder');
                            if (placeholder === 'true') {
                                continue;
                            }
                            const ariaHidden = parentEl.getAttribute && parentEl.getAttribute('aria-hidden');
                            if (
                                ariaHidden === 'true'
                                && !(parentEl.matches && parentEl.matches('[role=\"textbox\"], textarea, [contenteditable=\"true\"]'))
                            ) {
                                continue;
                            }
                        }
                        buffer += node.nodeValue;
                    }
                    return buffer;
                }
                """,
                chat_input,
            )
            if content:
                return str(content).strip()
        except Exception:
            try:
                return chat_input.inner_text().strip()
            except Exception:
                return ""
        return ""

    def _normalize_prompt_text(self, value: str) -> str:
        """Normalize prompt text for reliable comparison."""
        sanitized = (value or "").replace("\r\n", "\n").replace("\r", "\n")
        sanitized = sanitized.replace("\u200b", "")
        return sanitized.rstrip("\n")

    def _clear_chat_input(self, chat_input: Locator):
        """Clear existing content from the chat input using keyboard shortcuts."""
        modifier = "Meta" if sys.platform == "darwin" else "Control"

        try:
            chat_input.focus()
        except Exception:
            pass

        try:
            chat_input.press(f"{modifier}+A")
        except Exception:
            try:
                if self.page:
                    self.page.keyboard.press(f"{modifier}+A")
            except Exception:
                pass

        try:
            chat_input.press("Backspace")
        except Exception:
            try:
                if self.page:
                    self.page.keyboard.press("Backspace")
            except Exception:
                pass


    def _type_prompt_with_soft_returns(self, chat_input: Locator, prompt: str) -> bool:
        """Type text into the chat input while using soft returns for newlines."""
        if not self.page:
            return False

        segments = prompt.split("\n")
        try:
            for idx, segment in enumerate(segments):
                if segment:
                    inserted = False
                    try:
                        self.page.keyboard.insert_text(segment)
                        inserted = True
                    except Exception as insert_error:
                        print(f"Warning: insert_text failed: {insert_error}")
                    if not inserted:
                        try:
                            chat_input.type(segment, delay=0.01)
                            inserted = True
                        except Exception as type_error:
                            print(f"Warning: Locator.type failed: {type_error}")
                    if not inserted:
                        return False

                if idx < len(segments) - 1:
                    soft_return_sequences = ["Shift+Enter", "Shift+Return", "Alt+Enter"]
                    success = False
                    for sequence in soft_return_sequences:
                        try:
                            chat_input.press(sequence)
                            success = True
                            break
                        except Exception:
                            try:
                                self.page.keyboard.press(sequence)
                                success = True
                                break
                            except Exception:
                                continue
                    if not success:
                        print("Warning: failed to send a soft return for newline.")
                        return False
            return True
        except Exception as err:
            print(f"Warning: soft-return fallback hit an error: {err}")
            return False

    def _fill_chat_input(self, chat_input: Locator, prompt: str) -> str:
        """Simulate a human paste into the chat editor so Copilot treats URLs normally."""
        if not self.page:
            raise RuntimeError("Page is not initialized.")

        try:
            chat_input.scroll_into_view_if_needed()
        except Exception:
            pass

        chat_input.click()
        try:
            chat_input.focus()
        except Exception as focus_error:
            print(f"チャット入力欄: focus の適用に失敗しました: {focus_error}")

        self._clear_chat_input(chat_input)

        clipboard_value = prompt.replace("\n", "\r\n")
        clipboard_ready = False
        clipboard_confirmed = False
        try:
            pyperclip.copy(clipboard_value)
            clipboard_ready = True

            def _clipboard_matches(expected: str, timeout_sec: float = 1.0) -> bool:
                deadline = time.monotonic() + timeout_sec
                while time.monotonic() < deadline:
                    try:
                        current = pyperclip.paste()
                    except Exception as paste_error:
                        print(f"警告: クリップボード内容の確認に失敗しました: {paste_error}")
                        return False
                    if current == expected:
                        return True
                    time.sleep(0.05)
                return False

            clipboard_confirmed = _clipboard_matches(clipboard_value)
            if not clipboard_confirmed:
                print("警告: クリップボードに期待した内容が設定されていません。")
        except Exception:
            print("警告: クリップボードへのコピーに失敗したためキーボード挿入にフォールバックします。")

        current_text = ""
        modifier = "Meta" if sys.platform == "darwin" else "Control"

        if clipboard_ready and clipboard_confirmed:
            clipboard_success = False
            for attempt in range(3):
                try:
                    chat_input.focus()
                except Exception:
                    pass
                pasted = False
                try:
                    chat_input.press(f"{modifier}+V")
                    pasted = True
                except Exception:
                    try:
                        self.page.keyboard.press(f"{modifier}+V")
                        pasted = True
                    except Exception as paste_error:
                        print(f"警告: クリップボード貼り付けに失敗しました: {paste_error}")
                if pasted:
                    try:
                        self.page.wait_for_timeout(200)
                    except Exception:
                        time.sleep(0.2)
                    current_text = self._read_chat_input_text(chat_input)
                    if current_text:
                        clipboard_success = True
                        break
                try:
                    pyperclip.copy(clipboard_value)
                    clipboard_confirmed = False
                    confirm_deadline = time.monotonic() + 0.75
                    while time.monotonic() < confirm_deadline:
                        try:
                            if pyperclip.paste() == clipboard_value:
                                clipboard_confirmed = True
                                break
                        except Exception as confirm_error:
                            print(f"警告: クリップボード再確認に失敗しました: {confirm_error}")
                            break
                        time.sleep(0.05)
                    if not clipboard_confirmed:
                        print("警告: クリップボードの再コピー内容を確認できませんでした。")
                        clipboard_ready = False
                        break
                except Exception as recopy_error:
                    print(f"警告: クリップボードの再コピーに失敗しました: {recopy_error}")
                    clipboard_ready = False
                    break
            if not clipboard_success and clipboard_ready:
                print("警告: クリップボード貼り付け結果が空だったため代替手段を試みます。")
        else:
            try:
                self.page.wait_for_timeout(200)
            except Exception:
                time.sleep(0.2)

        if not current_text:
            current_text = self._read_chat_input_text(chat_input)

        if not current_text:
            try:
                chat_input.click()
            except Exception:
                pass
            try:
                self.page.keyboard.insert_text(prompt)
            except Exception:
                try:
                    chat_input.fill(prompt)
                except Exception:
                    pass
            try:
                chat_input.focus()
            except Exception:
                pass
            try:
                self.page.wait_for_timeout(200)
            except Exception:
                time.sleep(0.2)
            current_text = self._read_chat_input_text(chat_input)

        if not current_text:
            try:
                injected = self.page.evaluate(
                    """
                    (target, value) => {
                        if (!target) {
                            return false;
                        }

                        const doc = target.ownerDocument || document;
                        const dispatchEvents = (node) => {
                            if (!node) {
                                return;
                            }
                            try {
                                const inputEvt = new InputEvent('input', {
                                    bubbles: true,
                                    data: value,
                                    inputType: 'insertFromPaste',
                                });
                                node.dispatchEvent(inputEvt);
                            } catch (err) {
                                /* no-op */
                            }
                            try {
                                const changeEvt = new Event('change', { bubbles: true });
                                node.dispatchEvent(changeEvt);
                            } catch (err) {
                                /* no-op */
                            }
                        };

                        const hydrateEditable = (node) => {
                            if (!node) {
                                return null;
                            }
                            const localDoc = node.ownerDocument || doc;
                            try {
                                if (node.focus) {
                                    node.focus();
                                }
                            } catch (err) {
                                /* focus best-effort */
                            }

                            if (typeof node.value === 'string') {
                                node.value = value;
                                return node;
                            }

                            let updated = false;
                            if (localDoc && typeof localDoc.execCommand === 'function') {
                                try {
                                    const selection = localDoc.getSelection && localDoc.getSelection();
                                    if (selection && selection.removeAllRanges) {
                                        selection.removeAllRanges();
                                        const range = localDoc.createRange();
                                        range.selectNodeContents(node);
                                        selection.addRange(range);
                                    }
                                    updated = localDoc.execCommand('insertText', false, value);
                                } catch (err) {
                                    updated = false;
                                }
                            }

                            if (!updated) {
                                try {
                                    node.innerHTML = '';
                                    const lines = value.split('\\n');
                                    lines.forEach((line) => {
                                        const paragraph = localDoc.createElement('p');
                                        if (line) {
                                            paragraph.textContent = line;
                                        } else {
                                            paragraph.appendChild(localDoc.createElement('br'));
                                        }
                                        node.appendChild(paragraph);
                                    });
                                    updated = true;
                                } catch (err) {
                                    updated = false;
                                }
                            }

                            return updated ? node : null;
                        };

                        const attempted = new Set();
                        const targets = [target];
                        const nested = target.querySelector('[contenteditable=\"true\"], textarea, [role=\"textbox\"]');
                        if (nested) {
                            targets.push(nested);
                        }

                        for (const candidate of targets) {
                            if (!candidate || attempted.has(candidate)) {
                                continue;
                            }
                            attempted.add(candidate);
                            const hydrated = hydrateEditable(candidate);
                            if (hydrated) {
                                dispatchEvents(hydrated);
                                if (hydrated !== target) {
                                    dispatchEvents(target);
                                }
                                return true;
                            }
                        }

                        return false;
                    }
                    """,
                    chat_input,
                    prompt,
                )
            except Exception:
                injected = False
            if injected:
                try:
                    self.page.wait_for_timeout(250)
                except Exception:
                    time.sleep(0.25)
                current_text = self._read_chat_input_text(chat_input)

        if not current_text:
            try:
                chat_input.focus()
            except Exception:
                pass
            typed_success = self._type_prompt_with_soft_returns(chat_input, prompt)
            if typed_success:
                try:
                    self.page.wait_for_timeout(200)
                except Exception:
                    time.sleep(0.2)
                deadline = time.monotonic() + 1.0
                while not current_text and time.monotonic() < deadline:
                    current_text = self._read_chat_input_text(chat_input)
                    if current_text:
                        break
                    time.sleep(0.1)
                if not current_text:
                    copied_result = ""
                    try:
                        chat_input.focus()
                    except Exception:
                        pass
                    try:
                        chat_input.press(f"{modifier}+A")
                    except Exception:
                        try:
                            if self.page:
                                self.page.keyboard.press(f"{modifier}+A")
                        except Exception:
                            pass
                    try:
                        chat_input.press(f"{modifier}+C")
                    except Exception:
                        try:
                            if self.page:
                                self.page.keyboard.press(f"{modifier}+C")
                        except Exception:
                            pass
                    try:
                        copied_result = pyperclip.paste()
                    except Exception:
                        copied_result = ""
                    if copied_result:
                        current_text = copied_result
                        try:
                            pyperclip.copy(clipboard_value)
                        except Exception:
                            pass
                if not current_text:
                    print(
                        "Warning: keyboard fallback reported success but reading the input failed; "
                        "continuing with the original prompt text."
                    )
                    current_text = prompt
            else:
                print("Warning: soft-return keyboard fallback was unable to populate the prompt.")

            if not current_text:
                current_text = self._read_chat_input_text(chat_input)

        if not current_text:
            raise RuntimeError("Failed to populate the chat input with the prompt.")

        return current_text

    def _submit_chat_input_via_keyboard(self, chat_input: Locator) -> bool:
        """Fallback submission when the explicit send button is unavailable."""
        if not self.page:
            return False

        modifier = "Meta" if sys.platform == "darwin" else "Control"
        key_sequences = [f"{modifier}+Enter", "Enter"]

        for sequence in key_sequences:
            for target in (chat_input, None):
                try:
                    if target is None:
                        self.page.keyboard.press(sequence)
                    else:
                        target.press(sequence)
                    print(f"キーボードショートカット '{sequence}' で送信を試みました。")
                    return True
                except Exception as press_error:
                    print(f"キーボード送信 '{sequence}' に失敗: {press_error}")
                    continue
        return False

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
            ("contenteditable div", lambda: self.page.locator("div[contenteditable]")),
            ("contenteditable section", lambda: self.page.locator("section[contenteditable]")),
            ("contenteditable span", lambda: self.page.locator("span[contenteditable]")),
            ("plaintext-only div", lambda: self.page.locator("div[contenteditable='plaintext-only']")),
            ("plaintext-only section", lambda: self.page.locator("section[contenteditable='plaintext-only']")),
            ("plaintext-only span", lambda: self.page.locator("span[contenteditable='plaintext-only']")),
            ("quill editor root", lambda: self.page.locator("div.ql-editor")),
            ("rich text container", lambda: self.page.locator("div[class*='rich-text'], div[class*='compose-area']")),
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
                (f"iframe#{idx + 1} contenteditable role textbox", lambda frame=frame: frame.locator('[contenteditable][role=\"textbox\"]')),
                (f"iframe#{idx + 1} contenteditable", lambda frame=frame: frame.locator('[contenteditable]')),
                (f"iframe#{idx + 1} plaintext-only", lambda frame=frame: frame.locator('[contenteditable=\"plaintext-only\"]')),
                (f"iframe#{idx + 1} textbox", lambda frame=frame: frame.locator('[role=\"textbox\"]')),
                (f"iframe#{idx + 1} textarea", lambda frame=frame: frame.locator('textarea')),
                (f"iframe#{idx + 1} paragraph", lambda frame=frame: frame.get_by_role('paragraph')),
                (f"iframe#{idx + 1} data-testid prompt", lambda frame=frame: frame.locator('[data-testid*=\"prompt\" i]')),
                (f"iframe#{idx + 1} rich text container", lambda frame=frame: frame.locator("div.ql-editor, div[class*='rich-text']")),
            ])

            for placeholder in placeholders:
                factories.append((f"iframe#{idx + 1} placeholder={placeholder}", lambda frame=frame, placeholder=placeholder: frame.get_by_placeholder(placeholder)))

        return factories

    def _resolve_chat_input_target(self, locator: Locator) -> Locator:
        """チャット欄を操作できる contenteditable コンテナを指すロケーターに正規化する"""
        try:
            enriched = locator.locator(
                "xpath=ancestor-or-self::*[@contenteditable and normalize-space(@contenteditable) != 'false'][1]"
            )
            if enriched.count() > 0:
                print("チャット入力欄: contenteditable な親要素にフォーカスを切り替えます。")
                return enriched.first
        except Exception as resolve_error:
            print(f"チャット入力欄: contenteditable 親要素の特定に失敗しました: {resolve_error}")

        descendant_selectors = [
            "[contenteditable='true']",
            "[contenteditable]:not([contenteditable='false'])",
            "div[contenteditable]",
            "section[contenteditable]",
            "span[contenteditable]",
            "textarea",
            "[role='textbox']",
            "[data-content-editable-root='true']",
            "[data-slate-editor='true']",
        ]

        for selector in descendant_selectors:
            try:
                descendant = locator.locator(selector)
                if descendant.count() > 0:
                    print(f"チャット入力欄: {selector} の子孫要素に切り替えます。")
                    return descendant.first
            except Exception as descendant_error:
                print(f"チャット入力欄: {selector} の子孫要素探索に失敗しました: {descendant_error}")

        return locator

    def _initialize_copilot_mode(self):
        """GPT-5 チャットモードを有効化し、入力欄を準備する"""
        try:
            copilot_url_patterns = [
                re.compile(r"^https://m365\.cloud\.microsoft/(?:chat|copilot)(?:/|\?|$)", re.IGNORECASE),
                re.compile(r"^https://copilot\.microsoft\.com/(?:chat)?(?:/|\?|$)", re.IGNORECASE),
                re.compile(r"^https://www\.office\.com/launch/copilot", re.IGNORECASE),
            ]

            def _matches_copilot_url(url: str) -> bool:
                if not url:
                    return False
                sanitized = url.split('#', 1)[0]
                return any(pattern.search(sanitized) for pattern in copilot_url_patterns)

            current_url = self.page.url if self.page else ""
            if _matches_copilot_url(current_url):
                print("Copilot ページのURLを確認しました。")
            else:
                print("Copilot ページが表示されるのを待機しています...")
                try:
                    self.page.wait_for_url(lambda url: _matches_copilot_url(url), timeout=120000)
                    current_url = self.page.url if self.page else current_url
                    print("Copilot ページのロードを確認しました。")
                except PlaywrightTimeoutError as wait_error:
                    current_url = self.page.url if self.page else current_url
                    print(f"警告: Copilot ページのURLが既定のパターンに到達しません (現在: {current_url})。引き続き初期化を試みます。詳細: {wait_error}")
            self.page.wait_for_load_state('domcontentloaded')
            self.page.wait_for_timeout(500)

            gpt5_button_factories = [
                ("button.fui-ToggleButton:has-text('GPT-5 を試す')", lambda: self.page.locator("button.fui-ToggleButton:has-text('GPT-5 を試す')")),
                ("button role GPT-5 を試す", lambda: self.page.get_by_role("button", name="GPT-5 を試す")),
                ("button role GPT-5 を試す (部分一致)", lambda: self.page.get_by_role("button", name="GPT-5 を試す", exact=False)),
                ("button role GPT-5 で質問", lambda: self.page.get_by_role("button", name="GPT-5 で質問")),
                ("button role GPT-5 を使用", lambda: self.page.get_by_role("button", name="GPT-5 を使用")),
                ("button role GPT-5", lambda: self.page.get_by_role("button", name="GPT-5")),
                ("button role Copilot GPT-5", lambda: self.page.get_by_role("button", name="Copilot GPT-5")),
                ("button role regex GPT-5", lambda: self.page.get_by_role("button", name=re.compile(r"GPT[-\s]?5", re.IGNORECASE))),
                ("menuitem regex GPT-5", lambda: self.page.get_by_role("menuitem", name=re.compile(r"GPT[-\s]?5", re.IGNORECASE))),
                ("button:has-text GPT-5", lambda: self.page.locator("button", has_text=re.compile(r"GPT[-\s]?5", re.IGNORECASE))),
                ("aria-label=GPT-5 を試す", lambda: self.page.locator('[aria-label="GPT-5 を試す" i]')),
                ("aria-label contains GPT-5", lambda: self.page.locator('[aria-label*="GPT-5"]')),
                ("data-testid=gpt5Button", lambda: self.page.locator('[data-testid="gpt5Button"]')),
                ("data-testid=GPT5Button", lambda: self.page.locator('[data-testid="GPT5Button"]')),
                ("text=GPT-5 を試す", lambda: self.page.get_by_text("GPT-5 を試す", exact=False)),
                ("text=Try GPT-5", lambda: self.page.get_by_text("Try GPT-5", exact=False)),
            ]

            gpt5_button = None
            try:
                gpt5_button = self._wait_for_first_visible(
                    "GPT-5 モード切り替えボタン",
                    gpt5_button_factories,
                    timeout=20000,
                )
            except RuntimeError as gpt5_error:
                print("GPT-5 モードボタンが見つからなかったため、フォールバック探索を行います。")
                print(gpt5_error)
                fallback_factories = [
                    ("role=button GPT-5 を試す (再探索)", lambda: self.page.get_by_role("button", name="GPT-5 を試す", exact=False)),
                    ("text=GPT-5 を試す", lambda: self.page.get_by_text("GPT-5 を試す", exact=False)),
                    ("text=Try GPT-5", lambda: self.page.get_by_text("Try GPT-5", exact=False)),
                ]
                for description, factory in fallback_factories:
                    try:
                        candidate = factory()
                        candidate.wait_for(state="visible", timeout=5000)
                        gpt5_button = candidate.first
                        print(f"GPT-5 モードボタンのフォールバック ({description}) に成功しました。")
                        break
                    except Exception as fallback_error:
                        print(f"GPT-5 モードボタンのフォールバック ({description}) は失敗しました: {fallback_error}")
            if gpt5_button:
                print("GPT-5 モードボタンをクリックします...")
                try:
                    gpt5_button.scroll_into_view_if_needed()
                except Exception:
                    pass
                click_attempts = [
                    ("通常クリック", lambda: gpt5_button.click(timeout=5000)),
                    ("force オプション", lambda: gpt5_button.click(force=True, timeout=5000)),
                    ("evaluate click", lambda: gpt5_button.evaluate("el => el.click()")),
                ]
                for attempt_label, click_action in click_attempts:
                    try:
                        click_action()
                        self.page.wait_for_timeout(800)
                        break
                    except Exception as click_error:
                        print(f"GPT-5 モードボタンのクリック ({attempt_label}) に失敗しました: {click_error}")
                else:
                    print("GPT-5 モードボタンのクリックに失敗しました。直近のエラーをご確認ください。")
            else:
                print("GPT-5 モードボタンが見つからなかったため、既定のモードで続行します。")

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
                        timeout=45000,
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


    def ask(self, prompt: str, stop_event: Optional[Event] = None) -> str:
        """プロンプトを送信し、Copilotからの応答をクリップボード経由で取得する"""
        if not self.page:
            raise ConnectionError("ブラウザが初期化されていません。start()を呼び出してください。")

        def _ensure_not_stopped():
            if stop_event and stop_event.is_set():
                print("Stop requested: attempting to cancel Copilot response.")
                try:
                    self.request_stop()
                except Exception:
                    pass
                raise UserStopRequested("ユーザーによる中断が要求されました。")

        try:
            _ensure_not_stopped()
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

            _ensure_not_stopped()
            try:
                chat_input.click(timeout=5000)
            except TypeError:
                chat_input.click()
            except Exception as click_error:
                print(f"チャット入力欄でのフォーカス時に警告: {click_error}")

            _ensure_not_stopped()
            filled_text = self._fill_chat_input(chat_input, prompt)

            expected_text = self._normalize_prompt_text(prompt)
            normalized_current = self._normalize_prompt_text(filled_text)

            if normalized_current != expected_text:
                print(
                    "警告: チャット入力欄の内容が期待するプロンプトと一致しません。"
                    f" expected_len={len(expected_text)} actual_len={len(normalized_current)}"
                )
                mismatch_deadline = time.monotonic() + 2.0
                while time.monotonic() < mismatch_deadline:
                    current_value = self._normalize_prompt_text(self._read_chat_input_text(chat_input))
                    if current_value == expected_text:
                        print("チャット入力欄: 遅延後にプロンプト全体を検知できました。")
                        normalized_current = current_value
                        break
                    time.sleep(0.1)

                retry_count = 0
                max_retries = 2
                while normalized_current != expected_text and retry_count < max_retries:
                    retry_count += 1
                    print(f"チャット入力欄: 再入力リトライ {retry_count}/{max_retries} を実行します。")
                    try:
                        self._clear_chat_input(chat_input)
                    except Exception as clear_error:
                        print(f"警告: チャット入力欄のクリアに失敗しました: {clear_error}")
                    try:
                        self.page.wait_for_timeout(150)
                    except Exception:
                        time.sleep(0.15)
                    filled_text = self._fill_chat_input(chat_input, prompt)
                    normalized_current = self._normalize_prompt_text(filled_text)

                    if normalized_current != expected_text:
                        print(
                            "警告: 再入力後もチャット入力欄の内容が一致しません。"
                            f" expected_len={len(expected_text)} actual_len={len(normalized_current)}"
                        )

                if normalized_current != expected_text:
                    raise RuntimeError("チャット入力欄にプロンプト全体を反映できませんでした。")

            # Allow the pasted content to settle before sending
            try:
                self.page.wait_for_timeout(350)
            except Exception:
                time.sleep(0.35)

            _ensure_not_stopped()
            try:
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
                        ("data-testid^=send", lambda: self.page.locator('[data-testid^="send" i]')),
                        ("data-testid^=chat-send", lambda: self.page.locator('[data-testid^="chat-send" i]')),
                        ("role=button name~=Send", lambda: self.page.get_by_role("button", name=re.compile("send", re.IGNORECASE))),
                        ("aria-label~=送信", lambda: self.page.locator('[aria-label*="送信" i]')),
                        ("type=submit", lambda: self.page.locator('button[type="submit"]')),
                    ],
                    timeout=15000,
                )
                _ensure_not_stopped()
                send_button.click()
            except Exception as send_error:
                print(f"送信ボタンをクリックできませんでした: {send_error}. キーボード送信を試みます。")
                _ensure_not_stopped()
                if not self._submit_chat_input_via_keyboard(chat_input):
                    raise RuntimeError("送信ボタンが見つからず、キーボード送信にも失敗しました。") from send_error

            _ensure_not_stopped()
            # 新しいコピーボタン（＝ユーザー入力・応答）が出現するのを待つ
            print("Copilotの応答を待っています...")
            new_copy_button_locator = self.page.locator(copy_button_selector).nth(initial_copy_button_count)
            deadline = time.monotonic() + 180
            while True:
                try:
                    new_copy_button_locator.wait_for(state="visible", timeout=1000)
                    break
                except PlaywrightTimeoutError:
                    _ensure_not_stopped()
                    if time.monotonic() >= deadline:
                        raise
                    continue

            def _looks_like_prompt_echo(captured: str) -> bool:
                sanitized = (captured or "").strip()
                if not sanitized:
                    return True
                lowered = sanitized.lower()
                prompt_sample = (prompt or "").strip().lower()[:40]
                if prompt_sample and lowered.startswith(prompt_sample):
                    return True
                if lowered.startswith("system:") or lowered.startswith("user:"):
                    return True
                if lowered.startswith("[translation mode request]") or lowered.startswith("[translation review mode request]"):
                    return True
                if lowered.endswith("assistant:"):
                    return True
                if prompt and sanitized == prompt.strip():
                    return True
                return False

            response_text = None
            while True:
                _ensure_not_stopped()
                copy_buttons = self.page.locator(copy_button_selector)
                total_buttons = copy_buttons.count()
                if total_buttons <= initial_copy_button_count:
                    if time.monotonic() >= deadline:
                        return "エラー: 新しい応答のコピーボタンが見つかりませんでした。"
                    time.sleep(0.4)
                    continue

                for index in range(total_buttons - 1, initial_copy_button_count - 1, -1):
                    candidate_button = copy_buttons.nth(index)
                    try:
                        candidate_button.wait_for(state="visible", timeout=1000)
                    except PlaywrightTimeoutError:
                        continue
                    except Exception:
                        continue

                    try:
                        candidate_button.click()
                    except Exception:
                        continue

                    time.sleep(0.5)
                    clipboard_text = pyperclip.paste().strip()
                    if _looks_like_prompt_echo(clipboard_text):
                        continue

                    response_text = clipboard_text
                    break

                if response_text is not None:
                    break

                if time.monotonic() >= deadline:
                    return "エラー: Copilotからの応答を取得できませんでした。"
                time.sleep(0.5)

            print("応答が完了したと判断しました。")
            if response_text is None:
                return "エラー: Copilotからの応答を取得できませんでした。"

            thought_pos = response_text.find("Thought:")
            if thought_pos != -1:
                return response_text[thought_pos:]
            return response_text

        except PlaywrightTimeoutError:
            return "エラー: Copilotからの応答がタイムアウトしました。応答に時間がかかりすぎるか、Copilotが応答不能になっている可能性があります。"
        except UserStopRequested:
            raise
        except Exception as e:
            return f"エラー: ブラウザ操作中に予期せぬエラーが発生しました: {e}"

    def request_stop(self) -> bool:
        if not self.page:
            return False

        stop_candidates = [
            lambda: self.page.get_by_role("button", name="停止"),
            lambda: self.page.get_by_role("button", name="中止"),
            lambda: self.page.get_by_role("button", name="中断"),
            lambda: self.page.get_by_role("button", name=re.compile("Stop", re.IGNORECASE)),
            lambda: self.page.get_by_role("button", name=re.compile("Cancel", re.IGNORECASE)),
            lambda: self.page.get_by_test_id("StopGenerating"),
        ]

        for factory in stop_candidates:
            try:
                locator = factory()
                if locator.count() == 0:
                    continue
                locator.first.click()
                print("Copilot応答の停止ボタンをクリックしました。")
                return True
            except Exception:
                continue

        try:
            self.page.keyboard.press("Escape")
            print("Copilot応答の停止を Escape キーで試みました。")
            return True
        except Exception:
            pass

        return False


    def restart(self):
        """ブラウザセッションを終了し、起動時と同じ状態に再初期化する"""

        self._logger.info("ブラウザセッションを再初期化します。")
        self.close()
        self.start()


    def close(self):
        """ブラウザとPlaywrightセッションを安全に閉じる"""
        if self.context:
            try:
                self.context.close()
                self._logger.info("ブラウザコンテキストを閉じました。")
            except Exception as e:
                self._logger.warning("ブラウザコンテキストを閉じる際にエラーが発生しました: %s", e)
        if self.playwright:
            try:
                self.playwright.stop()
                self._logger.info("Playwrightを停止しました。")
            except Exception as e:
                self._logger.warning("Playwrightを停止する際にエラーが発生しました: %s", e)
        self.page = None
        self.context = None
        self.playwright = None
        self._focus_suppressed_once = False

    def _ensure_browser_visible(self):
        if self.headless or not self.page:
            return

        try:
            self.page.bring_to_front()
            self._logger.debug("ブラウザウィンドウを前面に表示しました。")
        except Exception as exc:
            self._logger.debug("ブラウザウィンドウを前面に表示できませんでした: %s", exc)

    def _suppress_browser_focus(self):
        if (
            not COPILOT_SUPPRESS_BROWSER_FOCUS
            or self.headless
            or not self.context
            or not self.page
            or self._focus_suppressed_once
        ):
            return

        try:
            session = self.context.new_cdp_session(self.page)
            window_info = session.send("Browser.getWindowForTarget")
            window_id = window_info.get("windowId")
            if not window_id:
                return
            current_bounds = window_info.get("bounds", {})
            new_bounds = {
                "left": COPILOT_SUPPRESS_BROWSER_FOCUS_LEFT,
                "top": COPILOT_SUPPRESS_BROWSER_FOCUS_TOP,
            }
            if "width" in current_bounds:
                new_bounds["width"] = current_bounds["width"]
            if "height" in current_bounds:
                new_bounds["height"] = current_bounds["height"]

            session.send(
                "Browser.setWindowBounds",
                {"windowId": window_id, "bounds": new_bounds},
            )
            self._focus_suppressed_once = True
            self._logger.debug(
                "ブラウザウィンドウをバックグラウンド位置へ移動しました (left=%s, top=%s)。",
                COPILOT_SUPPRESS_BROWSER_FOCUS_LEFT,
                COPILOT_SUPPRESS_BROWSER_FOCUS_TOP,
            )
        except Exception as exc:
            self._logger.debug("ブラウザウィンドウ位置の調整に失敗しました: %s", exc)
