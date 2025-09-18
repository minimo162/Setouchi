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
from typing import Optional, Callable, List

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

    def _wait_for_first_visible(self, description: str, locator_factories: List[Callable[[], Locator]], timeout: float) -> Locator:
        """与えられたロケーター群の中から、最初に可視化された要素を返す"""
        last_exception: Optional[Exception] = None
        for factory in locator_factories:
            try:
                locator = factory()
                # first() は常に存在する前提のため、先にcountで存在確認
                if locator.count() == 0:
                    continue
                first_visible = locator.first
                first_visible.wait_for(state="visible", timeout=timeout)
                return first_visible
            except PlaywrightTimeoutError as e:
                last_exception = e
            except Exception as e:  # 予期せぬエラーも記録し、次の候補を試す
                last_exception = e
        raise RuntimeError(f"{description}が見つかりません。UI が変更された可能性があります。") from last_exception

    def _fill_chat_input(self, chat_input: Locator, prompt: str):
        """入力欄の種類に応じて、テキストを確実に入力する"""
        try:
            chat_input.fill(prompt)
            # contenteditable要素ではinner_textの変化を確認
            current_text = chat_input.inner_text().strip()
            if current_text:
                return
        except Exception:
            pass

        chat_input.click()
        modifier = "Meta" if sys.platform == "darwin" else "Control"
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

        # contenteditable な場合に備えて JavaScript で直接値を書き込む
        injected = False
        try:
            self.page.evaluate(
                """
                (target, value) => {
                    if (!target) return false;
                    target.innerHTML = '';
                    const paragraph = document.createElement('p');
                    paragraph.textContent = value;
                    target.appendChild(paragraph);
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
            injected = True
        except Exception:
            injected = False

        if not injected:
            try:
                chat_input.type(prompt, delay=20)
            except Exception:
                try:
                    self.page.keyboard.type(prompt, delay=20)
                except Exception:
                    pass

        self.page.wait_for_timeout(120)
        try:
            current_text = chat_input.inner_text().strip()
        except Exception:
            current_text = ""

        if not current_text:
            try:
                self.page.keyboard.insert_text(prompt)
            except Exception:
                pass
            self.page.wait_for_timeout(120)
            try:
                current_text = chat_input.inner_text().strip()
            except Exception:
                current_text = ""

        if not current_text:
            raise RuntimeError("チャット入力欄へのテキスト入力に失敗しました。")

    def _initialize_copilot_mode(self):
        """GPT-5 チャットモードを有効化し、入力欄を準備する"""
        try:
            gpt5_locator = self.page.get_by_role("button", name="GPT-5 を試す")
            if gpt5_locator.count() > 0:
                try:
                    print("「GPT-5 を試す」ボタンを待機中...")
                    gpt5_locator.first.wait_for(state="visible", timeout=15000)
                    print("「GPT-5 を試す」ボタンをクリックします...")
                    gpt5_locator.first.click()
                except PlaywrightTimeoutError:
                    print("「GPT-5 を試す」ボタンの表示待機がタイムアウトしました。既に GPT-5 が有効か UI が変更されています。")
                except Exception as click_error:
                    print(f"「GPT-5 を試す」ボタンのクリック中に問題が発生しました: {click_error}。既に GPT-5 が選択済みの可能性があります。")
            else:
                print("「GPT-5 を試す」ボタンが見つかりませんでした。既に GPT-5 が選択されている可能性があります。")

            # チャット入力欄が表示され、編集可能になるのを待つ
            chat_input = self._wait_for_first_visible(
                "チャット入力欄",
                [
                    lambda: self.page.locator('[contenteditable="true"][role="textbox"]'),
                    lambda: self.page.get_by_role("paragraph"),
                    lambda: self.page.get_by_role("textbox", name="チャット入力"),
                ],
                timeout=45000,
            )

            print("チャット入力欄をフォーカスします...")
            chat_input.click()

            print("準備完了。GPT-5 での自動操作を開始できます。")

        except PlaywrightTimeoutError as e:
            print(f"エラー: チャット入力欄の初期化がタイムアウトしました。UIの構造が変更された可能性があります。: {e}")
            raise
        except Exception as e:
            print(f"エラー: GPT-5 への切り替え中に予期せぬエラーが発生しました: {e}")
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
                [
                    lambda: self.page.locator('[contenteditable="true"][role="textbox"]'),
                    lambda: self.page.get_by_role("paragraph"),
                    lambda: self.page.get_by_role("textbox", name="チャット入力"),
                ],
                timeout=45000,
            )

            chat_input.click()  # フォーカスを当てる
            self._fill_chat_input(chat_input, prompt)
            send_button = self._wait_for_first_visible(
                "送信ボタン",
                [
                    lambda: self.page.get_by_label("送信"),
                    lambda: self.page.get_by_label("送信 (Ctrl+Enter)"),
                    lambda: self.page.get_by_role("button", name="送信"),
                    lambda: self.page.get_by_role("button", name="送信 (Ctrl+Enter)"),
                    lambda: self.page.get_by_role("button", name="Send"),
                    lambda: self.page.get_by_label("Send"),
                    lambda: self.page.get_by_test_id("SendButtonTestId"),
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
