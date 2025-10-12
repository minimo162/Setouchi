"""
End-to-end regression for the Excel Copilot desktop app.

This test drives the Flet UI with Playwright in order to cover a real
translation (reference mode) scenario.  Because it relies on an actual Excel
workbook, SharePoint-accessible references, and a logged-in Copilot session,
the test is **opt-in**: set the environment variable `PLAYWRIGHT_E2E=1` on a
machine where `python desktop_app.py` is running before executing pytest.

Example:

    PLAYWRIGHT_E2E=1 FLET_APP_URL=http://127.0.0.1:63644 pytest tests/e2e/test_desktop_app_e2e.py

The test assumes the workbook "test.xlsx" is open in Excel (Sheet "Sheet") with
the source text in cell A1, matching the instructions in the repository README.
"""

from __future__ import annotations

import os
import time
from typing import Optional

import pytest
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from excel_copilot.config import COPILOT_USER_DATA_DIR

FLET_APP_URL = os.getenv("FLET_APP_URL", "http://127.0.0.1:63644")
PROMPT_TEXT = (
    "A1セルを日本語参照: "
    "https://ralleti.sharepoint.com/:b:/s/test/ES7NUzlbE29Ng1c-8smAb-wBBUmzQGNtqenf0XcOoi5xeg?e=9IxUSS "
    "英語参照: "
    "https://ralleti.sharepoint.com/:b:/s/test/EZ5EWz_WX4tAvKZNY_unLSgBaWRXIZJxIvf2lOnN8xHXxg?e=oMnLBU "
    "で英訳し、B列以降に出力してください。"
)


def _launch_context():
    """Launch a Playwright persistent context using the Copilot profile."""
    user_data_dir = os.getenv("PLAYWRIGHT_USER_DATA_DIR", COPILOT_USER_DATA_DIR)
    headless = os.getenv("PLAYWRIGHT_HEADLESS", "0") in {"1", "true", "TRUE", "yes"}
    playwright = sync_playwright().start()
    browser = None
    errors: list[Exception] = []
    for channel in (os.getenv("PLAYWRIGHT_CHANNEL"), "msedge", "chrome", None):
        launch_args = {"headless": headless, "slow_mo": 50}
        if channel:
            launch_args["channel"] = channel
        try:
            browser = playwright.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                **launch_args,
            )
            break
        except Exception as exc:  # pragma: no cover - best effort for E2E
            errors.append(exc)
            continue
    if browser is None:
        playwright.stop()
        raise RuntimeError(
            "Failed to launch persistent browser context. "
            "Ensure Edge or Chrome is installed and that the Copilot profile directory is accessible."
        ) from (errors[-1] if errors else None)
    return playwright, browser


@pytest.mark.skipif(
    os.getenv("PLAYWRIGHT_E2E") not in {"1", "true", "TRUE", "yes"},
    reason="Set PLAYWRIGHT_E2E=1 to enable the desktop E2E test.",
)
def test_reference_translation_flow():
    playwright, context = _launch_context()
    page = context.pages[0] if context.pages else context.new_page()

    try:
        page.goto(FLET_APP_URL, wait_until="domcontentloaded", timeout=30_000)
    except PlaywrightTimeoutError as exc:  # pragma: no cover - environment dependent
        context.close()
        playwright.stop()
        raise AssertionError(
            f"Failed to reach Flet app at {FLET_APP_URL}. "
            "Ensure `python desktop_app.py` is running with `--no-browser` or view enabled."
        ) from exc

    # Wait for initialization to finish (worker signals "初期化が完了しました。指示をどうぞ。")
    page.wait_for_timeout(1_000)
    initialization_selector = "text=初期化が完了しました"
    try:
        page.wait_for_selector(initialization_selector, timeout=120_000)
    except PlaywrightTimeoutError as exc:  # pragma: no cover - environment dependent
        context.close()
        playwright.stop()
        raise AssertionError("Desktop app did not finish initialization in time.") from exc

    # Switch to 翻訳（参照あり） if necessary.
    try:
        reference_tab = page.get_by_text("翻訳（参照あり）")
        if reference_tab.is_visible():
            reference_tab.click()
    except Exception:
        pass  # already selected or control not visible

    # Select workbook and sheet dropdowns if available.
    def _select_dropdown(text_to_select: str, index: int = 0) -> None:
        try:
            combobox = page.locator("select").nth(index)
            if combobox.count():
                combobox.select_option(label=text_to_select)
        except Exception:
            pass

    workbook_name = os.getenv("E2E_WORKBOOK", "test")
    sheet_name = os.getenv("E2E_SHEET", "Sheet")
    _select_dropdown(workbook_name, index=0)
    _select_dropdown(sheet_name, index=1)

    # Type prompt into the chat textbox (textarea is used for multiline input).
    textarea = page.locator("textarea")
    if not textarea.count():
        context.close()
        playwright.stop()
        raise AssertionError("Chat input textarea was not found on the page.")
    textarea.fill(PROMPT_TEXT)

    # Click the send button (tooltip 送信).
    send_button = page.get_by_role("button", name="送信")
    if not send_button.is_visible():
        # Some builds render the IconButton differently; fall back to first matching FAB.
        send_button = page.locator("button").first
    send_button.click()

    # Wait for translation progress cues and completion.
    progress_markers = [
        "キーフレーズ生成: Copilotに依頼中",
        "日本語参照文章抽出: Copilotに依頼中",
        "対になる英語参照文抽出: Copilotに依頼中",
        "Row",
    ]

    deadline = time.monotonic() + 300  # up to 5 minutes for slow runs
    for marker in progress_markers:
        remaining = max(0, int((deadline - time.monotonic()) * 1000))
        if remaining <= 0:
            raise AssertionError(f"Timed out waiting for progress marker: {marker}")
        page.wait_for_selector(f"text={marker}", timeout=remaining)

    # Final check: ensure no bibliography text leaked into the output panel.
    bibliography = page.locator("text=書誌")
    assert bibliography.count() == 0, "Bibliography text should not appear in the UI output."

    context.close()
    playwright.stop()
