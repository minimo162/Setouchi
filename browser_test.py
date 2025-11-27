#!/usr/bin/env python3
"""Test Flet UI rendering with headless browser."""
import asyncio
import subprocess
import time
import sys

async def test_flet_ui():
    from playwright.async_api import async_playwright

    # Start Flet server
    print("Starting Flet test server...")
    server = subprocess.Popen(
        [sys.executable, "test_ui.py"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd="/home/user/Setouchi"
    )

    # Wait for server to start
    time.sleep(3)

    try:
        async with async_playwright() as p:
            print("Launching browser...")
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page()

            print("Navigating to Flet app...")
            await page.goto("http://127.0.0.1:8551", timeout=10000)

            # Wait for Flet to load
            await page.wait_for_timeout(3000)

            # Check page content
            content = await page.content()
            print(f"Page content length: {len(content)}")

            # Take screenshot
            await page.screenshot(path="/home/user/Setouchi/screenshot.png")
            print("Screenshot saved to screenshot.png")

            # Check for specific text
            body_text = await page.inner_text("body")
            print(f"Body text (first 500 chars): {body_text[:500] if body_text else 'EMPTY'}")

            # Check for debug text
            if "DEBUG" in body_text:
                print("SUCCESS: Debug text found!")
            else:
                print("WARNING: Debug text NOT found")

            await browser.close()
    finally:
        server.terminate()
        server.wait()

if __name__ == "__main__":
    asyncio.run(test_flet_ui())
