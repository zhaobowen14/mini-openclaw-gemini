from __future__ import annotations

import argparse
import asyncio
import json
import subprocess
import time
from datetime import datetime
from pathlib import Path

from playwright.async_api import Browser, BrowserContext, Page
from playwright.async_api import TimeoutError as PlaywrightTimeoutError
from playwright.async_api import async_playwright

URL = "https://gemini.google.com/app"
DEFAULT_PROMPT = "テスト"
CHROME_PATH = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
DEBUG_PORT = 9223
USER_DATA_DIR = Path(r"C:\Temp\ChromeProfile9223")

UNAVAILABLE_PATTERNS = [
    "502. That's an error.",
    "The server encountered a temporary error",
    "現在ご利用いただけません",
    "ご利用いただけません",
    "This app isn't available",
    "isn't available in your country",
    "not available",
]

INPUT_SELECTORS = [
    'rich-textarea .ql-editor[contenteditable="true"]',
    'rich-textarea div[contenteditable="true"][role="textbox"]',
    '.ql-editor[contenteditable="true"][role="textbox"]',
    'div[contenteditable="true"][role="textbox"]',
    'div[contenteditable="true"]',
    'textarea',
    '[role="textbox"]',
]

SEND_SELECTORS = [
    'button.send-button.submit.has-input',
    'button.send-button.submit',
    'gem-icon-button.send-button.submit.has-input',
    'gem-icon-button.send-button.submit',
    'button[aria-label*="送信"]',
    'button[aria-label*="Send"]',
    'button:has-text("送信")',
    'button:has-text("Send")',
]

ANSWER_SELECTORS = [
    'model-response message-content',
    'message-content.model-response-text',
    '.model-response-text',
    '[data-response-index]',
    'message-content',
    '[role="article"]',
]

SKILL_DIR = Path(__file__).resolve().parent.parent
OUTPUTS_DIR = SKILL_DIR / "outputs"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gemini with Playwright")
    parser.add_argument("prompt", nargs="?", default=DEFAULT_PROMPT, help="単発で送るプロンプト")
    parser.add_argument("--prompts-file", help="複数ターン用のテキストファイル。1行=1プロンプト")
    parser.add_argument("--output", default="", help="回答保存先。相対パスは skills/gemini-playwright-chat/outputs/ 配下に保存")
    parser.add_argument("--headful", action="store_true", help="ブラウザを画面表示で起動する")
    parser.add_argument("--delay", type=int, default=5, help="キーボード入力の1文字あたり遅延(ms)")
    return parser.parse_args()


def load_prompts(args: argparse.Namespace) -> list[str]:
    if args.prompts_file:
        path = Path(args.prompts_file)
        lines = [line.strip() for line in path.read_text(encoding="utf-8").splitlines()]
        prompts = [line for line in lines if line]
        if not prompts:
            raise ValueError("prompts file に有効な行がありません。")
        return prompts
    return [args.prompt]


async def first_visible(page: Page, selectors, timeout_ms: int = 20000):
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while asyncio.get_running_loop().time() < deadline:
        for selector in selectors:
            locator = page.locator(selector).first
            try:
                if await locator.count() > 0 and await locator.is_visible():
                    return locator
            except Exception:
                pass
        await asyncio.sleep(0.1)
    raise TimeoutError(f"Visible element not found. Selectors={selectors}")



async def page_text(page: Page) -> str:
    try:
        return await page.locator("body").inner_text()
    except Exception:
        return ""




async def extract_answer(page: Page, previous_last_text: str = "", timeout_ms: int = 90000) -> str:
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    last_text = ""
    stable_count = 0

    while asyncio.get_running_loop().time() < deadline:


        for selector in ANSWER_SELECTORS:
            locator = page.locator(selector).last
            try:
                if await locator.count() == 0 or not await locator.is_visible():
                    continue
                text = (await locator.inner_text()).strip()
                if not text or text == previous_last_text:
                    continue
                if text == last_text:
                    stable_count += 1
                else:
                    last_text = text
                    stable_count = 0
                if stable_count >= 10 and len(text) >= 2:
                    return text
            except Exception:
                pass
        await asyncio.sleep(0.5)

    if last_text:
        return last_text
    raise TimeoutError("Gemini の回答取得に失敗しました。")


async def connect_over_cdp(playwright) -> tuple[Browser, bool]:
    cdp_url = f"http://127.0.0.1:{DEBUG_PORT}"
    try:
        browser = await playwright.chromium.connect_over_cdp(cdp_url)
        return browser, False
    except Exception:
        pass

    if not CHROME_PATH.exists():
        raise RuntimeError(f"Chrome が見つかりません: {CHROME_PATH}")

    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(
        [
            str(CHROME_PATH),
            f"--remote-debugging-port={DEBUG_PORT}",
            f"--user-data-dir={USER_DATA_DIR}",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=creationflags,
    )

    deadline = time.time() + 20
    last_error: Exception | None = None
    while time.time() < deadline:
        try:
            browser = await playwright.chromium.connect_over_cdp(cdp_url)
            return browser, True
        except Exception as e:
            last_error = e
            await asyncio.sleep(0.5)
    raise RuntimeError(f"CDP 接続に失敗しました: {last_error}")


async def get_or_create_context(browser: Browser) -> BrowserContext:
    if browser.contexts:
        return browser.contexts[0]
    return await browser.new_context(locale="ja-JP")


async def get_or_create_gemini_page(context: BrowserContext) -> Page:
    for page in context.pages:
        try:
            if "gemini.google.com" in page.url:
                await page.bring_to_front()
                return page
        except Exception:
            pass

    page = await context.new_page()
    await page.goto(URL, wait_until="domcontentloaded")
    return page


async def send_prompt(page: Page, prompt: str, delay_ms: int, previous_answer: str = "") -> str:
    input_box = await first_visible(page, INPUT_SELECTORS, timeout_ms=30000)
    await input_box.click()
    await page.keyboard.press("Control+A")
    await input_box.press("Space")
    await page.evaluate("text => navigator.clipboard.writeText(text)", prompt)
    await page.keyboard.press("Control+v")
    await page.wait_for_timeout(100)

    send_clicked = False
    for selector in SEND_SELECTORS:
        button = page.locator(selector).first
        try:
            if await button.count() > 0 and await button.is_visible():
                aria_disabled = await button.get_attribute("aria-disabled")
                disabled = await button.get_attribute("disabled")
                if aria_disabled == "true" or disabled is not None:
                    continue
                await button.click()
                send_clicked = True
                break
        except Exception:
            pass

    if not send_clicked:
        await page.keyboard.press("Enter")

    return await extract_answer(page, previous_last_text=previous_answer, timeout_ms=90000)


def safe_print(text: str = "") -> None:
    if text is None:
        text = ""
    try:
        print(text)
    except UnicodeEncodeError:
        safe = text.encode("cp932", errors="replace").decode("cp932", errors="replace")
        print(safe)


def save_output(path_str: str, results: list[dict]) -> None:
    if not path_str:
        return

    raw_path = Path(path_str)
    path = raw_path if raw_path.is_absolute() else OUTPUTS_DIR / raw_path
    path.parent.mkdir(parents=True, exist_ok=True)

    if path.suffix.lower() == ".json":
        path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        return

    lines: list[str] = []
    for item in results:
        lines.append(f"[Turn {item['turn']}]")
        lines.append(f"Prompt: {item['prompt']}")
        lines.append("Answer:")
        lines.append(item["answer"])
        lines.append("")
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


async def main() -> None:
    args = parse_args()
    prompts = load_prompts(args)
    results: list[dict] = []

    async with async_playwright() as p:
        browser, started_chrome = await connect_over_cdp(p)
        owned_context: BrowserContext | None = None
        try:
            context = await get_or_create_context(browser)
            if not browser.contexts:
                owned_context = context
            page = await get_or_create_gemini_page(context)
            page.set_default_timeout(10000)

            await page.goto(URL, wait_until="domcontentloaded")
            await page.wait_for_timeout(100)


            previous_answer = ""
            for index, prompt in enumerate(prompts, start=1):
                answer = await send_prompt(page, prompt, args.delay, previous_answer=previous_answer)
                previous_answer = answer
                item = {
                    "turn": index,
                    "prompt": prompt,
                    "answer": answer,
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                }
                results.append(item)

                # safe_print(f"=== Turn {index} ===")
                # safe_print(f"Prompt: {prompt}")
                # safe_print("Answer:")
                # safe_print(answer)
                # safe_print()

                await page.wait_for_timeout(1500)

            save_output(args.output, results)
            if args.output:
                raw_path = Path(args.output)
                saved_path = raw_path if raw_path.is_absolute() else OUTPUTS_DIR / raw_path
                safe_print(f"保存先: {saved_path}")

        except (TimeoutError, PlaywrightTimeoutError, ValueError, RuntimeError) as e:
            safe_print(f"実行エラー/タイムアウト: {e}")
            try:
                safe_print(f"最終URL: {page.url}")
                await page.screenshot(path="gemini_timeout.png", full_page=True)
                safe_print("スクリーンショットを gemini_timeout.png に保存しました。")
            except Exception:
                pass
        finally:
            if owned_context is not None:
                await owned_context.close()
            if started_chrome:
                await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
