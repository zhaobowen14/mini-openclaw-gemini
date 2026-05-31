# python "C:\Users\zhaob\Desktop\mini-openclaw-gemini_single\scripts\playwright_dom_analyzer.py" "https://chatgpt.com/" --headful --wait-ms 5000 --mode visible-only

from __future__ import annotations

import argparse
import asyncio
import json
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description='JS実行後の現在DOMを取得して分析する Playwright ツール'
    )
    parser.add_argument('url', help='分析対象URL')
    parser.add_argument(
        '--output',
        default='',
        help='出力先JSON。未指定なら scripts/../outputs/dom_analysis_<timestamp>.json',
    )
    parser.add_argument('--headful', action='store_true', help='ブラウザを表示して実行する')
    parser.add_argument('--wait-ms', type=int, default=3000, help='page.goto後に追加で待つ時間(ms)')
    parser.add_argument('--timeout-ms', type=int, default=30000, help='page.gotoなどのタイムアウト(ms)')
    parser.add_argument('--screenshot', action='store_true', help='スクリーンショットも保存する')
    parser.add_argument(
        '--mode',
        choices=['full', 'visible-only'],
        default='full',
        help='full=全分析, visible-only=可視要素一覧中心の軽量出力',
    )
    return parser.parse_args()


async def collect_dom_snapshot(page, mode: str = 'full'):
    return await page.evaluate(
        """
(mode) => {
  const normalizeText = (value, maxLen = 200) => {
    return (value || '').replace(/\s+/g, ' ').trim().slice(0, maxLen);
  };

  const cssPath = (el) => {
    if (!(el instanceof Element)) return '';
    const path = [];
    let current = el;
    while (current && current.nodeType === Node.ELEMENT_NODE && path.length < 6) {
      let selector = current.tagName.toLowerCase();
      if (current.id) {
        selector += '#' + current.id;
        path.unshift(selector);
        break;
      }
      const classList = Array.from(current.classList || []).slice(0, 3);
      if (classList.length) {
        selector += '.' + classList.join('.');
      }
      const parent = current.parentElement;
      if (parent) {
        const sameTagSiblings = Array.from(parent.children).filter(x => x.tagName === current.tagName);
        if (sameTagSiblings.length > 1) {
          selector += `:nth-of-type(${sameTagSiblings.indexOf(current) + 1})`;
        }
      }
      path.unshift(selector);
      current = current.parentElement;
    }
    return path.join(' > ');
  };

  const isVisible = (el) => {
    const style = window.getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    return style.display !== 'none' &&
           style.visibility !== 'hidden' &&
           style.opacity !== '0' &&
           rect.width > 0 &&
           rect.height > 0;
  };

  const summarize = (el) => {
    const rect = el.getBoundingClientRect();
    return {
      tag: el.tagName.toLowerCase(),
      id: el.id || '',
      className: typeof el.className === 'string' ? el.className : '',
      role: el.getAttribute('role') || '',
      name: el.getAttribute('name') || '',
      type: el.getAttribute('type') || '',
      href: el.getAttribute('href') || '',
      src: el.getAttribute('src') || '',
      ariaLabel: el.getAttribute('aria-label') || '',
      placeholder: el.getAttribute('placeholder') || '',
      dataTestId: el.getAttribute('data-testid') || '',
      contentEditable: el.getAttribute('contenteditable') || '',
      text: normalizeText(el.innerText || el.textContent || ''),
      visible: isVisible(el),
      disabled: !!el.disabled || el.getAttribute('aria-disabled') === 'true',
      rect: {
        x: Math.round(rect.x),
        y: Math.round(rect.y),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
      },
      cssPath: cssPath(el),
      outerHTML: (el.outerHTML || '').slice(0, 2000),
    };
  };

  const all = Array.from(document.querySelectorAll('*'));
  const visible = all.filter(isVisible);

  const forms = all.filter(el =>
    ['input', 'textarea', 'select', 'button'].includes(el.tagName.toLowerCase()) ||
    el.getAttribute('role') === 'textbox' ||
    el.getAttribute('contenteditable') === 'true'
  );

  const buttons = all.filter(el =>
    el.tagName.toLowerCase() === 'button' ||
    el.getAttribute('role') === 'button' ||
    /send|submit|送信/i.test(el.getAttribute('aria-label') || '') ||
    /send|submit|送信/i.test(el.innerText || '')
  );

  const articles = all.filter(el =>
    ['article', 'main', 'section'].includes(el.tagName.toLowerCase()) ||
    ['article', 'main'].includes(el.getAttribute('role') || '') ||
    !!el.getAttribute('data-message-author-role')
  );

  const roleSummary = {};
  for (const el of all) {
    const role = el.getAttribute('role');
    if (!role) continue;
    roleSummary[role] = (roleSummary[role] || 0) + 1;
  }

  const testIdSummary = [];
  const seenTestIds = new Set();
  for (const el of all) {
    const testid = el.getAttribute('data-testid');
    if (!testid || seenTestIds.has(testid)) continue;
    seenTestIds.add(testid);
    testIdSummary.push({
      dataTestId: testid,
      tag: el.tagName.toLowerCase(),
      text: normalizeText(el.innerText || el.textContent || '', 80),
      cssPath: cssPath(el),
    });
  }

  const base = {
    title: document.title,
    url: location.href,
    bodyText: normalizeText(document.body ? document.body.innerText : '', 10000),
    counts: {
      all: all.length,
      visible: visible.length,
      forms: forms.length,
      buttons: buttons.length,
      articles: articles.length,
    },
    visibleElements: visible.slice(0, 300).map(summarize),
  };

  if (mode === 'visible-only') {
    return base;
  }

  return {
    ...base,
    html: document.documentElement ? document.documentElement.outerHTML : '',
    roleSummary,
    testIdSummary: testIdSummary.slice(0, 300),
    formLikeElements: forms.slice(0, 100).map(summarize),
    buttonLikeElements: buttons.slice(0, 100).map(summarize),
    articleLikeElements: articles.slice(0, 100).map(summarize),
  };
}
        """,
        mode,
    )


async def main() -> None:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    project_dir = script_dir.parent
    outputs_dir = project_dir / 'outputs'
    outputs_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = Path(args.output) if args.output else outputs_dir / f'dom_analysis_{timestamp}.json'
    screenshot_path = output_path.with_suffix('.png')

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=not args.headful)
        page = await browser.new_page(viewport={'width': 1440, 'height': 900})
        page.set_default_timeout(args.timeout_ms)

        await page.goto(args.url, wait_until='domcontentloaded', timeout=args.timeout_ms)
        if args.wait_ms > 0:
            await page.wait_for_timeout(args.wait_ms)

        snapshot = await collect_dom_snapshot(page, args.mode)
        snapshot['meta'] = {
            'capturedAt': datetime.now().isoformat(timespec='seconds'),
            'waitMs': args.wait_ms,
            'timeoutMs': args.timeout_ms,
            'headful': args.headful,
            'mode': args.mode,
        }

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding='utf-8')

        if args.screenshot:
            await page.screenshot(path=str(screenshot_path), full_page=True)

        print(f'保存先: {output_path}')
        if args.screenshot:
            print(f'スクリーンショット: {screenshot_path}')
        print(f"title: {snapshot['title']}")
        print(f"url: {snapshot['url']}")
        print(f"counts: {snapshot['counts']}")

        await browser.close()


if __name__ == '__main__':
    asyncio.run(main())
