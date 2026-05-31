from __future__ import annotations

import sys
sys.stdout.reconfigure(encoding="utf-8")

import argparse
import json
import re
from urllib.parse import parse_qs, urlparse

from playwright.sync_api import sync_playwright
from youtube_transcript_api import YouTubeTranscriptApi

def extract_video_id(url: str) -> str | None:
    url = url.strip().strip("'\"")
    parsed = urlparse(url)

    # youtu.be
    if "youtu.be" in parsed.netloc:
        return parsed.path.lstrip("/") or None

    # watch?v=
    qs = parse_qs(parsed.query)
    if "v" in qs and qs["v"]:
        return qs["v"][0]

    # /shorts/
    m = re.search(r"/shorts/([A-Za-z0-9_-]{11})", parsed.path)
    if m:
        return m.group(1)

    # /embed/ or /live/
    m = re.search(r"/(embed|live)/([A-Za-z0-9_-]{11})", parsed.path)
    if m:
        return m.group(2)

    return None


def get_title_with_playwright(url: str) -> str | None:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(locale="ja-JP")
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(3000)
            title = page.title()
            return title.replace(" - YouTube", "") if title else None
        finally:
            browser.close()


def fetch_transcript(video_id: str, language: str = "ja"):
    api = YouTubeTranscriptApi()
    return api.fetch(video_id, languages=[language])


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("url")
    parser.add_argument("--language", default="ja")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    video_id = extract_video_id(args.url)
    if not video_id:
        raise SystemExit("Could not extract video ID from URL")

    title = get_title_with_playwright(args.url)

    # title = None
    transcript = fetch_transcript(video_id, language=args.language)

    items = [
        {"text": item.text, "start": item.start, "duration": item.duration}
        for item in transcript
    ]

    if args.json:
        print(json.dumps({"video_id": video_id, "title": title, "transcript": items}, ensure_ascii=False, indent=2))
    else:
        print(f"video_id: {video_id}")
        if title:
            print(f"title: {title}")
        print("\n".join(item["text"] for item in items))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
