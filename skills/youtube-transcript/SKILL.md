---
name: youtube-transcript
description: Fetch subtitles/transcripts from YouTube videos and Shorts, then return clean text or timestamped JSON. 
---

# YouTube Transcript

URL/タイトル確認だけ Playwright を使い、字幕本文は `youtube-transcript-api` で取得する。

## Quick start

1. YouTube URL から動画IDを抽出する。
2. `scripts/youtube_transcript_from_url.py` を実行する。
3. 必要に応じて、取得した字幕を要約・整形・保存する。

基本コマンド:

```powershell
py ./skills/youtube-transcript/scripts/youtube_transcript_from_url.py 'https://youtu.be/VIDEO_ID'
```

JSONで欲しいとき:

```powershell
py ./skills/youtube-transcript/scripts/youtube_transcript_from_url.py 'https://youtu.be/VIDEO_ID' --json
```

言語を変えるとき(基本的に日本語で実施):

```powershell
py ./skills/youtube-transcript/scripts/youtube_transcript_from_url.py 'https://youtu.be/VIDEO_ID' --language jp --json
```

## Workflow

### 1. URLを受け取る

対応対象:
- `https://www.youtube.com/watch?v=...`
- `https://youtu.be/...`
- `https://www.youtube.com/shorts/...`

### 2. まずローカル取得を試す

`scripts/youtube_transcript_from_url.py` を使う。このスクリプトは:
- URLから動画IDを抽出する
- Playwrightで動画タイトルを確認する
- `youtube-transcript-api` で字幕本文を取得する
- プレーンテキストまたはJSONで返す

### 3. 失敗時の判断


## Practical guidance

- `--language ja` を先に試す。だめなら英語を検討する。

## Output patterns

### プレーンテキスト向き
- ざっくり内容確認
- 要約の下準備
- 人間がそのまま読む用途

### JSON向き
- 後処理
- タイムスタンプ付き要約
- ファイル保存
- 他スクリプトへの受け渡し

## Resources

### scripts/

- `youtube_transcript_from_url.py`: YouTube URL からタイトル確認と字幕取得を行うメインスクリプト
