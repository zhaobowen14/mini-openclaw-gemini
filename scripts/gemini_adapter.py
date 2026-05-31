from __future__ import annotations

import json
import re
import subprocess
from pathlib import Path
from typing import Any


class GeminiAdapter:
    def __init__(self, workspace: str | Path):
        self.workspace = Path(workspace)
        self.script = self.workspace / "scripts" / "gemini_playwright_chat.py"
        self.output_dir = self.workspace / "outputs"
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _normalize_response_text(self, text: str) -> str:
        if not isinstance(text, str):
            return ""
        return text.replace("実行する", "")

    def complete(self, prompt: str) -> str:
        output_path = self.output_dir / "last_response.json"
        prompt_path = self.output_dir / "last_prompt.json"
        prompt_path.write_text(prompt, encoding="utf-8")

        command = [
            "python",
            str(self.script),
            "--prompts-file",
            str(prompt_path),
            "--output",
            str(output_path),
            "--delay",
            "0",
        ]

        completed = subprocess.run(
            command,
            cwd=str(self.workspace),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )

        # if completed.stdout:
        #     print(completed.stdout)

        if completed.returncode != 0:
            raise RuntimeError(
                "Gemini adapter failed\n"
                f"stdout:\n{completed.stdout}\n\n"
                f"stderr:\n{completed.stderr}"
            )

        return self._extract_text(output_path)

    def _extract_text(self, output_path: Path) -> str:
        if not output_path.exists():
            raise FileNotFoundError(f"Gemini output not found: {output_path}")

        raw = json.loads(output_path.read_text(encoding="utf-8"))

        if isinstance(raw, dict):
            for key in ("response_text", "response", "text", "content", "answer"):
                value = raw.get(key)
                if isinstance(value, str) and value.strip():
                    return self._normalize_response_text(value).strip()

        if isinstance(raw, list) and raw:
            last = raw[-1]
            if isinstance(last, dict):
                for key in ("answer", "response_text", "response", "text", "content"):
                    value = last.get(key)
                    if isinstance(value, str) and value.strip():
                        return self._normalize_response_text(value).strip()

        return self._normalize_response_text(json.dumps(raw, ensure_ascii=False))


def build_protocol_prompt(
    system_prompt: str,
    history: list[dict[str, Any]],
    user_message: str,
    tools: list[dict[str, Any]],
) -> str:
    payload = {
        "system": system_prompt,
        "tools": tools,
        "history": history,
        "user_message": user_message,
        "required_output_format": {
            "final": {"type": "final", "text": "string"},
            "tool_calls": {
                "type": "tool_calls",
                "tool_calls": [
                    {"id": "call_1", "name": "tool_name", "arguments": {}}
                ],
            },
            "python_code_block": "```python\\n# code here\\n```",
        },
        "rules": [
            "For normal answers, return JSON only.",
            "For substantial Python code generation, return only a Python code block.",
            "ユーザーの最初の依頼を確認し、目的を達成しているかチェック",
            "エラーがでたら、内容を確認して修正。必要に応じてtool readしてプログラムの内容を確認して",
            "If there are any additional requests after the final, must not return final again but proceed them.",
        ],
    }
    return json.dumps(payload, ensure_ascii=False, indent=None)


def parse_protocol_response(text: str) -> dict[str, Any]:
    original_text = text
    text = text.strip() if isinstance(text, str) else ""

    errors: list[str] = []

    def _short(s: str, limit: int = 4000) -> str:
        if len(s) <= limit:
            return s
        return s[:limit] + "\n... [truncated] ..."

    def _escape_newlines_in_text_field(raw: str) -> str:
        pattern = r'("text"\s*:\s*")([\s\S]*?)(")(?=\s*[},])'

        def repl(match: re.Match[str]) -> str:
            prefix, body, suffix = match.groups()
            body = body.replace("\\", "\\\\")
            body = body.replace('"', '\\"')
            body = body.replace("\r", "\\r").replace("\n", "\\n")
            return f"{prefix}{body}{suffix}"

        return re.sub(pattern, repl, raw, count=1)

    def _coerce(obj: Any) -> dict[str, Any]:
        if isinstance(obj, dict):
            if "type" not in obj and "tool_calls" in obj:
                obj["type"] = "tool_calls"
            elif "type" not in obj and "text" in obj:
                obj["type"] = "final"
            return obj

        if isinstance(obj, list):
            return {"type": "final", "text": json.dumps(obj, ensure_ascii=False)}

        if isinstance(obj, str):
            return {"type": "final", "text": obj}

        return {"type": "final", "text": json.dumps(obj, ensure_ascii=False)}

    def _fix_broken_json(raw: str) -> str:
        # command内のURLを囲む " が未エスケープで壊れたケースを救済する
        # 例:
        # "command": "py script.py "https://example.com?a=1&b=2" --mode visible-only"
        pattern = r'("command"\s*:\s*")([\s\S]*?)("\s*[,}])'

        def repl(match: re.Match[str]) -> str:
            prefix, body, suffix = match.groups()
            body = body.replace("\\", "\\\\")
            body = body.replace('"', '\\"')
            body = body.replace("\r", "\\r").replace("\n", "\\n")
            return f"{prefix}{body}{suffix}"

        return re.sub(pattern, repl, raw, count=1)

    # 1. Pythonコードブロック最優先
    code_block = re.search(
        r"```python\s*(.*?)\s*```",
        text,
        re.DOTALL | re.IGNORECASE,
    )
    if code_block:
        code = code_block.group(1).strip()
        filename = "generated_script.py"
        if "from pptx" in code or "Presentation(" in code:
            filename = "generated_presentation.py"

        return {
            "type": "tool_calls",
            "tool_calls": [
                {
                    "id": "auto_extracted_write",
                    "name": "write",
                    "arguments": {
                        "path": filename,
                        "content": code,
                    },
                }
            ],
        }

    # 2. "Python\n..." 形式も救済
    python_plain = re.match(r"^Python\s*\n([\s\S]*)", text, re.IGNORECASE)
    if python_plain:
        code = python_plain.group(1).strip()
        filename = "generated_script.py"
        if "from pptx" in code or "Presentation(" in code:
            filename = "generated_presentation.py"

        return {
            "type": "tool_calls",
            "tool_calls": [
                {
                    "id": "auto_extracted_write_plain",
                    "name": "write",
                    "arguments": {
                        "path": filename,
                        "content": code,
                    },
                }
            ],
        }

    # 3. 前処理
    text = re.sub(r"^\s*(json|JSON)\s*", "", text)

    def extract_first_json(raw: str) -> str | None:
        start = raw.find("{")
        if start == -1:
            return None

        depth = 0
        in_string = False
        escape = False

        for i in range(start, len(raw)):
            ch = raw[i]

            if in_string:
                if escape:
                    escape = False
                elif ch == "\\":
                    escape = True
                elif ch == '"':
                    in_string = False
                continue

            if ch == '"':
                in_string = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return raw[start : i + 1]

        return None

    json_text = extract_first_json(text)
    if json_text:
        text = json_text

    repaired_text = _fix_broken_json(text)

    candidates = [
        ("raw", text),
        ("repaired", repaired_text),
        ("escaped_text_field", _escape_newlines_in_text_field(repaired_text)),
    ]

    # 4. raw_decode
    for label, candidate in candidates:
        try:
            decoder = json.JSONDecoder()
            obj, end = decoder.raw_decode(candidate)
            return _coerce(obj)
        except Exception as e:
            msg = f"{label} RAW_DECODE ERROR: {repr(e)}"
            print(msg)
            errors.append(msg)

    # 5. json.loads
    for label, candidate in candidates:
        try:
            obj = json.loads(candidate)
            return _coerce(obj)
        except Exception as e:
            msg = f"{label} JSON_LOAD ERROR: {repr(e)}"
            print(msg)
            errors.append(msg)

    # 6. 失敗時は、empty responseで潰さず、原因と生テキストを返す
    return {
        "type": "error",
        "text": (
            "parse_protocol_response failed\n\n"
            "errors:\n"
            + "\n".join(errors)
            + "\n\nraw_text:\n"
            + _short(str(original_text))
            + "\n\nextracted_text:\n"
            + _short(text)
            + "\n\nrepaired_text:\n"
            + _short(repaired_text)
        ),
    }