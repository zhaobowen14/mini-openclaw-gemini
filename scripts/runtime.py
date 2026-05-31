from __future__ import annotations

import subprocess
import atexit
import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any

# Force UTF-8 for stdout/stderr to prevent UnicodeEncodeError on Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from cron_manager import CronManager, NaturalLanguageCronParser
from gemini_adapter import GeminiAdapter, build_protocol_prompt, parse_protocol_response
from session_store import SessionStore
from skill_loader import SkillLoader
from tool_runner import ToolRunner

RED = "\033[31m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
RESET = "\033[0m"


WORKSPACE = Path(__file__).resolve().parents[1]
SESSIONS_DIR = WORKSPACE / 'sessions'
OUTPUTS_DIR = WORKSPACE / 'outputs'
SKILLS_DIR = WORKSPACE / 'skills'
CRON_FILE = WORKSPACE / 'cron' / 'jobs.json'

SYSTEM_PROMPT = r"""You are a strict protocol agent.
You MUST follow the output format exactly.
ALLOWED OUTPUTS ONLY:
1. Valid JSON(example):
{"type": "final", "text": "..."}

2. A Python code block only:
```python
# code
```
要件:
・指定フォーマット以外は出力しない
・有効なJSONまたはPythonコードブロックのみ
・不完全なJSONは禁止
・JSON内の改行はすべて\nでエスケープ
・JSON内のコマンド引数はシングルクォートを使用
・Windowsパスは「/」を使用
・説明文の前後出力は禁止
・read/write/execはPython関数として使わない

ツール利用:
・ローカルファイルは、readツールを使用して読み込み
・httpから始まる文字列はURLとみなし、readツールを使わないこと
・ツールは必ずJSONのtool_calls形式で呼び出す
・write完了後にexecを実行

Python実行:
・pyコマンドを使用
・シングルクォートを使用
・処理停止系（pltなど）は使わない

データ処理:
・CSVでカンマを含む文字列はダブルクォートで囲む
・Unicodeエラー時は「set PYTHONIOENCODING=utf-8 &&」を付与

数式:
・textには書かない
・必要な場合はmathフィールドにLaTeXで出力

Web:
・最新情報はsearch_webを使用（最大5回とすること）
・URL取得後は必ずfetch_urlで確認

その他:
・フォルダ指定時はPythonで一覧取得
・readツールでlsやdirは使わない
・違反した場合は再出力する
・パワポはpython-pptx、エクセルはopenpyxl、ワードはpython-docxで生成
・エクセルの読み込みは、openpyxl.load_workbookで行うこと
"""

TOOLS = [
    {
        'name': 'read',
        'description': 'Read a local text file or PDF from workspace.',
        'arguments': {'path': 'string'},
    },
    {
        'name': 'write',
        'description': 'Write a text file in workspace outputs.',
        'arguments': {'path': 'string', 'content': 'string'},
    },
    {
        'name': 'exec',
        'description': 'Run a shell command in workspace outputs.',
        'arguments': {'command': 'string', 'cwd': 'string?', 'timeout': 'int?'},
    },
    {
        'name': 'fetch_url',
        'description': 'Fetch a web page and extract readable text.',
        'arguments': {'url': 'string', 'max_chars': 'int?'},
    },
    {
        'name': 'search_web',
        'description': 'Search the web using Brave Search API and return top results.',
        'arguments': {'query': 'string', 'count': 'int?'},
    },
]


class MiniRuntime:
    def __init__(self) -> None:
        self.store = SessionStore(SESSIONS_DIR)
        self.tools = ToolRunner(WORKSPACE)
        self.gemini = GeminiAdapter(WORKSPACE)
        self.skill_loader = SkillLoader(SKILLS_DIR)
        self.cron_manager = CronManager(CRON_FILE, WORKSPACE / 'scripts' / 'runtime.py', WORKSPACE)
        self.cron_parser = NaturalLanguageCronParser()

    def _slugify_skill_name(self, text: str) -> str:
        value = unicodedata.normalize('NFKC', text).strip().lower()
        value = re.sub(r'[^a-z0-9\-\s_]+', '-', value)
        value = re.sub(r'[\s_]+', '-', value)
        value = re.sub(r'-+', '-', value).strip('-')
        return value or 'new-skill'

    def _skill_creation_request(self, text: str) -> dict[str, str] | None:
        text = text.strip()

        parts = text.split(maxsplit=1)
        if len(parts) < 2:
            return None

        base_name = parts[0]
        description = parts[1].strip('"')

        return {
            'name': self._slugify_skill_name(base_name),
            'description': description or 'fill this skill description',
            'request': text,
        }

    def _create_skill_from_request(self, payload: dict[str, str]) -> str:
        skill_name = payload['name']
        description = payload['description']

        skill_dir = SKILLS_DIR / skill_name
        scripts_dir = skill_dir / 'scripts'   # ★追加

        skill_dir.mkdir(parents=True, exist_ok=True)
        scripts_dir.mkdir(parents=True, exist_ok=True)  # ★追加

        skill_file = skill_dir / 'SKILL.md'
        title = skill_name

        body = (
            '---\n'
            f'name: {skill_name}\n'
            f'description: {description}\n'
            '---\n\n'
            f'# {title}\n\n'
            '## Purpose\n\n'
            f'- {description}\n'
            '## Behavior\n\n'
            '- 必要なら read / write / exec を使ってローカル作業を進める。\n'
            '- まず状況確認をしてから、最小限の手順で目的を達成する。\n'
            '- 結果は簡潔に返し、必要なら次のアクションも提案する。\n'
        )

        skill_file.write_text(body, encoding='utf-8')

        return (
            f'skill を作成しました: {skill_name}\n'
            f'- path: skills/{skill_name}/SKILL.md\n'
            f'- scripts: skills/{skill_name}/scripts/\n'  # ★追加
            '- 同名 skill があった場合は上書きしています。\n'
            f'- description: {description}'
        )
    
    def _user_requires_tool(self, text: str) -> bool:
        lowered = text.lower()
        explicit_patterns = ['read tool', 'write tool', 'exec tool', 'search_web', 'fetch_url']
        return any(p in lowered for p in explicit_patterns)

    def _history_has_tool_result(self, session: Any) -> bool:
        return any(isinstance(m, dict) and m.get('role') == 'tool' for m in session.messages)

    def _needs_web_search(self, text: str) -> bool:
        lowered = text.lower()
        jp_terms_lower = ['web検索', 'ウェブ検索', 'ネットで調べて', 'webで調べて']
        en_terms = ['web search', 'search the web']
        return any(term in lowered for term in en_terms) or any(term in lowered for term in jp_terms_lower)

    def _needs_fetch_after_search(self, text: str) -> bool:
        return any(k in text.lower() for k in [
            'summarize', 'summary', 'content', 'body', 'title', 'fetch', 'open the page'
        ]) or any(k in text for k in ['要約', '本文', '内容', '見て', '確認', 'タイトル'])

    def _history_has_tool_name(self, session: Any, tool_name: str) -> bool:
        for message in session.messages:
            if not isinstance(message, dict) or message.get('role') != 'tool':
                continue
            content = message.get('content', [])
            if not isinstance(content, list):
                continue
            for item in content:
                if isinstance(item, dict) and item.get('name') == tool_name:
                    return True
        return False

    def _suggest_output_filename(self, source_path: str, content: str) -> str:
        # source = Path(source_path).name if source_path else 'generated_script.py'
        return source_path

    def _fix_generated_python(self, content: str) -> str:
        if not isinstance(content, str):
            return ''

        content = content.replace("if name == 'main':", 'if __name__ == "__main__":')
        content = content.replace('if name == "main":', 'if __name__ == "__main__":')
        content = content.replace('if __name__ == \\"__main__\\":', 'if __name__ == "__main__":')
        content = content.replace('\rif __name__ ==', '\nif __name__ ==')
        content = content.replace('\r\n', '\n')
        return content

    def _handle_command(self, session_id: str, user_message: str) -> str | None:
        text = user_message.strip()
        if not text:
            return None

        if text.lower() in ['/skills', '/skill list']:
            skills = self.skill_loader.list_skills()
            if not skills:
                return 'skill はまだありません。mini-openclaw-gemini/skills/<name>/SKILL.md を追加してください。'
            lines = ['利用可能な skill:']
            for skill in skills:
                lines.append(f'- {skill.name}: {skill.description}')
            return '\n'.join(lines)

        if text.lower().startswith('/skill show'):
            name = text.split(None, 2)[2].strip()
            skill = self.skill_loader.get_skill(name)
            if not skill:
                return f'skill が見つかりません: {name}'
            return f'{skill.name}\n{skill.description}\n\n{skill.instructions[:3000]}'


        if text.lower().startswith('/skill add'):
            prefix = '/skill add'
            rest = text[len(prefix):].strip()

            # ① コマンド形式（kabu "説明"）
            creation_request = self._skill_creation_request(rest)

            if creation_request is not None:
                return self._create_skill_from_request(creation_request)

            return '形式が不正です。例: /skill add test-skill "～のためのスキル"'

        if text.lower().startswith('/skill use'):
            explicit_skill_name, remaining = self.skill_loader.strip_explicit_skill_prefix(text)
            if not explicit_skill_name:
                return 'skill 名を指定してください。'
            skill = self.skill_loader.get_skill(explicit_skill_name)
            if not skill:
                return f'skill が見つかりません: {explicit_skill_name}'

        if text.lower() in ['/cron', '/cron list']:
            jobs = self.cron_manager.load_jobs()
            if not jobs:
                return 'cron はまだありません。'
            lines = ['cron 一覧:']
            for job in jobs:
                lines.append(
                    f'- {job.id} | enabled={job.enabled} | type={job.schedule_type} | next={job.next_run_at} | message={job.message}'
                )
            return '\n'.join(lines)

        if text.lower().startswith('/cron remove'):
            job_id = text.split(None, 2)[2].strip()
            removed = self.cron_manager.remove_job(job_id)
            return f'cron 削除: {job_id} -> {removed}'

        if text.lower().startswith('/cron disable'):
            job_id = text.split(None, 2)[2].strip()
            changed = self.cron_manager.set_enabled(job_id, False)
            return f'cron 停止: {job_id} -> {changed}'

        if text.lower().startswith('/cron enable'):
            job_id = text.split(None, 2)[2].strip()
            changed = self.cron_manager.set_enabled(job_id, True)
            return f'cron 再開: {job_id} -> {changed}'

        parsed = self.cron_parser.parse(text, session=session_id)
        if parsed:
            action = parsed['action']
            if action == 'list':
                jobs = self.cron_manager.load_jobs()
                if not jobs:
                    return 'cron はまだありません。'
                lines = ['cron 一覧:']
                for job in jobs:
                    lines.append(
                        f'- {job.id} | enabled={job.enabled} | type={job.schedule_type} | next={job.next_run_at} | message={job.message}'
                    )
                return '\n'.join(lines)
            if action == 'disable':
                changed = self.cron_manager.set_enabled(parsed['id'], False)
                return f'cron 停止: {parsed["id"]} -> {changed}'
            if action == 'enable':
                changed = self.cron_manager.set_enabled(parsed['id'], True)
                return f'cron 再開: {parsed["id"]} -> {changed}'
            if action == 'add':
                existing = next((job for job in self.cron_manager.load_jobs() if job.id == parsed['id']), None)
                if existing is not None:
                    return (
                        f'cron は既に存在します: {existing.id}\n'
                        f'- type: {existing.schedule_type}\n'
                        f'- next: {existing.next_run_at}\n'
                        f'- message: {existing.message}'
                    )
                job = self.cron_manager.add_job(
                    parsed['id'],
                    parsed['message'],
                    parsed['interval_seconds'],
                    session=parsed.get('session', session_id),
                    schedule_type=parsed.get('schedule_type', 'interval'),
                    time_of_day=parsed.get('time_of_day'),
                    weekday=parsed.get('weekday'),
                )
                return (
                    f'cron 登録: {job.id}\n'
                    f'- type: {job.schedule_type}\n'
                    f'- next: {job.next_run_at}\n'
                    f'- message: {job.message}'
                )

        return None

    def _build_user_message(self, original_user_message: str) -> str:
        explicit_skill_name, stripped_message = self.skill_loader.strip_explicit_skill_prefix(original_user_message)

        skill = None
        base_message = original_user_message

        if explicit_skill_name:
            skill = self.skill_loader.get_skill(explicit_skill_name)
            base_message = stripped_message or f'Use skill {explicit_skill_name}.'
        else:
            skill = self.skill_loader.match_skill(original_user_message)
            base_message = original_user_message

        skill_prompt = self.skill_loader.build_skill_prompt(skill)
        # print(f'DEBUG: matched skill: {skill.name if skill else None}, explicit_skill_name: {explicit_skill_name}, skill_prompt: {skill_prompt}')


        if not skill_prompt:
            return base_message
        return (
            f'{base_message}\n\n'
            'The following skill is active. Follow it when it helps answer the user.\n'
            f'{skill_prompt}'
        )

    def handle_one_turn(self, session_id: str, user_message: str) -> str:
        command_reply = self._handle_command(session_id, user_message)

        if command_reply is not None:
            session = self.store.load(session_id)
            # self.store.append(session, 'user', user_message)
            self.store.append(session, 'assistant', command_reply, kind='command')
            self.store.save(session)
            return command_reply

        if user_message.strip().lower() == '/reset':
            try:
                session_path = self.store._get_session_path(session_id)
                if session_path.exists():
                    session_path.unlink()
                return 'セッションをリセットしました。'
            except Exception as e:
                return f'リセット失敗: {e}'
            
        if user_message.strip().lower() == '/clear':
            try:
                session = self.store.load(session_id)

                session.messages = [
                    m for m in session.messages
                    if not (
                        m.get("role") == "tool" or
                        (m.get("role") == "assistant" and m.get("kind") == "tool_calls")
                    )
                ]

                self.store.save(session)

                return 'tool履歴のみ削除しました。'

            except Exception as e:
                return f'削除失敗: {e}'

        session = self.store.load(session_id)
        

        final_text = None
        tool_required = self._user_requires_tool(user_message)
        model_user_message = self._build_user_message(user_message)

        



        for _ in range(5):
            self.store.append(session, 'user', model_user_message)

            prompt = build_protocol_prompt(
                system_prompt=SYSTEM_PROMPT,
                history=session.messages,
                user_message=model_user_message,
                tools=TOOLS,
            )

            raw = self.gemini.complete(prompt)
            print(f'DEBUG: raw: {raw}')

            action = parse_protocol_response(raw)

            if not isinstance(action, dict):
                final_text = str(action)
                self.store.append(session, 'assistant', final_text)
                break

            action_type = action.get('type')
            

            if action_type == 'error':
                print("DEBUG: invalid response from LLM")
                self.store.append(session, 'assistant', raw, kind='parse_error')
                model_user_message = (
                    'Your response was invalid. Return either '
                    '{"type": "final", "text": "..."} '
                    'JSONは必ずエスケープ処理'
                    'For Python code generation, return only a Python code block.'
                    'JSON strings must be enclosed in double quotes (")- Double quotes inside JSON strings must be escaped as \\" Do not use unescaped double quotes inside a JSON string as it will break JSON parsing'
                    '- JSON内のコマンド引数は原則シングルクォートを使用。ただしWindows cmdでURLを渡す場合は、URLをダブルクォートで囲み、JSON内では必ず \" としてエスケープする。'
                )
                continue

            if action_type == 'final':
                if tool_required and not self._history_has_tool_result(session):
                    self.store.append(session, 'assistant', raw, kind='rejected_final_before_tool')
                    model_user_message = (
                        'You returned final before using the required tool. '
                        'Return tool_calls or a Python code block first.'
                    )
                    continue

                if self._needs_web_search(user_message) and not self._history_has_tool_name(session, 'search_web'):
                    self.store.append(session, 'assistant', raw, kind='rejected_final_need_web_search')
                    model_user_message = (
                        'The user explicitly asked for web search. '
                        'Use search_web before returning final.'
                    )
                    continue

                if (
                    self._needs_web_search(user_message)
                    and self._needs_fetch_after_search(user_message)
                    and self._history_has_tool_name(session, 'search_web')
                    and not self._history_has_tool_name(session, 'fetch_url')
                ):
                    self.store.append(session, 'assistant', raw, kind='rejected_final_need_fetch_after_search')
                    model_user_message = (
                        'You searched the web, but the user asked for page content/title/summary. '
                        'Fetch the most relevant URL with fetch_url before returning final.'
                    )
                    continue

                final_text = action.get('text', '')
                self.store.append(session, 'assistant', final_text)
                break

            if action_type == 'tool_calls':
                tool_calls = action.get('tool_calls', [])
                self.store.append(session, 'assistant', raw, kind='tool_calls')

                tool_results: list[dict[str, Any]] = []
                for tc in tool_calls:
                    if not isinstance(tc, dict):
                        tool_results.append({'ok': False, 'error': f'invalid tool call: {tc}'})
                        continue

                    name = tc.get('name')
                    arguments = tc.get('arguments', {})
                    if not isinstance(name, str):
                        fn = tc.get('function')
                        if isinstance(fn, dict):
                            name = fn.get('name')
                            arguments = fn.get('arguments', {})
                    if isinstance(arguments, str):
                        try:
                            arguments = json.loads(arguments)
                        except Exception:
                            arguments = {}
                    if not isinstance(name, str):
                        tool_results.append({'ok': False, 'error': f'missing tool name: {tc}'})
                        continue
                    if not isinstance(arguments, dict):
                        arguments = {}

                    if name == 'write':
                        content = arguments.get('content', '')
                        if not isinstance(content, str):
                            content = ''
                        # fixed_content = self._fix_generated_python(content)
                        fixed_content = content

                        source_path = arguments.get('path', 'generated_script.py')
                        if not isinstance(source_path, str):
                            source_path = 'generated_script.py'

                        output_path = self._suggest_output_filename(source_path, fixed_content)
                        write_args = {'path': output_path, 'content': fixed_content}
                        write_result = self.tools.run('write', write_args)
                        result: dict[str, Any] = {'write': write_result}

                        if write_result.get('ok'):
                            syntax_result = self.tools.run(
                                'exec',
                                {
                                    'command': f'python -m py_compile "{output_path}"',
                                    'cwd': '.',
                                    'timeout': 30,
                                },
                            )
                            result['syntax_check'] = syntax_result

                            exec_result = self.tools.run(
                                'exec',
                                {
                                    'command': f'python "{output_path}"',
                                    'cwd': '.',
                                    'timeout': 120,
                                },
                            )
                            result['exec'] = exec_result
                    else:
                        result = self.tools.run(name, arguments)

                    tool_results.append(
                        {
                            'tool_call_id': tc.get('id'),
                            'name': name,
                            'result': result,
                        }
                    )


                self.store.append(session, 'tool', tool_results)
                model_user_message = 'Continue. You have received tool results in history.'
                continue

            final_text = action.get('text', raw)
            self.store.append(session, 'assistant', final_text)
            break

        if final_text is None:
            final_text = 'Stopped after max steps.'
            self.store.append(session, 'assistant', final_text)

        self.store.save(session)
        return final_text


from rich.console import Console
from rich.markdown import Markdown

console = Console()

def pretty_print(text: str):
    try:
        text = text.strip()
        text = text.replace("\\n", "\n")
        console.print(Markdown(text))
    except Exception:
        console.print(text)


def main() -> None:

    try:
        session_dir = Path("./sessions")
        if not session_dir.exists():
            print("sessionsフォルダが存在しません")
        else:
            for f in session_dir.glob("*"):
                if f.is_file():
                    f.unlink()

        print('全セッションを削除しました。')

    except Exception as e:
        print(f'sessionsリセット失敗: {e}')

    parser = argparse.ArgumentParser()
    parser.add_argument('--session', default='default')
    parser.add_argument('--message', default=None)
    parser.add_argument('--message-file', default=None)
    args = parser.parse_args()

    runtime = MiniRuntime()

    message = args.message
    if args.message_file:
        message = Path(args.message_file).read_text(encoding='utf-8')

    if message:
        reply = runtime.handle_one_turn(args.session, message)
        print(f'{GREEN}Assistant:{RESET} ', end='')
        pretty_print(reply)
        return

    print("Interactive mode. Type 'exit' to quit.\n")

    while True:
        try:
            user_input = input(f'{YELLOW}You:{RESET} ').strip()
        except (EOFError, KeyboardInterrupt):
            print('\nBye.')
            break

        if user_input.lower() in ['exit', 'quit']:
            print('Bye.')
            break

        if not user_input:
            continue

        reply = runtime.handle_one_turn(args.session, user_input)
        print(f'{GREEN}Assistant:{RESET} ', end='')
        pretty_print(reply)


if __name__ == '__main__':
    main()
