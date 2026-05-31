from __future__ import annotations

from importlib.resources import path
import os
from os import path
import re
import json
from typing import Optional
import socket
import subprocess
from html import unescape
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader


class ToolRunner:
    def __init__(self, workspace: str | Path):
        self.workspace = Path(workspace)

    def run(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        print(f"[TOOL CALL] name={name} arguments={arguments}")

        try:
            if name == 'read':
                return self._read(arguments)
            if name == 'write':
                return self._write(arguments)
            if name == 'exec':
                return self._exec(arguments)
            if name == 'fetch_url':
                return self._fetch_url(arguments)
            if name == 'search_web':
                return self._search_web(arguments)
            return {'ok': False, 'error': f'unknown tool: {name}'}
        except Exception as e:
            return {'ok': False, 'error': str(e)}

    def _base_dir(self) -> Path:
        base = self.workspace 
        # base.mkdir(parents=True, exist_ok=True)
        return base.resolve()

    def _resolve_in_outputs(self, path: str) -> Path:
        if not isinstance(path, str) or not path.strip():
            raise ValueError('path must be a non-empty string')

        base = self.workspace.resolve()
        fallback = (base / "outputs" / Path(path).name).resolve()

        return fallback


    def _resolve_allowed_path(self, path: str) -> Path:
        if not isinstance(path, str) or not path.strip():
            raise ValueError('path must be a non-empty string')

        p = (self.workspace / path).resolve()

        print("workspace =", self.workspace.resolve())
        print("requested path =", path)
        print("resolved path =", p)

        allowed_dirs = [
            (self.workspace / "outputs").resolve(),
            (self.workspace / "skills").resolve(),
            (self.workspace / "cron").resolve(),
        ]
        
        for base in allowed_dirs:
            if p == base or base in p.parents:
                return p

        # フォールバック処理
        fallback = (self.workspace / "outputs" / Path(path).name).resolve()
        print(f"[WARN] path not allowed, fallback to: {fallback}")
        return fallback 

    def _resolve_path(self, path: str) -> Path:
        if not isinstance(path, str) or not path.strip():
            raise ValueError('path must be a non-empty string')

        p = (self.workspace / path).resolve()


        # ✔ 存在するならOK
        if p.exists():
            return p

        # ✔ フォールバック
        fallback = (self.workspace / "outputs" / Path(path).name).resolve()

        if fallback.exists():
            print(f"[WARN] fallback to: {fallback}")
            return fallback

        # ✔ fallbackも無ければエラー（ここ重要）
        raise FileNotFoundError(f'file not found: {fallback}')


    def _resolve_read_path(self, path: str) -> Path:
        if not isinstance(path, str) or not path.strip():
            raise ValueError('path must be a non-empty string')

        p = (self.workspace / path).resolve()

        print("workspace =", self.workspace.resolve())
        print("requested path =", path)
        print("resolved path =", p)

        allowed_dirs = [
            (self.workspace / "inputs").resolve(),
            (self.workspace / "outputs").resolve(),
            (self.workspace / "skills").resolve(),
            (self.workspace / "cron").resolve(),
        ]

        allowed = any(p == base or base in p.parents for base in allowed_dirs)
        allowed = True

        # ✔ 許可内かつ存在するならOK
        if allowed and p.exists():
            return p

        # ✔ フォールバック
        fallback = (self.workspace / "outputs" / Path(path).name).resolve()

        if fallback.exists():
            print(f"[WARN] fallback to: {fallback}")
            return fallback

        # ✔ fallbackも無ければエラー
        raise FileNotFoundError(f'file not found: {path}')

    def _read_pdf(self, path: Path) -> str:
        reader = PdfReader(str(path))
        texts: list[str] = []

        for page in reader.pages:
            text = page.extract_text() or ''
            if text:
                texts.append(text)

        return '\n\n'.join(texts)

    def _read(self, arguments: dict[str, Any]) -> dict[str, Any]:
        path = self._resolve_read_path(arguments['path'])
        print(f'[READ] path={path}')

        if not path.exists():
            return {'ok': False, 'error': f'file not found: {path}'}

        if path.suffix.lower() == '.pdf':
            text = self._read_pdf(path)
        else:
            text = path.read_text(encoding='utf-8')

        print(f'[READ SUCCESS] {len(text)} chars')
        return {'ok': True, 'path': str(path), 'content': text[:20000]}

    def _write(self, arguments: dict[str, Any]) -> dict[str, Any]:
        path = self._resolve_in_outputs(arguments['path'])

        content = arguments.get('content', '')
        if not isinstance(content, str):
            raise ValueError('content must be a string')

        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding='utf-8')
        return {'ok': True, 'path': str(path)}

    def _exec(self, arguments: dict[str, Any]) -> dict[str, Any]:
        command = arguments['command']
        if not isinstance(command, str) or not command.strip():
            raise ValueError('command must be a non-empty string')

        cwd_arg = arguments.get('cwd')
        if isinstance(cwd_arg, str) and cwd_arg.strip():
            cwd = self._resolve_in_outputs(cwd_arg)
        else:
            cwd = self._base_dir()

        completed = subprocess.run(
            command,
            cwd=str(cwd),
            shell=True,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            timeout=int(arguments.get('timeout', 30)),
        )
        return {
            'ok': completed.returncode == 0,
            'returncode': completed.returncode,
            'stdout': completed.stdout[-20000:],
            'stderr': completed.stderr[-20000:],
        }

    def _fetch_url(self, arguments: dict[str, Any]) -> dict[str, Any]:
        url = arguments.get('url', '')
        max_chars = int(arguments.get('max_chars', 20000))
        self._validate_public_url(url)

        response = requests.get(
            url,
            timeout=20,
            allow_redirects=True,
            headers={'User-Agent': 'mini-openclaw-gemini/0.1'},
        )
        response.raise_for_status()

        content_type = response.headers.get('Content-Type', '').lower()
        if 'text/html' in content_type:
            title, content = self._extract_html_text(response.text)
        elif 'text/' in content_type or 'json' in content_type or not content_type:
            title = None
            content = response.text
        else:
            return {
                'ok': False,
                'url': url,
                'status': response.status_code,
                'content_type': content_type,
                'error': f'unsupported content type: {content_type}',
            }

        truncated = len(content) > max_chars
        return {
            'ok': True,
            'url': response.url,
            'status': response.status_code,
            'content_type': content_type,
            'title': title,
            'content': content[:max_chars],
            'truncated': truncated,
        }

    def _search_web(self, arguments: dict[str, Any]) -> dict[str, Any]:
        query = arguments.get('query', '')
        if not isinstance(query, str) or not query.strip():
            raise ValueError('query must be a non-empty string')

        api_key = self.get_api_key("brave")
        if not api_key:
            return {
                'ok': False,
                'error': 'API key not found. Set api.json.',
            }

        count = int(arguments.get('count', 5))
        count = max(1, min(count, 10))

        params = {
            'q': query,
            'count': count,
        }
        if arguments.get('search_lang'):
            params['search_lang'] = arguments.get('search_lang')
        if arguments.get('country'):
            params['country'] = arguments.get('country')

        response = requests.get(
            'https://api.search.brave.com/res/v1/web/search',
            headers={
                'Accept': 'application/json',
                'X-Subscription-Token': api_key,
                'User-Agent': 'mini-openclaw-gemini/0.1',
            },
            params=params,
            timeout=20,
        )
        if response.status_code >= 400:
            return {
                'ok': False,
                'status': response.status_code,
                'error': f'Brave API error: {response.status_code}',
                'body': response.text[:2000],
            }
        raw = response.json()
        web = raw.get('web', {})
        items = web.get('results', [])

        results = []
        for item in items[:count]:
            results.append(
                {
                    'title': item.get('title'),
                    'url': item.get('url'),
                    'snippet': item.get('description') or item.get('snippet'),
                }
            )

        return {
            'ok': True,
            'query': query,
            'count': count,
            'results': results,
        }

    def get_api_key(
        self,
        key_name: str,
        filename: str = "api.json",
    ) -> Optional[str]:
        json_file = (self.workspace / filename).resolve()

        if not json_file.exists():
            return None

        try:
            data = json.loads(json_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return None

        value = data.get(key_name)
        if value is None:
            return None

        return str(value).strip()
    
    def _extract_html_text(self, html_text: str) -> tuple[str | None, str]:
        soup = BeautifulSoup(html_text, 'html.parser')

        for tag in soup(['script', 'style', 'noscript', 'svg', 'footer', 'nav', 'header', 'aside']):
            tag.decompose()

        title = None
        if soup.title and soup.title.string:
            title = soup.title.string.strip()

        main = soup.find('main') or soup.find('article') or soup.body or soup
        text = main.get_text('\n', strip=True)
        text = unescape(text)
        text = re.sub(r'\n{3,}', '\n\n', text)
        return title, text

    def _validate_public_url(self, url: str) -> None:
        if not isinstance(url, str) or not url.strip():
            raise ValueError('url must be a non-empty string')

        parsed = urlparse(url)
        if parsed.scheme not in {'http', 'https'}:
            raise ValueError('only http/https URLs are allowed')

        host = parsed.hostname
        if not host:
            raise ValueError('invalid URL host')

        blocked_hosts = {'localhost', '127.0.0.1', '::1'}
        if host.lower() in blocked_hosts:
            raise ValueError('local URLs are not allowed')

        try:
            addresses = socket.getaddrinfo(host, None)
        except socket.gaierror:
            return

        for entry in addresses:
            ip = entry[4][0]
            if ip.startswith('10.') or ip.startswith('192.168.') or ip.startswith('172.16.') or ip.startswith('172.17.') or ip.startswith('172.18.') or ip.startswith('172.19.') or ip.startswith('172.2') or ip.startswith('172.30.') or ip.startswith('172.31.'):
                raise ValueError('private network URLs are not allowed')
