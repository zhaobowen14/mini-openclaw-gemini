from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class Session:
    session_id: str
    messages: list[dict[str, Any]] = field(default_factory=list)
    meta: dict[str, Any] = field(default_factory=dict)


class SessionStore:
    def __init__(self, base_dir: str | Path):
        self.base_dir = Path(base_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def _path(self, session_id: str) -> Path:
        return self.base_dir / f"{session_id}.json"

    def load(self, session_id: str) -> Session:
        path = self._path(session_id)
        if not path.exists():
            return Session(session_id=session_id)
        data = json.loads(path.read_text(encoding="utf-8"))
        return Session(
            session_id=data["session_id"],
            messages=data.get("messages", []),
            meta=data.get("meta", {}),
        )

    def save(self, session: Session) -> None:
        path = self._path(session.session_id)
        data = {
            "session_id": session.session_id,
            "messages": session.messages,
            "meta": session.meta,
        }
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def append(self, session: Session, role: str, content: Any, **extra: Any) -> None:
        message = {"role": role, "content": content}
        if extra:
            message.update(extra)
        session.messages.append(message)

    def _get_session_path(self, session_id: str) -> Path:
        return self.base_dir / f"{session_id}.json"
    
