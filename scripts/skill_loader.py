from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


@dataclass
class Skill:
    name: str
    description: str
    instructions: str
    path: Path
    score: int = 0


class SkillLoader:
    def __init__(self, skills_dir: str | Path):
        self.skills_dir = Path(skills_dir)
        self.skills_dir.mkdir(parents=True, exist_ok=True)

    def list_skills(self) -> list[Skill]:
        skills: list[Skill] = []
        for skill_dir in sorted(self.skills_dir.iterdir()):
            if not skill_dir.is_dir():
                continue
            skill_file = skill_dir / 'SKILL.md'
            if not skill_file.exists():
                continue
            skills.append(self._read_skill(skill_file))
        return skills

    def get_skill(self, name: str) -> Skill | None:
        normalized = name.strip().lower()
        for skill in self.list_skills():
            if skill.name.lower() == normalized:
                return skill
        return None

    def match_skill(self, user_message: str) -> Skill | None:
        text = user_message.strip()
        if not text:
            return None

        explicit = self._extract_explicit_skill_name(text)
        if explicit:
            return self.get_skill(explicit)

        tokens = self._tokenize(text)
        if not tokens:
            return None

        best: Skill | None = None
        for skill in self.list_skills():
            haystack = f'{skill.name} {skill.description} {skill.instructions}'
            skill_tokens = self._tokenize(haystack)
            overlap = len(tokens & skill_tokens)
            if overlap <= 0:
                continue
            candidate = Skill(
                name=skill.name,
                description=skill.description,
                instructions=skill.instructions,
                path=skill.path,
                score=overlap,
            )
            if best is None or candidate.score > best.score:
                best = candidate
        return best


    def build_skill_prompt(self, skill: Skill | None) -> str:
        if not skill:
            return ''

        sections = [
            f'Skill name: {skill.name}',
            f'Skill description: {skill.description}',
            'Skill instructions:',
            skill.instructions,
        ]

        scripts_dir = skill.path.parent / 'scripts'
        if scripts_dir.exists() and scripts_dir.is_dir():
            script_sections: list[str] = []
            for script_file in sorted(p for p in scripts_dir.rglob('*') if p.is_file()):
                try:
                    content = script_file.read_text(encoding='utf-8')
                except UnicodeDecodeError:
                    try:
                        content = script_file.read_text(encoding='utf-8', errors='replace')
                    except Exception:
                        content = '<read failed>'
                except Exception:
                    content = '<read failed>'

                rel_path = script_file.relative_to(skill.path.parent).as_posix()
                script_sections.append(
                    f'File: {rel_path}\n'
                    f'```\n{content}\n```'
                )

            if script_sections:
                sections.append('Skill scripts:')
                sections.extend(script_sections)

        return '\n\n'.join(sections)


    def _read_skill(self, skill_file: Path) -> Skill:
        raw = skill_file.read_text(encoding='utf-8')
        name = skill_file.parent.name
        description = ''
        instructions = raw

        if raw.startswith('---'):
            parts = raw.split('---', 2)
            if len(parts) >= 3:
                frontmatter = parts[1]
                instructions = parts[2].strip()
                for line in frontmatter.splitlines():
                    if ':' not in line:
                        continue
                    key, value = line.split(':', 1)
                    key = key.strip().lower()
                    value = value.strip().strip('"').strip("'")
                    if key == 'name' and value:
                        name = value
                    elif key == 'description' and value:
                        description = value

        if not description:
            first_non_empty = next((line.strip() for line in instructions.splitlines() if line.strip()), '')
            description = first_non_empty[:200]

        return Skill(name=name, description=description, instructions=instructions, path=skill_file)

    def _extract_explicit_skill_name(self, text: str) -> str | None:
        patterns = [
            r'^/skill\s+use\s+([\w\-]+)',
            r'^/skill\s+([\w\-]+)',
            r'^skill:\s*([\w\-]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text.strip(), flags=re.IGNORECASE)
            if match:
                return match.group(1)
        return None

    def strip_explicit_skill_prefix(self, text: str) -> tuple[str | None, str]:
        source = text.strip()
        match = re.match(
            r'^/skill(?:\s+use)?\s+([\w\-]+)\s*(.*)$',
            source,
            flags=re.IGNORECASE
        )
        if not match:
            return None, text

        skill_name = match.group(1)
        remaining = (match.group(2) or '').strip()
        return skill_name, remaining

    def _tokenize(self, text: str) -> set[str]:
        lowered = text.lower()
        pieces = re.findall(r'[a-z0-9_\-\u3040-\u30ff\u3400-\u9fff]+', lowered)
        return {p for p in pieces if len(p) >= 2}
