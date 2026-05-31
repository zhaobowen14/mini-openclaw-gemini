from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

JST = timezone(timedelta(hours=9))

@dataclass
class CronJob:
    id: str
    message: str
    session: str = 'default'
    interval_seconds: int = 3600
    enabled: bool = True
    next_run_at: str | None = None
    last_run_at: str | None = None
    schedule_type: str = 'interval'
    time_of_day: str | None = None
    weekday: int | None = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> 'CronJob':
        return cls(
            id=str(data['id']),
            message=str(data['message']),
            session=str(data.get('session', 'default')),
            interval_seconds=int(data.get('interval_seconds', 3600)),
            enabled=bool(data.get('enabled', True)),
            next_run_at=data.get('next_run_at'),
            last_run_at=data.get('last_run_at'),
            schedule_type=str(data.get('schedule_type', 'interval')),
            time_of_day=data.get('time_of_day'),
            weekday=data.get('weekday'),
        )


    def _to_jst_str(self, iso_z: str | None) -> str | None:
        if not iso_z:
            return None
        dt = datetime.fromisoformat(iso_z.replace('Z', '+00:00'))
        return dt.astimezone(JST).strftime('%Y-%m-%d %H:%M:%S')
    
    def to_dict(self) -> dict[str, Any]:
        return {
            'id': self.id,
            'message': self.message,
            'session': self.session,
            'interval_seconds': self.interval_seconds,
            'enabled': self.enabled,
            'next_run_at': self.next_run_at,
            'last_run_at': self.last_run_at,
            'next_run_at_jp':self._to_jst_str(self.next_run_at),
            'last_run_at_jp': self._to_jst_str(self.last_run_at),
            'schedule_type': self.schedule_type,
            'time_of_day': self.time_of_day,
            'weekday': self.weekday,
        }


class CronManager:
    def __init__(self, cron_file: str | Path, runtime_script: str | Path, workspace: str | Path):
        self.cron_file = Path(cron_file).resolve()
        self.runtime_script = Path(runtime_script)
        self.workspace = Path(workspace).resolve()
        if not self.runtime_script.is_absolute():
            self.runtime_script = (self.workspace.parent / self.runtime_script).resolve()
        else:
            self.runtime_script = self.runtime_script.resolve()
        self.cron_file.parent.mkdir(parents=True, exist_ok=True)
        if not self.cron_file.exists():
            self.save_jobs([])

    def load_jobs(self) -> list[CronJob]:
        data = json.loads(self.cron_file.read_text(encoding='utf-8-sig'))
        return [CronJob.from_dict(item) for item in data]

    def save_jobs(self, jobs: list[CronJob]) -> None:
        payload = [job.to_dict() for job in jobs]
        self.cron_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')

    def add_job(
        self,
        job_id: str,
        message: str,
        interval_seconds: int,
        session: str = 'default',
        schedule_type: str = 'interval',
        time_of_day: str | None = None,
        weekday: int | None = None,
    ) -> CronJob:
        jobs = self.load_jobs()
        if any(job.id == job_id for job in jobs):
            raise ValueError(f'cron job already exists: {job_id}')

        now = self._utc_now()
        job = CronJob(
            id=job_id,
            message=message,
            session=session,
            interval_seconds=interval_seconds,
            enabled=True,
            schedule_type=schedule_type,
            time_of_day=time_of_day,
            weekday=weekday,
        )
        job.next_run_at = self._calculate_next_run(job, now)
        jobs.append(job)
        self.save_jobs(jobs)
        return job

    def remove_job(self, job_id: str) -> bool:
        jobs = self.load_jobs()
        filtered = [job for job in jobs if job.id != job_id]
        changed = len(filtered) != len(jobs)
        if changed:
            self.save_jobs(filtered)
        return changed

    def set_enabled(self, job_id: str, enabled: bool) -> bool:
        jobs = self.load_jobs()
        changed = False
        now = self._utc_now()
        for job in jobs:
            if job.id != job_id:
                continue
            job.enabled = enabled
            if enabled:
                job.next_run_at = self._calculate_next_run(job, now)
            changed = True
        if changed:
            self.save_jobs(jobs)
        return changed

    def run_pending(self) -> list[dict[str, Any]]:
        jobs = self.load_jobs()
        now = self._utc_now()
        results: list[dict[str, Any]] = []
        changed = False

        for job in jobs:
            if not job.enabled:
                continue
            if not self._is_due(job, now):
                continue

            result = self._run_job(job)
            results.append(result)
            changed = True
            job.last_run_at = self._to_iso(now.timestamp())
            job.next_run_at = self._calculate_next_run(job, now)


        if changed:
            self.save_jobs(jobs)
        return results

    def run_loop(self, poll_seconds: int = 30) -> None:
        print(f'Cron loop started. poll_seconds={poll_seconds}')
        while True:
            results = self.run_pending()
            for result in results:
                print(json.dumps(result, ensure_ascii=False))
            time.sleep(max(1, poll_seconds))

    def _run_job(self, job: CronJob) -> dict[str, Any]:
        with tempfile.NamedTemporaryFile('w', encoding='utf-8', suffix='.txt', delete=False, dir=str(self.workspace)) as tmp:
            tmp.write(job.message)
            message_file = Path(tmp.name)

        try:
            command = [
                'python',
                str(self.runtime_script),
                '--session',
                job.session,
                '--message-file',
                str(message_file),
            ]
            completed = subprocess.run(
                command,
                cwd=str(self.workspace),
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                timeout=max(60, job.interval_seconds),
            )
            return {
                'id': job.id,
                'ok': completed.returncode == 0,
                'returncode': completed.returncode,
                'stdout': completed.stdout[-4000:],
                'stderr': completed.stderr[-4000:],
            }
        finally:
            try:
                message_file.unlink(missing_ok=True)
            except Exception:
                pass

    def _is_due(self, job: CronJob, now: datetime) -> bool:
        if not job.next_run_at:
            return True
        try:
            due_at = datetime.fromisoformat(job.next_run_at.replace('Z', '+00:00'))
        except ValueError:
            return True
        return due_at <= now

    def _calculate_next_run(self, job: CronJob, now: datetime) -> str:
        if job.schedule_type == 'daily' and job.time_of_day:
            hour, minute = self._parse_time_of_day(job.time_of_day)
            candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            if candidate <= now:
                candidate += timedelta(days=1)
            return self._to_iso(candidate.timestamp())

        if job.schedule_type == 'weekly' and job.time_of_day and job.weekday is not None:
            hour, minute = self._parse_time_of_day(job.time_of_day)
            candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            delta_days = (int(job.weekday) - candidate.weekday()) % 7
            candidate += timedelta(days=delta_days)
            if candidate <= now:
                candidate += timedelta(days=7)
            return self._to_iso(candidate.timestamp())

        return self._to_iso(now.timestamp() + max(1, job.interval_seconds))

    def _parse_time_of_day(self, value: str) -> tuple[int, int]:
        hour_s, minute_s = value.split(':', 1)
        return int(hour_s), int(minute_s)

    def _utc_now(self) -> datetime:
        return datetime.now(timezone.utc)

    def _to_iso(self, timestamp: float) -> str:
        return datetime.fromtimestamp(timestamp, tz=timezone.utc).isoformat().replace('+00:00', 'Z')


class NaturalLanguageCronParser:
    WEEKDAYS = {
        '月': 0, '月曜': 0, '月曜日': 0,
        '火': 1, '火曜': 1, '火曜日': 1,
        '水': 2, '水曜': 2, '水曜日': 2,
        '木': 3, '木曜': 3, '木曜日': 3,
        '金': 4, '金曜': 4, '金曜日': 4,
        '土': 5, '土曜': 5, '土曜日': 5,
        '日': 6, '日曜': 6, '日曜日': 6,
    }

    def parse(self, text: str, session: str = 'default') -> dict[str, Any] | None:
        source = text.strip()
        if not source:
            return None

        lowered = source.lower()

        if 'cron' in lowered and any(word in source for word in ['一覧', '見せて', '表示', '教えて']):
            return {'action': 'list'}

        m = re.search(r'([\w\-]+)\s*を\s*(止めて|停止|無効)', source)
        if m:
            return {'action': 'disable', 'id': m.group(1)}

        m = re.search(r'([\w\-]+)\s*を\s*(再開|有効)', source)
        if m:
            return {'action': 'enable', 'id': m.group(1)}

        if '毎朝' in source or '毎日' in source:
            time_of_day = self._extract_time(source) or '09:00'
            message = self._extract_message(source)
            if message:
                return {
                    'action': 'add',
                    'schedule_type': 'daily',
                    'time_of_day': time_of_day,
                    'interval_seconds': 24 * 3600,
                    'message': message,
                    'id': self._make_id('daily', message),
                    'session': session,
                }

        weekly_match = re.search(r'毎週\s*([月火水木金土日](?:曜(?:日)?)?)', source)
        if weekly_match:
            weekday_label = weekly_match.group(1)
            weekday = self.WEEKDAYS.get(weekday_label)
            time_of_day = self._extract_time(source) or '09:00'
            message = self._extract_message(source)
            if message and weekday is not None:
                return {
                    'action': 'add',
                    'schedule_type': 'weekly',
                    'weekday': weekday,
                    'time_of_day': time_of_day,
                    'interval_seconds': 7 * 24 * 3600,
                    'message': message,
                    'id': self._make_id(f'weekly-{weekday}', message),
                    'session': session,
                }

        m = re.search(r'(\d+)\s*分おきに', source)
        if m:
            minutes = int(m.group(1))
            message = self._extract_message(source)
            if message:
                return {
                    'action': 'add',
                    'schedule_type': 'interval',
                    'interval_seconds': minutes * 60,
                    'message': message,
                    'id': self._make_id(f'every-{minutes}m', message),
                    'session': session,
                }

        m = re.search(r'(\d+)\s*時間おきに', source)
        if m:
            hours = int(m.group(1))
            message = self._extract_message(source)
            if message:
                return {
                    'action': 'add',
                    'schedule_type': 'interval',
                    'interval_seconds': hours * 3600,
                    'message': message,
                    'id': self._make_id(f'every-{hours}h', message),
                    'session': session,
                }

        return None

    def _extract_time(self, text: str) -> str | None:
        m = re.search(r'([0-2]?\d)[:時]([0-5]\d)?', text)
        if not m:
            return None
        hour = int(m.group(1))
        minute = int(m.group(2) or '00')
        if hour > 23 or minute > 59:
            return None
        return f'{hour:02d}:{minute:02d}'

    def _extract_message(self, text: str) -> str:
        cleaned = text
        cleaned = re.sub(r'毎朝\s*[0-2]?\d(?::[0-5]\d)?時?', '', cleaned)
        cleaned = re.sub(r'毎日\s*[0-2]?\d(?::[0-5]\d)?時?', '', cleaned)
        cleaned = re.sub(r'毎週\s*[月火水木金土日](?:曜(?:日)?)?\s*[0-2]?\d(?::[0-5]\d)?時?', '', cleaned)
        cleaned = re.sub(r'\d+\s*分おきに', '', cleaned)
        cleaned = re.sub(r'\d+\s*時間おきに', '', cleaned)
        cleaned = cleaned.replace('cron', '')
        cleaned = re.sub(r'^(に|で|を)+', '', cleaned)
        cleaned = re.sub(r'(して|してね|してください|頼む|お願いします)$', '', cleaned).strip()
        return cleaned.strip(' 　、。')

    def _make_id(self, prefix: str, message: str) -> str:
        slug = re.sub(r'[^a-zA-Z0-9\-]+', '-', message.lower()).strip('-')
        if not slug:
            slug = 'job'
        return f'{prefix}-{slug[:24]}'


def main() -> None:
    import argparse

    dir_path = Path(__file__).resolve().parents[1]
    print(dir_path)
    parser = argparse.ArgumentParser()
    parser.add_argument('--cron-file', default=f'{dir_path}/cron/jobs.json')
    parser.add_argument('--runtime-script', default=f'{dir_path}/scripts/runtime.py')
    parser.add_argument('--workspace', default=f'{dir_path}/cron')

    subparsers = parser.add_subparsers(dest='command', required=True)

    add_parser = subparsers.add_parser('add')
    add_parser.add_argument('--id', required=True)
    add_parser.add_argument('--message', required=True)
    add_parser.add_argument('--interval-seconds', type=int, required=True)
    add_parser.add_argument('--session', default='default')
    add_parser.add_argument('--schedule-type', default='interval')
    add_parser.add_argument('--time-of-day', default=None)
    add_parser.add_argument('--weekday', type=int, default=None)

    remove_parser = subparsers.add_parser('remove')
    remove_parser.add_argument('--id', required=True)

    enable_parser = subparsers.add_parser('enable')
    enable_parser.add_argument('--id', required=True)

    disable_parser = subparsers.add_parser('disable')
    disable_parser.add_argument('--id', required=True)

    subparsers.add_parser('list')
    subparsers.add_parser('run-once')

    loop_parser = subparsers.add_parser('loop')
    loop_parser.add_argument('--poll-seconds', type=int, default=30)

    parse_parser = subparsers.add_parser('parse-text')
    parse_parser.add_argument('--text', required=True)
    parse_parser.add_argument('--session', default='default')

    args = parser.parse_args()

    manager = CronManager(args.cron_file, args.runtime_script, args.workspace)

    if args.command == 'add':
        job = manager.add_job(
            args.id,
            args.message,
            args.interval_seconds,
            session=args.session,
            schedule_type=args.schedule_type,
            time_of_day=args.time_of_day,
            weekday=args.weekday,
        )
        print(json.dumps(job.to_dict(), ensure_ascii=False, indent=2))
        return

    if args.command == 'remove':
        removed = manager.remove_job(args.id)
        print(json.dumps({'removed': removed, 'id': args.id}, ensure_ascii=False))
        return

    if args.command == 'enable':
        changed = manager.set_enabled(args.id, True)
        print(json.dumps({'enabled': changed, 'id': args.id}, ensure_ascii=False))
        return

    if args.command == 'disable':
        changed = manager.set_enabled(args.id, False)
        print(json.dumps({'disabled': changed, 'id': args.id}, ensure_ascii=False))
        return

    if args.command == 'list':
        print(json.dumps([job.to_dict() for job in manager.load_jobs()], ensure_ascii=False, indent=2))

        # jobs = manager.load_jobs()
        # result = []
        # for job in jobs:
        #     d = job.to_dict()
        #     d['next_run_at_jst'] = manager._to_jst_str(job.next_run_at)
        #     d['last_run_at_jst'] = manager._to_jst_str(job.last_run_at)
        #     result.append(d)

        # print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    if args.command == 'run-once':
        print(json.dumps(manager.run_pending(), ensure_ascii=False, indent=2))
        return

    if args.command == 'loop':
        manager.run_loop(poll_seconds=args.poll_seconds)
        return

    if args.command == 'parse-text':
        parsed = NaturalLanguageCronParser().parse(args.text, session=args.session)
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
        return


if __name__ == '__main__':
    main()
