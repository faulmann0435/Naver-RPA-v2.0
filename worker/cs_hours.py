"""Business hours of the CS alerts (the owner does not read Telegram outside them).

.env keys (read every cycle): CS_WORK_HOURS (default "09:00-18:10", start inclusive, end exclusive)
and CS_WORK_DAYS (default "월화수목금", Korean weekday letters, Monday = 월). A bad value is reported
as a warning and replaced by its default; it never raises.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from datetime import datetime, time

from worker.naver_adapter import KST
from worker.naver_client import ENV_FILE, read_env_file

HOURS_KEY, DAYS_KEY = "CS_WORK_HOURS", "CS_WORK_DAYS"
DEFAULT_HOURS, DEFAULT_DAYS = "09:00-18:10", "월화수목금"  # 18:10: the last 10-minute check still catches questions sent before 18:00
WEEKDAYS = "월화수목금토일"  # index = datetime.weekday()
_HOURS = re.compile(r"^\s*(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})\s*$")


@dataclass(frozen=True)
class WorkSchedule:
    start: time
    end: time
    days: frozenset[int]  # datetime.weekday() values

    def describe(self) -> str:
        letters = "".join(WEEKDAYS[d] for d in sorted(self.days))
        return f"{letters} {self.start:%H:%M}-{self.end:%H:%M}"


def parse_hours(text: str) -> tuple[time, time]:
    """("09:00-18:00") -> (time(9), time(18)); ValueError unless start < end on the same day."""
    match = _HOURS.match(text or "")
    if not match:
        raise ValueError(f"시간 형식이 올바르지 않습니다: {text!r}")
    sh, sm, eh, em = (int(g) for g in match.groups())
    if sh > 23 or eh > 23 or sm > 59 or em > 59:
        raise ValueError(f"시각 범위를 벗어났습니다: {text!r}")
    start, end = time(sh, sm), time(eh, em)
    if start >= end:
        raise ValueError(f"시작이 끝보다 늦거나 같습니다: {text!r}")
    return start, end


def parse_days(text: str) -> frozenset[int]:
    """("월화수목금") -> {0..4}; ValueError for an empty text or any letter that is not a weekday."""
    letters = "".join((text or "").split())
    if not letters or any(c not in WEEKDAYS for c in letters):
        raise ValueError(f"요일 형식이 올바르지 않습니다: {text!r}")
    return frozenset(WEEKDAYS.index(c) for c in letters)


def build_schedule(hours_text: str, days_text: str) -> tuple[WorkSchedule, list[str]]:
    """(schedule, warnings): an empty value means the default; a bad one falls back to it with a warning."""
    warnings: list[str] = []
    try:
        start, end = parse_hours(hours_text or DEFAULT_HOURS)
    except ValueError as err:
        warnings.append(f"{HOURS_KEY} 값을 쓸 수 없어 기본값 {DEFAULT_HOURS}을 씁니다: {err}")
        start, end = parse_hours(DEFAULT_HOURS)
    try:
        days = parse_days(days_text or DEFAULT_DAYS)
    except ValueError as err:
        warnings.append(f"{DAYS_KEY} 값을 쓸 수 없어 기본값 {DEFAULT_DAYS}을 씁니다: {err}")
        days = parse_days(DEFAULT_DAYS)
    return WorkSchedule(start, end, days), warnings


def _setting(key: str) -> str:
    return (os.environ.get(key) or read_env_file(ENV_FILE).get(key, "")).strip()


def load_schedule() -> tuple[WorkSchedule, list[str]]:
    """The schedule from the environment or .env (read on every call) and its warnings."""
    return build_schedule(_setting(HOURS_KEY), _setting(DAYS_KEY))


def is_work_time(now: datetime, schedule: WorkSchedule) -> bool:
    """True on a work day at or after start and before end (Korean time)."""
    local = now.astimezone(KST)
    return local.weekday() in schedule.days and schedule.start <= local.time().replace(tzinfo=None) < schedule.end
