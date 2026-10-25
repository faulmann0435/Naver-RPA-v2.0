"""Business-hours settings of the CS alerts: parsing, fallbacks, boundaries (pure, fake times)."""
from datetime import datetime, time, timezone

import pytest

from worker import cs_hours
from worker.cs_hours import build_schedule, is_work_time, load_schedule, parse_days, parse_hours
from worker.naver_adapter import KST

DEFAULT, _ = build_schedule("", "")


def at(day, hour, minute=0):
    return datetime(2026, 10, day, hour, minute, tzinfo=KST)  # 2026-10-12 is a Monday


def test_defaults():
    schedule, warnings = build_schedule("", "")
    assert (schedule.start, schedule.end) == (time(9), time(18, 10)) and schedule.days == {0, 1, 2, 3, 4}
    assert warnings == [] and schedule.describe() == "월화수목금 09:00-18:10"


def test_custom_values():
    schedule, warnings = build_schedule(" 8:30 - 17:00 ", "월 수 토")
    assert (schedule.start, schedule.end, schedule.days) == (time(8, 30), time(17), frozenset({0, 2, 5}))
    assert warnings == []


@pytest.mark.parametrize("text", ["", "09:00", "9-18", "18:00-09:00", "09:00-09:00", "25:00-26:00",
                                  "09:61-10:00", "a:b-c:d"])
def test_bad_hours_raise(text):
    with pytest.raises(ValueError):
        parse_hours(text)


@pytest.mark.parametrize("text", ["", "   ", "mon", "월화x", "월요일"])
def test_bad_days_raise(text):
    with pytest.raises(ValueError):
        parse_days(text)


def test_invalid_values_fall_back_to_defaults_with_one_warning_each():
    schedule, warnings = build_schedule("late", "xyz")
    assert schedule == DEFAULT and len(warnings) == 2
    assert cs_hours.HOURS_KEY in warnings[0] and cs_hours.DAYS_KEY in warnings[1]
    schedule, warnings = build_schedule("10:00-12:00", "xyz")  # only the bad one falls back
    assert (schedule.start, schedule.end) == (time(10), time(12))
    assert schedule.days == DEFAULT.days and len(warnings) == 1


@pytest.mark.parametrize("moment, expected", [
    (at(12, 8, 59), False), (at(12, 9, 0), True), (at(12, 17, 59), True), (at(12, 18, 9), True), (at(12, 18, 10), False),
    (at(16, 12), True),   # Friday
    (at(17, 12), False),  # Saturday
    (at(18, 12), False),  # Sunday
])
def test_is_work_time_boundaries_and_weekend(moment, expected):
    assert is_work_time(moment, DEFAULT) is expected


def test_is_work_time_uses_korean_time_for_other_zones():
    assert is_work_time(datetime(2026, 10, 12, 0, 30, tzinfo=timezone.utc), DEFAULT)  # Monday 09:30 in Seoul
    assert not is_work_time(datetime(2026, 10, 11, 23, 59, tzinfo=timezone.utc), DEFAULT)  # Monday 08:59


def test_load_schedule_reads_environment_then_env_file(monkeypatch):
    monkeypatch.setattr(cs_hours, "read_env_file", lambda path: {cs_hours.HOURS_KEY: "07:00-08:00"})
    monkeypatch.delenv(cs_hours.HOURS_KEY, raising=False)
    monkeypatch.setenv(cs_hours.DAYS_KEY, "토일")
    schedule, warnings = load_schedule()
    assert (schedule.start, schedule.days) == (time(7), frozenset({5, 6})) and warnings == []
    monkeypatch.setenv(cs_hours.HOURS_KEY, "10:00-11:00")  # the environment wins
    assert load_schedule()[0].start == time(10)
