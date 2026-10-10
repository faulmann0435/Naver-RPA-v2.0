"""Downloads watcher: first run takes stock, later runs compare each finished export once."""
import os
import time

from worker.compare import CompareResult
from worker.watch_downloads import (
    FOLDER_KEY,
    SETTLE_SECONDS,
    check_once,
    configured_folder,
    load_state,
)


class Recorder:
    def __init__(self):
        self.lines = []

    def send(self, text):
        self.lines.append(text)


def _export(folder, name, age=SETTLE_SECONDS + 1):
    path = folder / name
    path.write_bytes(b"data")
    past = time.time() - age
    os.utime(path, (past, past))
    return path


def _run(tmp_path, folder, notifier, compared):
    def fake_compare(path):
        compared.append(path.name)
        return CompareResult(path.name, excel_rows=2, api_rows=2)
    return check_once(folder, notifier, fake_compare, tmp_path / "state.json", tmp_path / "out")


def test_first_run_skips_existing_exports(tmp_path):
    folder = tmp_path / "Downloads"
    folder.mkdir()
    _export(folder, "스마트스토어_old.xlsx")
    notifier, compared = Recorder(), []
    assert _run(tmp_path, folder, notifier, compared) == []
    assert compared == [] and "건너뜀" in notifier.lines[0]
    assert "스마트스토어_old.xlsx" in load_state(folder, tmp_path / "state.json")


def test_new_export_is_compared_once_and_reported(tmp_path):
    folder = tmp_path / "Downloads"
    folder.mkdir()
    notifier, compared = Recorder(), []
    _run(tmp_path, folder, notifier, compared)
    _export(folder, "스마트스토어_new.xlsx")
    _run(tmp_path, folder, notifier, compared)
    _run(tmp_path, folder, notifier, compared)
    assert compared == ["스마트스토어_new.xlsx"]
    assert notifier.lines[-1].startswith("✅ 스마트스토어_new.xlsx")
    assert len(list((tmp_path / "out").glob("*_스마트스토어_new.txt"))) == 1


def test_file_still_downloading_waits(tmp_path):
    folder = tmp_path / "Downloads"
    folder.mkdir()
    notifier, compared = Recorder(), []
    _run(tmp_path, folder, notifier, compared)
    _export(folder, "스마트스토어_fresh.xlsx", age=0)
    _run(tmp_path, folder, notifier, compared)
    assert compared == []


def test_other_files_are_ignored(tmp_path):
    folder = tmp_path / "Downloads"
    folder.mkdir()
    notifier, compared = Recorder(), []
    _run(tmp_path, folder, notifier, compared)
    _export(folder, "사진.xlsx")
    _export(folder, "스마트스토어_x.xlsx.crdownload")
    _run(tmp_path, folder, notifier, compared)
    assert compared == []


def test_a_failing_export_is_reported_and_watching_goes_on(tmp_path):
    folder = tmp_path / "Downloads"
    folder.mkdir()
    notifier = Recorder()
    check_once(folder, notifier, lambda p: CompareResult(p.name), tmp_path / "s.json", tmp_path / "out")
    _export(folder, "스마트스토어_bad.xlsx")
    _export(folder, "스마트스토어_good.xlsx", age=SETTLE_SECONDS + 2)

    def flaky(path):
        if "bad" in path.name:
            raise ValueError("비밀번호가 올바르지 않습니다.")
        return CompareResult(path.name, excel_rows=1, api_rows=1)

    results = check_once(folder, notifier, flaky, tmp_path / "s.json", tmp_path / "out")
    assert [r.matched for r in results] == [True, False]
    assert notifier.lines[-1].startswith("⚠ 스마트스토어_bad.xlsx: 대조 못 함")


def test_summary_flags_unexpected_differences():
    result = CompareResult("f.xlsx", 2, 2, column_diffs={"주문상태": [], "수량": []}, file_problems=("속초: 내용이 다름",))
    assert not result.matched
    assert result.summary().startswith("❌ f.xlsx")


def test_switching_folder_takes_stock_again(tmp_path):
    first, second = tmp_path / "A", tmp_path / "B"
    first.mkdir()
    second.mkdir()
    notifier, compared = Recorder(), []
    _run(tmp_path, first, notifier, compared)
    _export(second, "스마트스토어_old_in_b.xlsx")
    _run(tmp_path, second, notifier, compared)
    assert compared == [] and "건너뜀" in notifier.lines[-1]


def test_folder_comes_from_environment(monkeypatch, tmp_path):
    monkeypatch.setenv(FOLDER_KEY, str(tmp_path))
    assert configured_folder() == tmp_path
