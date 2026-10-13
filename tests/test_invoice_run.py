"""Invoice run for the bot: matching against API orders, pending state, ledger, dispatch."""
from datetime import datetime

import pytest

from tests.invoice_helpers import api_item, hanjin_row, hanjin_xlsx
from worker.dispatch_client import DispatchFailure, DispatchResult, dispatch_or_simulate
from worker.invoice_run import (
    add_pending,
    invoice_summary,
    mark_registered,
    prepare_invoices,
    read_ledger,
    read_pending,
    register,
    remove_pending,
)
from worker.naver_adapter import KST

WHEN = datetime(2026, 10, 13, 9, 5, tzinfo=KST)


def _prepare(tmp_path, items, *hanjin):
    return prepare_invoices(hanjin_xlsx(*hanjin), fetch=lambda: items, ledger=tmp_path / "ledger.json", when=WHEN)


def test_prepare_matches_api_orders_and_names_the_files(tmp_path):
    run = _prepare(tmp_path, [api_item("1"), api_item("2", package="91")], hanjin_row(1, "111"))
    (a,) = run.result.matched()
    assert a.product_order_ids == ("1", "2") and a.tracking == "111"
    assert [f.filename for f in run.files] == ["송장업로드_일치_1013_0905.xlsx"]


def test_second_file_appears_when_checks_are_includable(tmp_path):
    run = _prepare(tmp_path, [api_item("1"), api_item("2", "김철수")],
                   hanjin_row(1, "111"), hanjin_row(2, "222", "김철수", **{"받으시는 분 전화": "01099999999"}))
    assert [f.filename for f in run.files] == ["송장업로드_일치_1013_0905.xlsx", "송장업로드_확인포함_1013_0905.xlsx"]


def test_claims_and_already_registered_orders_are_left_out(tmp_path):
    ledger = tmp_path / "ledger.json"
    mark_registered({"1": "111"}, WHEN, ledger)
    items = [api_item("1"), api_item("2", "김철수", status="CANCELED")]
    run = prepare_invoices(hanjin_xlsx(hanjin_row(1, "111"), hanjin_row(2, "222", "김철수")),
                           fetch=lambda: items, ledger=ledger, when=WHEN)
    assert run.result.assignments == () and run.already_registered == 1
    assert len(run.result.invoices_without_order) == 2


def test_summary_lists_counts_and_problem_rows(tmp_path):
    run = _prepare(tmp_path, [api_item("1"), api_item("3", "박민수")],
                   hanjin_row(1, "111"), hanjin_row(2, "999", "이영희"))
    text = invoice_summary(run)
    assert "일치 1건" in text and "주문 못 찾은 송장 1건" in text and "송장 없는 발송대기 주문 1건" in text
    assert "이영희 · 999" in text and "박민수" in text


def test_summary_stays_under_telegram_limit(tmp_path):
    hanjin = [hanjin_row(i, f"T{i:06d}", f"고객{i}") for i in range(400)]
    assert len(invoice_summary(_prepare(tmp_path, [], *hanjin))) < 4096


def test_pending_roundtrip(tmp_path):
    path = tmp_path / "pending.json"
    run = _prepare(tmp_path, [api_item("1")], hanjin_row(1, "111"))
    run_id = add_pending(run, path)
    assert read_pending(path)[run_id]["matched"] == [{"ids": ["1"], "tracking": "111", "courier": "HANJIN"}]
    remove_pending(run_id, path)
    assert read_pending(path) == {}


def _entry():
    return {"when": "10/13 09:05", "matched": [["1", "2", "111"]], "includable": [["1", "2", "111"], ["3", "222"]]}


def test_register_real_success_is_written_to_the_ledger(tmp_path):
    ledger, sent = tmp_path / "ledger.json", []

    def dispatch(items):
        sent.append([(i.product_order_id, i.tracking) for i in items])
        return DispatchResult(("3",), (DispatchFailure("1", "E", "실패"), DispatchFailure("2", "E", "실패")))

    result = register(_entry(), True, dispatch=dispatch, ledger=ledger, now=lambda: WHEN)
    assert sent == [[("1", "111"), ("2", "111"), ("3", "222")]] and len(result.failures) == 2
    assert read_ledger(ledger) == {"3": {"tracking": "222", "when": "2026-10-13 09:05"}}


def test_register_simulated_leaves_the_ledger_alone(tmp_path):
    ledger = tmp_path / "ledger.json"
    result = register(_entry(), False, dispatch=lambda items: dispatch_or_simulate(items, enabled=False),
                      ledger=ledger, now=lambda: WHEN)
    assert result.simulated and result.success_ids == ("1", "2")
    assert not ledger.exists()


def test_register_skips_orders_already_in_the_ledger(tmp_path):
    ledger, sent = tmp_path / "ledger.json", []
    mark_registered({"1": "111"}, WHEN, ledger)
    register(_entry(), False, dispatch=lambda items: sent.append([i.product_order_id for i in items])
             or DispatchResult(()), ledger=ledger)
    assert sent == [["2"]]


@pytest.mark.parametrize("bad", [b"junk"])
def test_bad_file_raises_a_readable_error(tmp_path, bad):
    with pytest.raises(ValueError, match="xlsx"):
        prepare_invoices(bad, fetch=list, ledger=tmp_path / "l.json")


def test_supplier_purchase_order_with_tracking_column_is_matched_and_blank_rows_named(tmp_path):
    from tests.invoice_helpers import returned_row, returned_xlsx
    data = returned_xlsx(returned_row("111"), returned_row(None, None, 전화번호=None, 주소=None, 구분="초벌장어 9kg"))
    run = prepare_invoices(data, fetch=lambda: [api_item("1")], ledger=tmp_path / "ledger.json", when=WHEN)
    assert [a.tracking for a in run.result.matched()] == ["111"]
    text = invoice_summary(run)
    assert "송장번호가 없는 줄 1개" in text and "3번 줄 · 받는 사람 없음 · 초벌장어 9kg" in text


def test_register_sends_each_shipment_with_its_own_courier(tmp_path):
    entry = {"when": "10/13 09:05", "includable": [], "matched": [
        {"ids": ["1"], "tracking": "600663392321", "courier": "CJGLS"},
        {"ids": ["2", "3"], "tracking": "6890179430672", "courier": "EPOST"}]}
    sent = []
    register(entry, False, dispatch=lambda items: sent.extend(items) or DispatchResult(()),
             ledger=tmp_path / "ledger.json")
    assert [(i.product_order_id, i.courier_code) for i in sent] == [("1", "CJGLS"), ("2", "EPOST"), ("3", "EPOST")]
