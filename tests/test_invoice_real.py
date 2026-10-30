"""Matching against the owner's real files (kept outside the repo); skipped when they are absent."""
from pathlib import Path

import pytest

from core.invoice_match import match_invoices, read_invoice_file, read_order_export

REAL_DIR = Path(__file__).resolve().parent.parent.parent / "Version 2.0"
HANJIN = REAL_DIR / "한진 송장파일.xlsx"
HUMAN = REAL_DIR / "송장 업로드용.xlsx"

pytestmark = pytest.mark.skipif(not (HANJIN.exists() and HUMAN.exists()), reason="real invoice files not present")


def test_every_order_gets_the_tracking_number_a_human_chose():
    human = read_order_export(HUMAN.read_bytes())
    expected = dict(zip(human["상품주문번호"], human["송장번호"], strict=True))
    orders = human.assign(송장번호="", 택배사="", 주문상태="발송대기")
    result = match_invoices(orders, read_invoice_file(HANJIN.read_bytes()))
    got = {pid: a.tracking for a in result.assignments for pid in a.product_order_ids}
    wrong = sorted(pid for pid in expected if got.get(pid) != expected[pid])
    assert wrong == [], f"상품주문번호 {wrong} 의 송장번호가 사람이 넣은 것과 다릅니다"
    assert len(result.assignments) == 11 and len(got) == 12
    assert all(a.status == "MATCHED" for a in result.assignments)
    assert len(result.invoices_without_order) == 6 and result.orders_without_invoice == ()


SUPPLIER_LIST = REAL_DIR / "갑오징어 생물오징어 등 송장.xls"


@pytest.mark.skipif(not SUPPLIER_LIST.exists(), reason="real supplier list not present")
def test_real_supplier_xls_list_is_read_with_its_couriers():
    invoices = read_invoice_file(SUPPLIER_LIST.read_bytes())
    assert len(invoices) == 66 and all(i.tracking.isdigit() for i in invoices)
    assert {i.carrier for i in invoices} == {"CJ대한통운", "우체국택배", "로젠택배"}


BOOKING_FORM = REAL_DIR / "프리미엄 과메기 송장번호.xlsx"


@pytest.mark.skipif(not BOOKING_FORM.exists(), reason="real booking form not present")
def test_real_booking_form_is_read_as_hanjin_and_skips_the_stray_mark():
    invoices = read_invoice_file(BOOKING_FORM.read_bytes())
    assert len(invoices) == 10 and all(i.tracking.isdigit() and len(i.tracking) == 12 for i in invoices)
    assert all(i.name and len(i.zip_code) == 5 and i.address for i in invoices)
    assert {i.carrier for i in invoices} == {"한진택배"}
