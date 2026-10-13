"""송장 입력 page logic (pure) plus a smoke run of the page."""
from datetime import datetime
from io import BytesIO

import msoffcrypto
import openpyxl
import pytest
from streamlit.testing.v1 import AppTest

from tests.invoice_helpers import hanjin_row, hanjin_xlsx, order_export_xlsx, order_row
from ui import invoice_logic as logic
from core.invoice_match import KST


def _files():
    orders = order_export_xlsx(order_row("1"), order_row("2", "김철수"), order_row("3", "박민수"))
    invoices = hanjin_xlsx(hanjin_row(1, "111"), hanjin_row(2, "222", "김철수", **{"받으시는 분 전화": "01099999999"}),
                           hanjin_row(3, "333", "이영희"))
    return orders, invoices


def test_metrics_and_tables():
    outcome = logic.match_files(*_files())
    assert logic.metrics(outcome.result) == {"일치": 1, "확인 필요": 1, "주문 못 찾은 송장": 1, "송장 없는 주문": 1}
    assert list(logic.matched_table(outcome.result)["송장번호"]) == ["111"]
    assert list(logic.check_table(outcome.result)["사유"]) == ["전화번호 다름"]
    assert list(logic.loose_invoice_table(outcome.result)["받는분"]) == ["이영희"]
    assert list(logic.unshipped_table(outcome.result)["수취인"]) == ["박민수"]


def test_download_with_and_without_checks():
    outcome = logic.match_files(*_files())
    when = datetime(2026, 10, 13, 9, 5, tzinfo=KST)

    def ids(include):
        name, data = logic.build_download(outcome, include, when)
        assert name == "송장업로드_20261013_0905.xlsx"
        sheet = openpyxl.load_workbook(BytesIO(data)).active
        return [row[0].value for row in sheet.iter_rows(min_row=2)]

    assert ids(False) == ["1"] and ids(True) == ["1", "2"]


def test_encrypted_order_file_needs_the_password():
    orders, invoices = _files()
    plain = BytesIO(orders)
    locked = BytesIO()
    try:
        msoffcrypto.OfficeFile(plain).encrypt("pw-for-test", locked)
    except Exception:  # noqa: BLE001
        pytest.skip("this msoffcrypto version cannot encrypt")
    with pytest.raises(ValueError, match="암호화"):
        logic.match_files(locked.getvalue(), invoices)
    with pytest.raises(ValueError, match="비밀번호"):
        logic.match_files(locked.getvalue(), invoices, "wrong")
    assert logic.match_files(locked.getvalue(), invoices, "pw-for-test").result.matched()


def test_bad_files_give_readable_errors():
    orders, invoices = _files()
    with pytest.raises(ValueError, match="운송장번호"):
        logic.match_files(orders, orders)
    with pytest.raises(ValueError, match="상품주문번호"):
        logic.match_files(invoices, invoices)


def test_page_renders_without_files():
    app = AppTest.from_string("from ui import invoice_page\ninvoice_page.render()").run()
    assert not app.exception and "두 파일" in app.info[0].value
