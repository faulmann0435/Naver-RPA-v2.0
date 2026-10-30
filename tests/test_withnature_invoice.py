"""자연이랑 channel in the 송장 flow: split from Naver, portal upload file, ledger, bot routing."""
import json
from datetime import date, datetime
from functools import partial
from io import BytesIO

import openpyxl

from core.invoice_match import Assignment
from core.withnature import read_withnature_export, to_order_frame
from tests.invoice_helpers import (
    ADDRESS,
    api_item,
    hanjin_row,
    hanjin_xlsx,
    list_row,
    list_xlsx,
)
from tests.test_telegram_bot import OWNER, FakeTelegram, _buttons, _doc, _press
from tests.withnature_helpers import export_row, export_xlsx
from worker.dispatch_client import DispatchResult
from worker.invoice_run import (
    _portal_rows,
    invoice_summary,
    prepare_invoices,
    read_withnature_ledger,
    register,
)
from worker.naver_adapter import KST
from worker.telegram_bot import OrderBot
from worker.withnature_run import WithnatureRun, save_orders

WHEN = datetime(2026, 10, 30, 17, 5, tzinfo=KST)
WN_ADDRESS = "경기도 가나시 다라구 마바로 1 (사아동* 자차아파트) 101동 202호"
WN_ADDRESS_RESTORED = WN_ADDRESS.replace("* ", ", ")


def _save(tmp_path, *rows):
    frame = to_order_frame(read_withnature_export(export_xlsx(*rows)))
    save_orders(frame, datetime.now(KST).date(), tmp_path / "wn_orders.json")


def _wn_hanjin(seq, tracking, name="김철수"):
    return hanjin_row(seq, tracking, name, 받는분우편번호="13482", 받는분총주소=WN_ADDRESS_RESTORED)


def _prepare(tmp_path, items, *rows, xlsx=None):
    return prepare_invoices(xlsx or hanjin_xlsx(*rows), fetch=lambda: items, ledger=tmp_path / "ledger.json",
                            when=WHEN)


def _portal_file(run, kind=""):
    (f,) = [f for f in run.files if f.filename == f"자연이랑_송장업로드{kind}_1030_1705.xlsx"]
    return list(openpyxl.load_workbook(BytesIO(f.data)).active.iter_rows(values_only=True))


def test_mixed_file_splits_naver_and_withnature(tmp_path):
    _save(tmp_path, export_row("O1", "김철수", 받는사람주소=WN_ADDRESS))
    run = _prepare(tmp_path, [api_item("1")], hanjin_row(1, "111"), _wn_hanjin(2, "222"))
    (a,) = run.result.matched()
    assert a.product_order_ids == ("1",) and a.tracking == "111"
    assert run.result.invoices_without_order == () and run.result.orders_without_invoice == ()
    assert [f.filename for f in run.files] == ["송장업로드_일치_1030_1705.xlsx", "자연이랑_송장업로드_1030_1705.xlsx"]
    assert _portal_file(run) == [("주문번호", "운송장번호", "택배업체코드"), ("O1", "222", "03")]
    text = invoice_summary(run)
    assert "일치 1건" in text and "🌿 자연이랑" in text and "1건 → 파일을 사이트 '출고 및 운송장등록 > 일괄등록'" in text
    assert "'첫 행은 컬럼으로 제외'를 체크하세요" in text


def test_supplier_list_courier_name_gives_the_code(tmp_path):
    _save(tmp_path, export_row("O1", "김철수"))
    data = list_xlsx(list_row("555", "씨제이대한통운", "김철수"), list_row("666", "로젠택배", "김철수", 연락처="010-9"))
    run = _prepare(tmp_path, [], xlsx=data)
    assert _portal_file(run)[1:] == [("O1", "555", "02")]


def test_pairs_already_uploaded_are_skipped_and_counted(tmp_path):
    _save(tmp_path, export_row("O1", "김철수", 받는사람주소=WN_ADDRESS))
    first = _prepare(tmp_path, [], _wn_hanjin(1, "222"))
    assert read_withnature_ledger(tmp_path / "wn_ledger.json") == {"O1|222": "2026-10-30 17:05"}
    assert [f.filename for f in first.files] == ["자연이랑_송장업로드_1030_1705.xlsx"]
    again = _prepare(tmp_path, [], _wn_hanjin(1, "222"))
    assert again.files == () and again.withnature.skipped == 1
    assert "이미 올린 송장 1건 제외" in invoice_summary(again)
    other = _prepare(tmp_path, [], _wn_hanjin(1, "333"))  # a new tracking number is a new pair
    assert _portal_file(other)[1:] == [("O1", "333", "03")]


def test_needs_check_goes_to_its_own_file_and_is_not_recorded(tmp_path):
    _save(tmp_path, export_row("O1", "김철수", 받는사람주소=WN_ADDRESS))
    wrong_phone = _wn_hanjin(1, "222")
    wrong_phone["받으시는 분 전화"] = "01099999999"
    run = _prepare(tmp_path, [], wrong_phone)
    assert [f.filename for f in run.files] == ["자연이랑_송장업로드_확인필요_1030_1705.xlsx"]
    assert _portal_file(run, "_확인필요")[1:] == [("O1", "222", "03")]
    assert run.withnature.matched == 0 and run.withnature.checks == 1
    assert read_withnature_ledger(tmp_path / "wn_ledger.json") == {}


def test_unknown_courier_is_left_out_and_listed():
    a = Assignment(("O1",), "777", "김철수", "상품", "MATCHED", True, carrier="택배X")
    assert _portal_rows((a,), {}) == ([], 0, ["김철수 · 777 · 택배X"])


def test_one_box_matching_both_channels_is_not_included_anywhere(tmp_path):
    _save(tmp_path, export_row("O1", "홍길동", 받는사람주소=ADDRESS, 우편번호="08239"))
    run = _prepare(tmp_path, [api_item("1")], hanjin_row(1, "111"))
    assert run.result.includable() == () and run.files == ()
    assert [a.reason for a in run.result.needs_check()] == ["네이버·자연이랑 주문에 걸친 송장"] and run.withnature.excluded
    assert "네이버·자연이랑 주문에 걸친 송장" in invoice_summary(run)


def test_without_saved_orders_nothing_changes(tmp_path):
    run = _prepare(tmp_path, [api_item("1")], hanjin_row(1, "111"))
    assert run.withnature is None and [f.filename for f in run.files] == ["송장업로드_일치_1030_1705.xlsx"]
    assert "자연이랑" not in invoice_summary(run)
    assert not (tmp_path / "wn_ledger.json").exists()


def test_old_saved_orders_are_ignored(tmp_path):
    frame = to_order_frame(read_withnature_export(export_xlsx(export_row("O1", "김철수"))))
    save_orders(frame, date(2020, 1, 1), tmp_path / "wn_orders.json")
    assert _prepare(tmp_path, [], _wn_hanjin(1, "222")).withnature is None


# --- nothing of 자연이랑 reaches Naver ------------------------------------------------------------

def test_register_never_sends_a_withnature_order_number(tmp_path):
    entry = {"when": "10/30 17:05", "includable": [], "matched": [
        {"ids": ["O26100000001", "1"], "tracking": "222", "courier": "HANJIN"}]}
    sent = []
    register(entry, False, dispatch=lambda items: sent.extend(items) or DispatchResult(()),
             ledger=tmp_path / "ledger.json")
    assert [i.product_order_id for i in sent] == ["1"]


class _BotTelegram(FakeTelegram):
    data = b""

    def download(self, file_id):
        self.sent.append(("download", {"file_id": file_id}))
        return self.data


def _bot(tmp_path, items, withnature_runner=None):
    tg = _BotTelegram()
    ledger = tmp_path / "ledger.json"
    sent = []

    def dispatch(batch):
        sent.extend((i.product_order_id, i.tracking) for i in batch)
        return DispatchResult(tuple(i.product_order_id for i in batch))

    extra = {"withnature_runner": withnature_runner} if withnature_runner else {}
    bot = OrderBot(tg, frozenset({OWNER}), pending_file=tmp_path / "p.json",
                   invoice_preparer=partial(prepare_invoices, fetch=lambda: items, ledger=ledger, when=WHEN),
                   invoice_register=partial(register, dispatch=dispatch, ledger=ledger),
                   invoice_pending_file=tmp_path / "inv.json", **extra)
    return tg, bot, sent


def test_naver_buttons_register_only_naver_ids(tmp_path):
    _save(tmp_path, export_row("O1", "김철수", 받는사람주소=WN_ADDRESS))
    tg, bot, sent = _bot(tmp_path, [api_item("1")])
    tg.data = hanjin_xlsx(hanjin_row(1, "111"), _wn_hanjin(2, "222"))
    bot.handle(_doc())
    assert "자연이랑_송장업로드_1030_1705.xlsx" in [p["name"] for m, p in tg.sent if m == "sendDocument"]
    assert [text for text, _ in _buttons(tg)] == ["🚚 일치 1건 송장 등록", "✖ 취소"]
    bot.handle(_press(_buttons(tg)[0][1]))
    assert sent == [("1", "111")]


# --- document routing -------------------------------------------------------------------------------

_FILE = "자연이랑_속초 발주양식_20261030.xlsx"
_RUN = WithnatureRun(files=[{"filename": _FILE, "data": BytesIO(b"x"), "vendor": "v"}], order_count=1,
                     vendor_counts=((_FILE, 1),), unmatched=0, unclassified=0, needs_review=0, source="t",
                     when=WHEN)


def test_a_withnature_export_goes_to_the_withnature_runner(tmp_path):
    calls = []
    tg, bot, _ = _bot(tmp_path, [], withnature_runner=lambda data: calls.append(data) or _RUN)
    tg.data = export_xlsx(export_row())
    bot.handle(_doc("fulfillment020_2026-10-30.xlsx"))
    assert calls == [tg.data]
    assert [m for m, _ in tg.sent] == ["sendMessage", "download", "sendDocument", "sendMessage"]
    last = tg.sent[-1][1]
    assert last["text"].startswith("🌿 자연이랑 발주서 1개") and "inline_keyboard" not in last["reply_markup"]
    assert not (tmp_path / "inv.json").exists()


def test_other_excel_files_still_take_the_invoice_path(tmp_path):
    calls = []
    tg, bot, _ = _bot(tmp_path, [api_item("1")], withnature_runner=lambda data: calls.append(data))
    tg.data = hanjin_xlsx(hanjin_row(1, "111"))
    bot.handle(_doc())
    assert calls == [] and tg.texts()[-1].startswith("🚚 송장 확인")


def test_withnature_failure_is_reported(tmp_path):
    def broken(data):
        raise ValueError("깨짐")

    tg, bot, _ = _bot(tmp_path, [], withnature_runner=broken)
    tg.data = export_xlsx(export_row())
    bot.handle(_doc())
    assert tg.texts()[-1].startswith("⚠️ 자연이랑 발주서 실패")


def test_help_text_mentions_the_withnature_file(tmp_path):
    tg, bot, _ = _bot(tmp_path, [])
    bot.handle({"message": {"chat": {"id": OWNER}, "text": "안녕"}})
    assert "자연이랑 출고작업관리 엑셀" in tg.texts()[-1]


def test_ledger_file_is_plain_json(tmp_path):
    _save(tmp_path, export_row("O1", "김철수", 받는사람주소=WN_ADDRESS))
    _prepare(tmp_path, [], _wn_hanjin(1, "222"))
    assert list(json.loads((tmp_path / "wn_ledger.json").read_text(encoding="utf-8"))) == ["O1|222"]
