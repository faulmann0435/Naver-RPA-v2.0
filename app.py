"""Sokcho Order Processing System v15: Streamlit entry (navigation + sidebar).

UI only; processing logic lives in core/, storage in store/, screens in ui/.
"""
import streamlit as st

from ui import dictionary_page, history_page, order_page, preview_page, rules_page
from ui.context import (  # noqa: F401  # re-exported: tests and other modules import them from app
    CONFIG_PATH,
    get_config,
    load_config_local,
    render_sidebar,
)
from ui.password_gate import require_password

PAGE_TITLE = "속초 발주 처리 시스템 v15"


def main() -> None:
    st.set_page_config(page_title=PAGE_TITLE, layout="wide")
    if not require_password():
        return  # nothing below (config, sidebar, pages, data store) runs before the password is accepted
    try:
        _config, config_source = get_config()
    except FileNotFoundError:
        st.error(f"설정 파일을 찾을 수 없습니다. 경로: {CONFIG_PATH}")
        st.info("config.xlsx를 앱과 같은 폴더에 두거나 경로를 확인하세요.")
        return
    except ValueError as e:
        st.error(str(e))
        return
    render_sidebar(config_source)
    navigation = st.navigation(
        [
            st.Page(order_page.render, title="주문처리", icon="📦", url_path="order", default=True),
            st.Page(dictionary_page.render, title="품목 관리", icon="📖", url_path="dictionary"),
            st.Page(preview_page.render, title="결과 확인", icon="🧪", url_path="preview"),
            st.Page(rules_page.render, title="고급 설정", icon="⚙️", url_path="rules"),
            st.Page(history_page.render, title="변경 이력", icon="🕓", url_path="history"),
        ]
    )
    navigation.run()


if __name__ == "__main__":
    main()
