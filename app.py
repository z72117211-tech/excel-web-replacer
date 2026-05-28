import re
from io import BytesIO
from typing import Dict, Tuple

import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="Excel 關鍵字替換工具",
    page_icon="🧰",
    layout="centered",
)

DEFAULT_RULES = """蝦皮=>
聊聊=>
shopee=>
Shopee=>
SHOPEE=>"""

TRACKING_PARAMS = [
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
    "fbclid", "gclid", "yclid", "mc_cid", "mc_eid"
]


def parse_rules(text: str, ignore_case: bool) -> Dict[str, str]:
    rules: Dict[str, str] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=>" not in line:
            continue
        old, new = line.split("=>", 1)
        old = old.strip()
        new = new.strip()
        if old:
            rules[old] = new
    return rules


def replace_text(value: str, rules: Dict[str, str], ignore_case: bool) -> str:
    result = value
    for old, new in rules.items():
        if ignore_case:
            result = re.sub(re.escape(old), new, result, flags=re.IGNORECASE)
        else:
            result = result.replace(old, new)
    return result


def remove_tracking_params(url: str) -> str:
    if "?" not in url:
        return url

    base, query = url.split("?", 1)
    fragment = ""
    if "#" in query:
        query, fragment = query.split("#", 1)
        fragment = "#" + fragment

    kept = []
    for part in query.split("&"):
        if not part:
            continue
        key = part.split("=", 1)[0]
        if key not in TRACKING_PARAMS:
            kept.append(part)

    if kept:
        return base + "?" + "&".join(kept) + fragment
    return base + fragment


def process_excel(uploaded_file, rules: Dict[str, str], ignore_case: bool, clean_tracking: bool) -> Tuple[BytesIO, int, int]:
    wb = load_workbook(uploaded_file)
    changed_cells = 0
    changed_links = 0

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    original = cell.value
                    updated = replace_text(original, rules, ignore_case)
                    if clean_tracking and updated.startswith(("http://", "https://")):
                        updated = remove_tracking_params(updated)
                    if updated != original:
                        cell.value = updated
                        changed_cells += 1

                if cell.hyperlink and cell.hyperlink.target:
                    original_target = cell.hyperlink.target
                    updated_target = replace_text(original_target, rules, ignore_case)
                    if clean_tracking:
                        updated_target = remove_tracking_params(updated_target)
                    if updated_target != original_target:
                        cell.hyperlink = updated_target
                        changed_links += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, changed_cells, changed_links


st.title("Excel 關鍵字替換工具")
st.caption("上傳 Excel，批次替換儲存格文字與超連結網址。適合內部同事使用。")

with st.expander("使用方式", expanded=True):
    st.write(
        "每行一組替換規則，格式為 `原文字=>新文字`。如果右邊留空，就代表刪除該關鍵字。"
    )
    st.code("蝦皮=>\n聊聊=>\nshopee=>平台", language="text")

uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx"])

rules_text = st.text_area(
    "替換規則",
    value=DEFAULT_RULES,
    height=180,
    help="範例：蝦皮=>平台。右側留空代表刪除。",
)

col1, col2 = st.columns(2)
with col1:
    ignore_case = st.checkbox("英文大小寫不區分", value=True)
with col2:
    clean_tracking = st.checkbox("順便移除常見追蹤參數", value=True)

if st.button("開始處理", type="primary", use_container_width=True):
    if uploaded_file is None:
        st.error("請先上傳 Excel 檔案。")
    else:
        rules = parse_rules(rules_text, ignore_case)
        if not rules and not clean_tracking:
            st.error("請至少輸入一組替換規則，或勾選移除追蹤參數。")
        else:
            try:
                output, changed_cells, changed_links = process_excel(
                    uploaded_file, rules, ignore_case, clean_tracking
                )
                st.success(f"處理完成：已修改 {changed_cells} 個儲存格、{changed_links} 個超連結。")
                st.download_button(
                    label="下載處理後 Excel",
                    data=output,
                    file_name="processed_excel.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"處理失敗：{exc}")

st.divider()
st.caption("提醒：建議先用備份檔測試。此工具會在記憶體中處理檔案，不會主動保存上傳內容。")
