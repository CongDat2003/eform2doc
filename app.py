"""
app.py — Web UI cho eform2doc.

Chạy local:  streamlit run app.py
Deploy free: Streamlit Community Cloud / Hugging Face Spaces
"""

import io
import json
import time

import streamlit as st

from eform_parser import (
    parse_eform, all_placeholders, GRID_TYPES, MULTI_CHECKBOX_TYPES,
)
from doc_tools import scan_placeholders, generate_template

st.set_page_config(page_title="eform2doc", page_icon="📄", layout="wide")

st.title("📄 eform2doc")
st.caption("Sinh và kiểm tra DOC template từ eform JSON (Form.io)")


# ------------------------------------------------------------------ helpers

def fields_to_rows(fields):
    grid_keys = {f.key for f in fields if f.type in GRID_TYPES}
    rows = []
    for f in fields:
        if f.parent in grid_keys:
            rows.append({
                "Key": f"    └ {f.key}",
                "Nhãn": f.label,
                "Loại": f.type,
                "Bắt buộc": "✓" if f.required else "",
                "Placeholder": f"cột của ${{{f.parent}#table}}",
            })
        elif f.type in MULTI_CHECKBOX_TYPES:
            rows.append({
                "Key": f.key, "Nhãn": f.label, "Loại": f.type,
                "Bắt buộc": "✓" if f.required else "",
                "Placeholder": f"{len(f.options)} option ↓",
            })
            for val, lbl, ph in f.option_placeholders():
                rows.append({
                    "Key": f"    · {val}", "Nhãn": lbl, "Loại": "",
                    "Bắt buộc": "", "Placeholder": ph,
                })
        else:
            rows.append({
                "Key": f.key, "Nhãn": f.label, "Loại": f.type,
                "Bắt buộc": "✓" if f.required else "",
                "Placeholder": f.placeholder,
            })
    return rows


def build_markdown(fields):
    grid_keys = {f.key for f in fields if f.type in GRID_TYPES}
    lines = ["# Bảng mapping eform ↔ DOC template", "",
             "| Key | Nhãn | Loại | Bắt buộc | Placeholder |", "|---|---|---|---|---|"]
    for f in fields:
        if f.parent in grid_keys:
            key_txt, ph = f"&nbsp;&nbsp;└ `{f.key}`", f"cột của `${{{f.parent}#table}}`"
        else:
            key_txt, ph = f"`{f.key}`", f"`{f.placeholder}`"
        label = (f.label or "").replace("|", "\\|")
        lines.append(f"| {key_txt} | {label} | {f.type} | {'✓' if f.required else ''} | {ph} |")

    multi = [f for f in fields if f.type in MULTI_CHECKBOX_TYPES]
    if multi:
        lines += ["", "## Chi tiết nhóm checkbox", ""]
        for f in multi:
            lines += [f"### `{f.key}` — {f.label}", "",
                      "| Value | Nhãn | Placeholder |", "|---|---|---|"]
            for val, lbl, ph in f.option_placeholders():
                lines.append(f"| `{val}` | {lbl.replace('|', chr(92) + '|')} | `{ph}` |")
            lines.append("")
    return "\n".join(lines) + "\n"


# ------------------------------------------------------------------ input

uploaded = st.file_uploader("Tải file eform JSON", type=["json"])

with st.expander("Hoặc dán JSON trực tiếp"):
    pasted = st.text_area("JSON", height=180, label_visibility="collapsed")

raw = None
source_name = "form"
if uploaded is not None:
    raw = json.load(uploaded)
    source_name = uploaded.name.rsplit(".", 1)[0]
elif pasted.strip():
    try:
        raw = json.loads(pasted)
    except json.JSONDecodeError as e:
        st.error(f"JSON không hợp lệ: {e}")

if raw is None:
    st.info("Tải lên hoặc dán eform JSON để bắt đầu.")
    st.stop()


# ------------------------------------------------------------------ parse

t0 = time.perf_counter()
fields = parse_eform(raw)
parse_ms = (time.perf_counter() - t0) * 1000

placeholders = all_placeholders(fields)
grids = [f for f in fields if f.type in GRID_TYPES]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Field", len(fields))
c2.metric("Placeholder", len(placeholders))
c3.metric("Bảng (datagrid)", len(grids))
c4.metric("Thời gian đọc", f"{parse_ms:.0f} ms")

if grids:
    st.success("Bảng phát hiện được: " + " · ".join(
        f"`${{{g.key}#table}}`" for g in grids))

tab1, tab2, tab3 = st.tabs(["📋 Danh sách key", "📄 Sinh DOC template", "🔍 Đối chiếu DOC"])


# ------------------------------------------------------------------ tab 1

with tab1:
    st.dataframe(fields_to_rows(fields), use_container_width=True, height=460)
    st.download_button(
        "⬇ Tải bảng mapping (.md)",
        build_markdown(fields),
        file_name=f"{source_name}_mapping.md",
        mime="text/markdown",
    )


# ------------------------------------------------------------------ tab 2

with tab2:
    title = st.text_input("Tiêu đề văn bản", value="BIỂU MẪU")
    subtitle = st.text_input("Phụ đề (không bắt buộc)", value="")

    if st.button("Sinh DOC template", type="primary"):
        t0 = time.perf_counter()
        buf = io.BytesIO()
        import tempfile, os
        tmp = os.path.join(tempfile.mkdtemp(), "out.docx")
        generate_template(fields, tmp, title=title, subtitle=subtitle)
        with open(tmp, "rb") as fh:
            buf.write(fh.read())
        gen_ms = (time.perf_counter() - t0) * 1000

        st.success(f"Xong trong {gen_ms:.0f} ms — {len(placeholders)} placeholder")
        st.download_button(
            "⬇ Tải template.docx",
            buf.getvalue(),
            file_name=f"{source_name}_template.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    with st.expander("Danh sách placeholder sẽ có trong DOC"):
        st.code("\n".join(sorted(placeholders)), language="text")


# ------------------------------------------------------------------ tab 3

with tab3:
    st.write("Tải lên DOC template đang dùng để kiểm tra có lệch key với eform không.")
    doc_file = st.file_uploader("DOC template (.docx)", type=["docx"], key="check")

    if doc_file is not None:
        import tempfile, os
        tmp = os.path.join(tempfile.mkdtemp(), "check.docx")
        with open(tmp, "wb") as fh:
            fh.write(doc_file.read())

        t0 = time.perf_counter()
        found = scan_placeholders(tmp)
        scan_ms = (time.perf_counter() - t0) * 1000

        needed = set(placeholders)
        missing = sorted(needed - found)
        extra = sorted(found - needed)
        ok = needed & found

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("eform cần", len(needed))
        m2.metric("DOC đang có", len(found))
        m3.metric("Khớp", len(ok))
        m4.metric("Thời gian", f"{scan_ms:.0f} ms")

        if not missing and not extra:
            st.success("Khớp hoàn toàn — DOC template sẵn sàng merge.")
        else:
            if missing:
                st.error(f"THIẾU {len(missing)} placeholder — eform có nhưng DOC chưa có:")
                st.code("\n".join(missing), language="text")
            if extra:
                st.warning(f"THỪA {len(extra)} placeholder — DOC có nhưng eform không sinh ra:")
                st.code("\n".join(extra), language="text")


st.divider()
st.caption(
    "Quy tắc: `textfield/select/datetime → ${Key}` · "
    "`checkbox → ${Key#checkbox}` · "
    "`selectboxes → ${Key.value#checkbox}` · "
    "`datagrid/editgrid → ${TenBang#table}`"
)
