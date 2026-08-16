"""
eform_parser.py — Trích xuất schema field từ Form.io eform JSON.

Xuất ra danh sách Field phẳng, mỗi field có:
  - key          : property name (dùng làm ${key} trong DOC)
  - label        : nhãn hiển thị (đã strip HTML)
  - type         : loại component Form.io
  - placeholder  : cú pháp placeholder tương ứng trong DOC
  - required     : bắt buộc hay không
  - parent       : key của grid cha (nếu là subfield)
  - options      : list (value, label) cho selectboxes/select/radio
"""

import json
import re
from dataclasses import dataclass, field as dc_field
from typing import List, Optional, Tuple

# Component không sinh dữ liệu -> bỏ qua
LAYOUT_TYPES = {
    "panel", "fieldset", "columns", "well", "table", "tabs",
    "htmlelement", "content", "button", "column",
}

# Component chứa mảng con
GRID_TYPES = {"datagrid", "editgrid"}

# Component checkbox nhóm (nhiều key con dạng Key.value)
MULTI_CHECKBOX_TYPES = {"selectboxes"}


@dataclass
class Field:
    key: str
    label: str
    type: str
    required: bool = False
    parent: Optional[str] = None
    options: List[Tuple[str, str]] = dc_field(default_factory=list)
    max_length: Optional[int] = None
    hidden: bool = False

    @property
    def placeholder(self) -> str:
        """Cú pháp placeholder tương ứng trong DOC template."""
        if self.parent:
            return "${%s.%s}" % (self.parent, self.key)
        if self.type in GRID_TYPES:
            return "${%s#table}" % self.key
        if self.type == "checkbox":
            return "${%s#checkbox}" % self.key
        if self.type in MULTI_CHECKBOX_TYPES:
            return "(nhiều dòng - xem options)"
        return "${%s}" % self.key

    def option_placeholders(self) -> List[Tuple[str, str, str]]:
        """Với selectboxes: trả về [(value, label, placeholder), ...]"""
        if self.type not in MULTI_CHECKBOX_TYPES:
            return []
        return [
            (val, lbl, "${%s.%s#checkbox}" % (self.key, val))
            for val, lbl in self.options
        ]


def strip_html(text: str) -> str:
    """Bỏ tag HTML khỏi label, gọn khoảng trắng."""
    if not text:
        return ""
    text = re.sub(r"<[^>]+>", "", str(text))
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\s+", " ", text).strip()


def _extract_options(comp: dict) -> List[Tuple[str, str]]:
    """Lấy danh sách (value, label) từ select / selectboxes / radio."""
    # selectboxes & radio: values ở root
    values = comp.get("values")
    if isinstance(values, list) and values:
        return [
            (str(v.get("value", "")), strip_html(v.get("label", "")))
            for v in values
            if isinstance(v, dict) and v.get("value") not in (None, "")
        ]
    # select dataSrc=values: data.values
    data = comp.get("data")
    if isinstance(data, dict):
        dv = data.get("values")
        if isinstance(dv, list) and dv:
            return [
                (str(v.get("value", "")), strip_html(v.get("label", "")))
                for v in dv
                if isinstance(v, dict) and v.get("value") not in (None, "")
            ]
    return []


def _walk(comp, out: List[Field], parent: Optional[str]):
    """Duyệt đệ quy cây component."""
    if isinstance(comp, list):
        for c in comp:
            _walk(c, out, parent)
        return
    if not isinstance(comp, dict):
        return

    ctype = comp.get("type")
    key = comp.get("key")

    # Columns: duyệt vào từng cột
    if ctype == "columns":
        for col in comp.get("columns") or []:
            if isinstance(col, dict):
                _walk(col.get("components") or [], out, parent)
        return

    # Grid: ghi nhận grid rồi duyệt subfield với parent = key grid
    if ctype in GRID_TYPES and key:
        validate = comp.get("validate") or {}
        out.append(Field(
            key=key,
            label=strip_html(comp.get("label", "")),
            type=ctype,
            required=bool(validate.get("required")),
            parent=parent,
            hidden=bool(comp.get("hidden")),
        ))
        _walk(comp.get("components") or [], out, key)
        return

    # Component có dữ liệu
    if ctype and ctype not in LAYOUT_TYPES and comp.get("input") and key:
        validate = comp.get("validate") or {}
        max_len = validate.get("maxLength")
        out.append(Field(
            key=key,
            label=strip_html(comp.get("label", "")) or key,
            type=ctype,
            required=bool(validate.get("required")),
            parent=parent,
            options=_extract_options(comp),
            max_length=int(max_len) if isinstance(max_len, (int, float)) and max_len else None,
            hidden=bool(comp.get("hidden")),
        ))

    # Duyệt tiếp con (panel, fieldset, well...)
    _walk(comp.get("components") or [], out, parent)


def parse_eform(data) -> List[Field]:
    """Nhận eform JSON (dict hoặc list) -> danh sách Field."""
    if isinstance(data, dict) and "components" in data and "type" not in data:
        data = data["components"]
    out: List[Field] = []
    _walk(data, out, None)

    # Khử trùng lặp theo (parent, key), giữ bản đầu tiên
    seen = set()
    unique: List[Field] = []
    for f in out:
        sig = (f.parent, f.key)
        if sig in seen:
            continue
        seen.add(sig)
        unique.append(f)
    return unique


def parse_eform_file(path: str) -> List[Field]:
    with open(path, "r", encoding="utf-8") as fh:
        return parse_eform(json.load(fh))


def all_placeholders(fields: List[Field], with_columns: bool = True) -> List[str]:
    """
    Tập hợp mọi placeholder mà DOC template cần có.

    with_columns=True (mặc định): bảng sinh ra
        ${Grid#table}  + mỗi cột một ${Grid.SubKey}
    with_columns=False: bảng chỉ sinh ${Grid#table}
    """
    result = []
    grid_keys = {f.key for f in fields if f.type in GRID_TYPES}
    for f in fields:
        if f.parent in grid_keys:
            if with_columns:
                result.append("${%s.%s}" % (f.parent, f.key))
            continue
        if f.type in MULTI_CHECKBOX_TYPES:
            result.extend(ph for _, _, ph in f.option_placeholders())
        else:
            result.append(f.placeholder)
    return result
