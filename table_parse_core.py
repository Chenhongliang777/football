#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""表格解析共用逻辑（公告附件、申报表）。"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd

INVALID_NAME_CELLS = frozenset(
    {"姓名", "名称", "序号", "编号", "合计", "备注", "性别", "项目", "高校", "等级"}
)

_ROSTER_NAME_BAD_RE = re.compile(
    r"队|协会|印发|等级|年度|足球|批次|授予|关于|通知|体育局|人民政府|"
    r"免冠|照片|单位|电话|项目|运动员|一级|二级|三级|名单|公示|"
    r"日印|印发|增补|技术|称号|标准|批准"
)

# 18 位身份证（含 OCR 末位 x/X 及相邻噪声）
_ID_CARD_RE = re.compile(r"(?<!\d)(\d{17}[\dXx])(?!\d)")

# 性别列常见 OCR 误识 → 归一化
_GENDER_OCR_MAP = {
    "男": "男",
    "女": "女",
    "bl": "男",
    "bh": "男",
    "bh]": "男",
    "al": "女",
    "a": "女",
    "mw": "男",
    "mw]": "男",
}


def normalize_cell(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def map_header_to_field(header: str, aliases: Dict[str, str]) -> Optional[str]:
    h = normalize_cell(header).replace(" ", "").replace("\n", "")
    if not h:
        return None
    if h in aliases:
        return aliases[h]
    for key, field in aliases.items():
        if key in h or h in key:
            return field
    return None


def normalize_ocr_gender(cell: str) -> str:
    s = normalize_cell(cell).lower().replace(" ", "")
    if s in ("男", "女"):
        return s
    return _GENDER_OCR_MAP.get(s, "")


def extract_id_card(text: str) -> str:
    """从单元格或整行文本中提取 18 位身份证号。"""
    raw = normalize_cell(text).replace(" ", "").replace("$", "5")
    m = _ID_CARD_RE.search(raw)
    if not m:
        return ""
    return m.group(1).upper().replace("X", "X")


def is_plausible_person_name(name: str) -> bool:
    n = normalize_cell(name)
    if not n or n in INVALID_NAME_CELLS:
        return False
    if not re.fullmatch(r"[\u4e00-\u9fff·]{2,8}", n):
        return False
    if _ROSTER_NAME_BAD_RE.search(n):
        return False
    return True


def is_header_like_row(cells: List[str]) -> bool:
    joined = "".join(normalize_cell(c) for c in cells)
    return "姓名" in joined and ("序号" in joined or "性别" in joined or "编号" in joined)


def looks_like_data_row(cells: List[str]) -> bool:
    """区分「序号+姓名」数据行与第二行表头，避免漏掉每名表的第一人。"""
    if not cells:
        return False
    c0 = normalize_cell(cells[0])
    if re.fullmatch(r"\d+", c0):
        return True
    if any(normalize_cell(c) in ("男", "女") for c in cells[1:4]):
        return True
    name_like = 0
    for c in cells[1:5]:
        if is_plausible_person_name(normalize_cell(c)):
            name_like += 1
    return name_like >= 1 and "姓名" not in "".join(cells)


def find_header_row(df: pd.DataFrame, max_scan: int = 20) -> Optional[int]:
    for i in range(min(max_scan, len(df))):
        for cell in df.iloc[i]:
            if "姓名" in normalize_cell(cell):
                return i
    return None


def resolve_header_rows(
    df: pd.DataFrame,
    header_idx: int,
    header_aliases: Dict[str, str],
) -> tuple[List[str], int]:
    row_a = [normalize_cell(c) for c in df.iloc[header_idx]]
    data_start = header_idx + 1
    if header_idx + 1 >= len(df):
        return row_a, data_start

    row_b = [normalize_cell(c) for c in df.iloc[header_idx + 1]]
    has_name_a = any("姓名" in h for h in row_a)
    has_name_b = any("姓名" in h for h in row_b)

    # 表头下一行已是数据（如 1 祝宇），切勿合并为双行表头
    if looks_like_data_row(row_b):
        return row_a, header_idx + 1

    if not has_name_a and has_name_b:
        return row_b, header_idx + 2

    if has_name_a and not has_name_b:
        merged: List[str] = []
        width = max(len(row_a), len(row_b))
        for i in range(width):
            a = row_a[i] if i < len(row_a) else ""
            b = row_b[i] if i < len(row_b) else ""
            merged.append(f"{a}{b}" if a and b and a != b else (a or b))
        if any(map_header_to_field(h, header_aliases) for h in merged):
            return merged, header_idx + 2

    return row_a, data_start


def dataframe_to_records(
    df: pd.DataFrame,
    output_fields: Sequence[str],
    header_aliases: Dict[str, str],
) -> List[Dict[str, str]]:
    header_idx = find_header_row(df)
    if header_idx is None:
        return []

    headers, data_start = resolve_header_rows(df, header_idx, header_aliases)
    col_map: Dict[int, str] = {}
    for idx, h in enumerate(headers):
        field = map_header_to_field(h, header_aliases)
        if field:
            col_map[idx] = field

    if "姓名" not in col_map.values():
        return []

    records: List[Dict[str, str]] = []
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        cells = [
            normalize_cell(row.iloc[j]) if j < len(row) else "" for j in range(len(row))
        ]
        if is_header_like_row(cells):
            continue
        athlete = {f: "" for f in output_fields}
        has_name = False
        for idx, field in col_map.items():
            val = normalize_cell(row.iloc[idx]) if idx < len(row) else ""
            if field == "姓名" and val:
                has_name = True
            if field in athlete:
                athlete[field] = val
        _enrich_athlete_from_cells(cells, athlete, output_fields)
        if athlete.get("姓名"):
            has_name = True
        if has_name and is_plausible_person_name(athlete.get("姓名", "")):
            records.append(athlete)
    return records


def _enrich_athlete_from_cells(
    cells: List[str],
    athlete: Dict[str, str],
    output_fields: Sequence[str],
) -> None:
    """按列映射后，再从整行扫描姓名/性别/身份证号（应对 OCR 列错位）。"""
    if not athlete.get("身份证号"):
        for c in cells:
            cid = extract_id_card(c)
            if cid:
                athlete["身份证号"] = cid
                break
    if not athlete.get("性别"):
        for c in cells[:6]:
            g = normalize_ocr_gender(c)
            if g:
                athlete["性别"] = g
                break
    if not is_plausible_person_name(athlete.get("姓名", "")):
        for c in cells[:8]:
            n = normalize_cell(c)
            if is_plausible_person_name(n):
                athlete["姓名"] = n
                break


def dedupe_records(
    records: List[Dict[str, str]],
    key_fields: tuple[str, ...] = ("姓名",),
) -> List[Dict[str, str]]:
    seen: set = set()
    out: List[Dict[str, str]] = []
    for r in records:
        cid = r.get("身份证号", "").strip()
        if cid:
            key: tuple = ("id", cid)
        else:
            key = ("fields",) + tuple(r.get(f, "").strip() for f in key_fields)
            if not any(key[1:]):
                continue
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def _split_ocr_table_line(line: str) -> List[str]:
    if "|" in line:
        return [normalize_cell(c) for c in line.split("|")]
    return [normalize_cell(c) for c in re.split(r"\s{2,}|\t", line) if normalize_cell(c)]


def _row_from_pipe_cells(
    cells: List[str],
    output_fields: Sequence[str],
) -> Optional[Dict[str, str]]:
    if len(cells) < 2:
        return None
    if is_header_like_row(cells):
        return None

    athlete = {f: "" for f in output_fields}
    idx = 0
    if cells and re.fullmatch(r"\d{1,4}", cells[0]):
        idx = 1

    # 序号后常见：姓名、性别
    if idx < len(cells) and is_plausible_person_name(cells[idx]):
        athlete["姓名"] = cells[idx]
        idx += 1
    if idx < len(cells):
        g = normalize_ocr_gender(cells[idx])
        if g:
            athlete["性别"] = g
            idx += 1

    _enrich_athlete_from_cells(cells, athlete, output_fields)
    if is_plausible_person_name(athlete.get("姓名", "")):
        return athlete
    return None


def parse_ocr_roster_lines(
    text: str,
    output_fields: Sequence[str],
) -> List[Dict[str, str]]:
    """OCR 常见行：序号 姓名 男/女（或 | 分列、制表符分隔）。"""
    records: List[Dict[str, str]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line or ("姓名" in line and "序号" in line):
            continue

        if "|" in line:
            row = _row_from_pipe_cells(_split_ocr_table_line(line), output_fields)
            if row:
                records.append(row)
            continue

        m = re.match(
            r"^(?<!\d)(\d{1,4})[.、．]?\s+([\u4e00-\u9fff·]{2,8})\s*([男女])\b",
            line,
        )
        if m:
            row = {f: "" for f in output_fields}
            row["姓名"] = m.group(2)
            row["性别"] = m.group(3)
            if is_plausible_person_name(row["姓名"]):
                records.append(row)
            continue
        parts = re.split(r"\s+|\t", line)
        if len(parts) >= 3 and re.fullmatch(r"\d+", parts[0]):
            name = parts[1]
            gender = normalize_ocr_gender(parts[2]) or (
                parts[2] if parts[2] in ("男", "女") else ""
            )
            if is_plausible_person_name(name):
                row = {f: "" for f in output_fields}
                row["姓名"] = name
                row["性别"] = gender
                _enrich_athlete_from_cells(parts, row, output_fields)
                records.append(row)
    return records


def _is_continuation_row(cells: List[str]) -> bool:
    """行首为性别、无序号（OCR 常把下一人首列识别成 女| / 男|）。"""
    if not cells:
        return False
    if normalize_ocr_gender(cells[0]):
        return True
    if cells[0] in ("男", "女"):
        return True
    return False


def parse_text_lines_to_dataframe(text: str) -> List[pd.DataFrame]:
    """将 OCR/纯文本按行拆成多个候选表（供 dataframe_to_records）。"""
    frames: List[pd.DataFrame] = []
    block: List[List[str]] = []
    last_seq = 0

    def flush() -> None:
        nonlocal block
        if len(block) >= 2:
            frames.append(pd.DataFrame(block))
        block = []

    for line in text.splitlines():
        line = line.strip()
        if not line:
            flush()
            continue
        if "|" in line:
            cells = [c.strip() for c in line.split("|")]
            cells = [c for c in cells if c or len(cells) <= 3]
        else:
            m = re.match(
                r"^(\d{1,4})[.、．]?\s+([\u4e00-\u9fff·]{2,8})\s*([男女])(?:\s+(.+))?$",
                line,
            )
            if m:
                cells = [m.group(1), m.group(2), m.group(3)]
                if m.group(4):
                    cells.extend(re.split(r"\s{2,}", m.group(4).strip()))
                flush()
                block.append(cells)
                if m.group(1).isdigit():
                    last_seq = int(m.group(1))
                continue
            cells = re.split(r"\s{2,}|\t", line)
            if len(cells) < 3:
                cells = line.split()
            cells = [c.strip() for c in cells if c.strip()]

        if not cells:
            flush()
            continue

        if _is_continuation_row(cells) and not re.fullmatch(r"\d+", cells[0]):
            last_seq += 1
            cells = [str(last_seq)] + cells
        elif cells and re.fullmatch(r"\d{1,4}", normalize_cell(cells[0])):
            last_seq = int(cells[0])

        if len(cells) >= 2 and not is_header_like_row(cells):
            block.append(cells)
        elif is_header_like_row(cells):
            flush()
            block.append(cells)
        else:
            flush()
    flush()
    return frames


def parse_id_anchor_lines(
    text: str,
    output_fields: Sequence[str],
) -> List[Dict[str, str]]:
    """行内含有身份证号时，从整行再扫一遍姓名/性别（补救列错位、缺序号行）。"""
    records: List[Dict[str, str]] = []
    seen_ids: set[str] = set()
    for line in text.splitlines():
        line = line.strip()
        if not line or is_header_like_row(_split_ocr_table_line(line)):
            continue
        cid = extract_id_card(line)
        if not cid or cid in seen_ids:
            continue
        athlete = {f: "" for f in output_fields}
        athlete["身份证号"] = cid
        if "男" in line and "女" not in line:
            athlete["性别"] = "男"
        elif "女" in line:
            athlete["性别"] = "女"
        for m in re.finditer(r"[\u4e00-\u9fff·]{2,8}", line):
            n = m.group()
            if is_plausible_person_name(n):
                athlete["姓名"] = n
                break
        if is_plausible_person_name(athlete.get("姓名", "")):
            seen_ids.add(cid)
            records.append(athlete)
    return records


def parse_text_as_tables(
    text: str,
    output_fields: Sequence[str],
    header_aliases: Dict[str, str],
) -> List[Dict[str, str]]:
    records: List[Dict[str, str]] = []
    for df in parse_text_lines_to_dataframe(text):
        records.extend(dataframe_to_records(df, output_fields, header_aliases))
    records.extend(parse_ocr_roster_lines(text, output_fields))
    records.extend(parse_id_anchor_lines(text, output_fields))
    return dedupe_records(records, ("姓名", "身份证号"))


def check_tesseract_available() -> Tuple[bool, str]:
    """返回 (是否可用, 说明)。"""
    try:
        import pytesseract
    except ImportError:
        return False, "未安装 Python 包 pytesseract（pip install pytesseract）"
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:
        return (
            False,
            "未检测到 Tesseract 可执行文件，请安装系统 Tesseract 并加入 PATH，"
            f"并安装 chi_sim 语言包。详情: {e}",
        )
    return True, f"Tesseract {pytesseract.get_tesseract_version()}"
