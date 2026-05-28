#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""申报推荐表本地解析（Submit_athletes），无 LLM。"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Dict, List

import pandas as pd

from extraction_common import (
    PROJECT_ROOT,
    SUBMIT_JSON_FIELDS,
    TEMP_JSON_DIR,
    build_submit_json_record,
    merge_jsons_to_csv,
)
from table_parse_core import dataframe_to_records, dedupe_records

SUBMIT_DIR = PROJECT_ROOT / "Submit_athletes"
TEMP_DIR = PROJECT_ROOT / TEMP_JSON_DIR

SUBMIT_HEADER_ALIASES: Dict[str, str] = {
    "姓名": "姓名",
    "性别": "性别",
    "项目": "项目",
    "运动等级": "运动等级",
    "运动员等级": "运动等级",
    "证书授予单位": "证书授予单位",
    "授予单位": "证书授予单位",
    "高校": "高校",
    "母校": "高校",
    "学校": "高校",
    "就读学校": "高校",
    "运动等级证书编号": "运动等级证书编号",
    "运动员等级证书编号": "运动等级证书编号",
    "证书编号": "运动等级证书编号",
}


def safe_name(name: str) -> str:
    return re.sub(r"[^\w\-.]", "_", name)


def parse_xlsx(path: Path) -> List[Dict[str, str]]:
    records: List[Dict[str, str]] = []
    for sheet in pd.ExcelFile(path).sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
        records.extend(
            dataframe_to_records(df, SUBMIT_JSON_FIELDS, SUBMIT_HEADER_ALIASES)
        )
    return dedupe_records(records, ("姓名", "运动等级证书编号"))


def parse_pdf_tables(path: Path) -> List[Dict[str, str]]:
    import pdfplumber

    records: List[Dict[str, str]] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if table and len(table) >= 2:
                    records.extend(
                        dataframe_to_records(
                            pd.DataFrame(table),
                            SUBMIT_JSON_FIELDS,
                            SUBMIT_HEADER_ALIASES,
                        )
                    )
    return dedupe_records(records, ("姓名", "运动等级证书编号"))


def parse_submit_file(path: Path) -> List[Dict[str, str]]:
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        return parse_xlsx(path)
    if ext == ".pdf":
        return parse_pdf_tables(path)
    return []


def process_submit_file(path: Path, temp_dir: Path, skip_existing: bool) -> None:
    file_id = path.stem
    out_path = temp_dir / f"submit__{safe_name(file_id)}.json"

    if skip_existing and out_path.is_file():
        try:
            if json.loads(out_path.read_text(encoding="utf-8")):
                print(f"跳过: {path.name}")
                return
        except json.JSONDecodeError:
            pass

    print(f"处理: {path.name}", flush=True)
    athletes = parse_submit_file(path)
    records = [
        build_submit_json_record(str(path.resolve()), file_id, a) for a in athletes
    ]
    out_path.write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"  {len(records)} 条 -> {out_path.name}", flush=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="申报表本地解析")
    parser.add_argument("--skip-existing", action="store_true")
    parser.add_argument("--merge-only", action="store_true")
    parser.add_argument("--limit", type=int, default=None, metavar="N")
    args = parser.parse_args()

    temp_dir = TEMP_DIR
    temp_dir.mkdir(parents=True, exist_ok=True)

    if args.merge_only:
        merge_jsons_to_csv()
        return

    if not SUBMIT_DIR.is_dir():
        print(f"目录不存在: {SUBMIT_DIR}")
        return

    files = sorted(
        [
            p
            for p in SUBMIT_DIR.iterdir()
            if p.is_file() and p.suffix.lower() in (".pdf", ".xlsx", ".xls")
        ],
        key=lambda p: p.name,
    )
    if args.limit and args.limit > 0:
        files = files[: args.limit]

    if not files:
        print("未找到 PDF/xlsx")
        merge_jsons_to_csv()
        return

    print(f"共 {len(files)} 个申报表")
    for i, p in enumerate(files, 1):
        print(f"[{i}/{len(files)}] ", end="")
        process_submit_file(p, temp_dir, args.skip_existing)

    merge_jsons_to_csv()


if __name__ == "__main__":
    main()
