#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
字段定义与 CSV 合并。

- 公告 JSON（notice__*.json）：本地解析通知+附件
- 申报 JSON（submit__*.json）：本地解析申报表
- 合并为 final_athletes.csv
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Mapping

import pandas as pd


def configure_stdio_utf8() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        kernel32.SetConsoleOutputCP(65001)
        kernel32.SetConsoleCP(65001)
    except Exception:
        pass
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None or not hasattr(stream, "reconfigure"):
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


configure_stdio_utf8()

PROJECT_ROOT = Path(__file__).resolve().parent
TEMP_JSON_DIR = "temp_jsons"
OUTPUT_CSV = "final_athletes.csv"
MANUAL_FIX_DIR = "manual_fix"

SUBMIT_JSON_FIELDS = [
    "姓名",
    "性别",
    "项目",
    "运动等级",
    "证书授予单位",
    "高校",
    "运动等级证书编号",
]

# Kimi 回复中公示级 meta（只出现一次，再合并进每条运动员）
NOTICE_META_FIELDS = ["来源文件路径", "文件ID", "授予单位"]

NOTICE_JSON_FIELDS = [
    "来源文件路径",
    "文件ID",
    "授予单位",
    "文号",
    "发布日期",
    "运动员等级",
    "项目",
    "姓名",
    "性别",
    "身份证号",
    "证书编号",
    "赛事名称",
    "比赛成绩",
    "代表单位",
]

# 附件表内运动员列（公告元数据单独从正文正则提取）
NOTICE_ATHLETE_FIELDS = [
    "运动员等级",
    "项目",
    "姓名",
    "性别",
    "身份证号",
    "证书编号",
    "赛事名称",
    "比赛成绩",
    "代表单位",
]

CSV_FIELDS = [
    "来源文件路径",
    "文件ID",
    "授予单位",
    "文号",
    "发布日期",
    "运动员等级",
    "项目",
    "姓名",
    "性别",
    "身份证号",
    "证书编号",
    "赛事名称",
    "比赛成绩",
    "代表单位",
    "高校",
]


def empty_csv_row() -> Dict[str, str]:
    return {f: "" for f in CSV_FIELDS}


def notice_meta_for_source(source_url: str, file_id: str) -> Dict[str, str]:
    return {
        "来源文件路径": source_url,
        "文件ID": file_id,
        "授予单位": "",
        "文号": "",
        "发布日期": "",
    }


def build_notice_json_record(
    notice_meta: Dict[str, str],
    athlete: Dict[str, str],
) -> Dict[str, str]:
    row = {f: "" for f in NOTICE_JSON_FIELDS}
    for key, val in notice_meta.items():
        if key in row and val is not None:
            row[key] = str(val).strip()
    for key, val in athlete.items():
        if key in row and val is not None:
            row[key] = str(val).strip()
    return row


def build_submit_json_record(
    source_path: str,
    file_id: str,
    row: Dict[str, str],
) -> Dict[str, str]:
    out: Dict[str, str] = {f: "" for f in SUBMIT_JSON_FIELDS}
    for key in SUBMIT_JSON_FIELDS:
        if key in row and row[key] is not None:
            out[key] = str(row[key]).strip()
    out["来源文件路径"] = source_path
    out["文件ID"] = file_id
    return out


def _str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


def notice_json_to_csv_row(rec: Mapping[str, Any]) -> Dict[str, str]:
    row = empty_csv_row()
    for key in NOTICE_JSON_FIELDS:
        if key in rec:
            row[key] = _str(rec[key])
    if not row["项目"]:
        row["项目"] = _str(rec.get("项目类别", ""))
    return row


def submit_json_to_csv_row(rec: Mapping[str, Any]) -> Dict[str, str]:
    row = empty_csv_row()
    row["来源文件路径"] = _str(rec.get("来源文件路径", ""))
    row["文件ID"] = _str(rec.get("文件ID", ""))
    row["授予单位"] = _str(rec.get("证书授予单位", rec.get("授予单位", "")))
    row["运动员等级"] = _str(rec.get("运动等级", rec.get("运动员等级", "")))
    row["项目"] = _str(rec.get("项目", ""))
    row["姓名"] = _str(rec.get("姓名", ""))
    row["性别"] = _str(rec.get("性别", ""))
    row["证书编号"] = _str(
        rec.get("运动等级证书编号", rec.get("证书编号", ""))
    )
    row["高校"] = _str(rec.get("高校", ""))
    row["文号"] = _str(rec.get("文号", ""))
    row["发布日期"] = _str(rec.get("发布日期", ""))
    return row


def json_record_to_csv_row(rec: Mapping[str, Any], json_name: str) -> Dict[str, str]:
    if json_name.startswith("submit__"):
        return submit_json_to_csv_row(rec)
    return notice_json_to_csv_row(rec)


def merge_jsons_to_csv(
    temp_json_dir: Path | None = None,
    output_csv: Path | None = None,
) -> int:
    temp_dir = (temp_json_dir or (PROJECT_ROOT / TEMP_JSON_DIR)).resolve()
    out_path = (output_csv or (PROJECT_ROOT / OUTPUT_CSV)).resolve()

    csv_rows: List[Dict[str, str]] = []
    json_files = sorted(temp_dir.glob("*.json"))
    if not json_files:
        print("没有找到临时 JSON，请先运行抽取脚本。")
        return 0

    for json_path in json_files:
        with open(json_path, "r", encoding="utf-8") as f:
            records = json.load(f)
        if not isinstance(records, list):
            print(f"警告: {json_path.name} 不是列表，已跳过")
            continue
        for rec in records:
            if isinstance(rec, dict):
                csv_rows.append(json_record_to_csv_row(rec, json_path.stem))

    if not csv_rows:
        print("未找到任何运动员记录。")
        return 0

    df = pd.DataFrame(csv_rows)[CSV_FIELDS]
    before = len(df)
    df = df.drop_duplicates()
    removed = before - len(df)
    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    msg = f"合并完成：共 {len(df)} 条"
    if removed:
        msg += f"（去重 {removed} 条）"
    print(f"{msg}，已保存至 {out_path}")
    return len(df)
