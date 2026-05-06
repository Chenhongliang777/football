#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
运动员技术等级公告附件批量下载器（API 直连 + RequestId 版）
基于 requests + BeautifulSoup

功能：
1. 自动在请求头中生成 RequestId，POST JSON 到公告详情 API
2. 对每个 id，先 source=2，若无附件再 source=1
3. 多附件命名：id_1.pdf, id_2.docx；单附件命名：id.pdf
4. 两个 source 都无附件的 id 记录到 missing_ids.json
5. 断点续传：progress.json 记录已完成的 id，随时中断再运行自动跳过

依赖安装：
    pip install requests beautifulsoup4 lxml

使用方式：
    python notice_attachment_crawler.py
"""

import json
import os
import sys
import time
import random
import re
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

# ==================== 配置区域 ====================
INPUT_FILE = "ids.json"
OUTPUT_DIR = "data"
PROGRESS_FILE = "data/progress.json"
MISSING_FILE = "data/missing_ids.json"

NOTICE_API = "https://ydydj.univsport.com/api/system/document/public-grant-notice-detail"
INIT_URL = "https://ydydj.univsport.com/index.php?m=index&c=look&a=look"
TIMEOUT = 15

# 基础 Headers（每次请求会动态加入 Requestid）
BASE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Content-Type": "application/json; charset=UTF-8",
    "Origin": "https://ydydj.univsport.com",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "Connection": "keep-alive",
}

ATTACHMENT_EXTS = (
    ".pdf", ".doc", ".docx", ".xls", ".xlsx",
    ".zip", ".rar", ".7z", ".txt", ".wps", ".et", ".dps"
)
ATTACHMENT_KEYWORDS = ("附件", "下载", "点击下载", "document", "file", "上传")
# ==================================================


def ensure_dir(path: str) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def read_ids(filepath: str) -> list:
    p = Path(filepath)
    if not p.exists():
        return None
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read().strip()
    try:
        data = json.loads(content)
        if isinstance(data, list):
            return [str(x) for x in data]
    except (json.JSONDecodeError, ValueError):
        pass
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    return lines


def load_progress() -> dict:
    p = Path(PROGRESS_FILE)
    if p.exists():
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"completed": [], "missing": []}


def save_progress(progress: dict):
    ensure_dir(OUTPUT_DIR)
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)
    with open(MISSING_FILE, "w", encoding="utf-8") as f:
        json.dump(progress.get("missing", []), f, ensure_ascii=False, indent=2)


def gen_request_id() -> str:
    """
    生成 RequestId（模仿浏览器行为：12位数字字符串）。
    样本特征：01 开头 + 10位随机/时间戳数字。
    """
    # 01 前缀 + 10位随机数字
    return "01" + "".join(str(random.randint(0, 9)) for _ in range(10))


def get_ext_from_url(url: str) -> str:
    path = url.split("?")[0]
    ext = Path(path).suffix.lower()
    if ext in ATTACHMENT_EXTS:
        return ext
    return ""


def infer_ext_from_response(resp) -> str:
    ct = resp.headers.get("Content-Type", "").lower()
    mapping = {
        "pdf": ".pdf", "msword": ".doc", "wordprocessingml": ".docx",
        "excel": ".xls", "spreadsheetml": ".xlsx",
        "zip": ".zip", "rar": ".rar", "plain": ".txt",
    }
    for key, ext in mapping.items():
        if key in ct:
            return ext
    cd = resp.headers.get("Content-Disposition", "")
    fn_match = re.findall(r'filename[^;=]*=["\']?([^"\';]+)["\']?', cd)
    if fn_match:
        ext = Path(fn_match[0]).suffix.lower()
        if ext in ATTACHMENT_EXTS:
            return ext
    return ".bin"


def find_attachments_in_html(html_content: str, base_url: str) -> list:
    if not html_content:
        return []
    soup = BeautifulSoup(html_content, "lxml")
    attachments = []
    seen = set()

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        text = a.get_text(strip=True)
        title = a.get("title", "")
        onclick = a.get("onclick", "")

        if not href or href.startswith(("javascript:", "mailto:", "#", "void")):
            if not any(k in onclick.lower() for k in ("download", "file", "down", "export")):
                continue

        full_url = urljoin(base_url, href)

        is_attach = False
        if any(href.lower().endswith(ext) for ext in ATTACHMENT_EXTS):
            is_attach = True
        if not is_attach:
            for ext in ATTACHMENT_EXTS:
                if ext in text.lower() or ext in title.lower():
                    is_attach = True
                    break
        if not is_attach:
            if any(k in text or k in title for k in ATTACHMENT_KEYWORDS):
                is_attach = True

        if is_attach and full_url not in seen:
            seen.add(full_url)
            attachments.append((full_url, text or title))

    return attachments


def download_file(session: requests.Session, url: str, save_path: Path) -> Path:
    resp = session.get(url, stream=True, timeout=30, headers=BASE_HEADERS)
    resp.raise_for_status()
    if save_path.suffix == ".bin" or not save_path.suffix:
        inferred = infer_ext_from_response(resp)
        if inferred and inferred != ".bin":
            save_path = save_path.with_suffix(inferred)
    with open(save_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)
    return save_path


def fetch_notice_api(session: requests.Session, nid: str, source: str) -> tuple:
    """
    POST JSON 到公告详情 API，返回 (attachments, error)。
    每次请求自动生成新的 Requestid 放入 Header。
    """
    payload = {
        "identity": "3",
        "source": source,
        "id": nid,
    }

    # 动态生成请求头（每次不同 Requestid + 对应 Referer）
    headers = dict(BASE_HEADERS)
    headers["Requestid"] = gen_request_id()
    headers["Referer"] = f"https://ydydj.univsport.com/level/noticedetail?identity=3&source={source}&id={nid}"

    try:
        resp = session.post(NOTICE_API, json=payload, headers=headers, timeout=TIMEOUT)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        data = resp.json()

        if not isinstance(data, dict):
            return [], f"返回非字典: {type(data)}"
        if data.get("code") != 0 and data.get("code") != "0":
            return [], f"API 错误: code={data.get('code')}, msg={data.get('msg')}"

        inner = data.get("data", {})
        if not isinstance(inner, dict):
            return [], "data 字段非字典"

        content_html = inner.get("content", "")
        if not content_html:
            return [], None

        attachments = find_attachments_in_html(content_html, "https://ydydj.univsport.com/")
        return attachments, None

    except requests.exceptions.RequestException as e:
        return [], f"网络请求失败: {e}"
    except json.JSONDecodeError as e:
        return [], f"JSON 解析失败: {e}"
    except Exception as e:
        return [], f"异常: {e}"


def download_attachments(session, attachments, nid):
    for idx, (aurl, atext) in enumerate(attachments, start=1):
        ext = get_ext_from_url(aurl)
        if not ext:
            ext = ".bin"
        filename = f"{nid}{ext}" if len(attachments) == 1 else f"{nid}_{idx}{ext}"
        save_path = Path(OUTPUT_DIR) / filename
        try:
            final_path = download_file(session, aurl, save_path)
            print(f"  下载成功: {final_path.name} ({atext[:20]})")
        except Exception as e:
            print(f"  [下载失败] {filename}: {e}")


def process_id(session: requests.Session, nid: str, progress: dict) -> dict:
    print(f"\n[{nid}] 开始处理...")

    attachments2, err2 = fetch_notice_api(session, nid, "2")
    if err2:
        print(f"  [source=2 API 错误] {err2}")
        return progress

    if attachments2:
        print(f"  source=2 发现 {len(attachments2)} 个附件")
        download_attachments(session, attachments2, nid)
        if nid not in progress["completed"]:
            progress["completed"].append(nid)
        return progress

    print(f"  source=2 无附件，尝试 source=1...")

    attachments1, err1 = fetch_notice_api(session, nid, "1")
    if err1:
        print(f"  [source=1 API 错误] {err1}")
        return progress

    if attachments1:
        print(f"  source=1 发现 {len(attachments1)} 个附件")
        download_attachments(session, attachments1, nid)
    else:
        print(f"  source=1 也无附件")
        if nid not in progress["missing"]:
            progress["missing"].append(nid)

    if nid not in progress["completed"]:
        progress["completed"].append(nid)
    return progress


def main():
    ensure_dir(OUTPUT_DIR)

    ids = read_ids(INPUT_FILE)
    if ids is None:
        ids = [str(i) for i in range(1, 1905)]
        with open(INPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(ids, f, ensure_ascii=False, indent=2)
        print(f"已创建示例输入文件: {INPUT_FILE}（范围 1~1904）")

    print(f"共读取 {len(ids)} 个待处理 id")

    progress = load_progress()
    completed_set = set(progress["completed"])
    print(f"已完成的 id: {len(completed_set)}，剩余: {len(ids) - len(completed_set)}")

    session = requests.Session()
    session.headers.update(BASE_HEADERS)

    # 初始化 Cookie
    try:
        print(f"初始化会话: {INIT_URL}")
        r = session.get(INIT_URL, timeout=TIMEOUT)
        print(f"初始化状态: {r.status_code}")
    except Exception as e:
        print(f"[警告] 初始化失败: {e}")

    total_start = time.perf_counter()

    for i, nid in enumerate(ids, start=1):
        if nid in completed_set:
            continue

        progress = process_id(session, nid, progress)
        save_progress(progress)

        if i % 10 == 0:
            print(f"\n--- 进度: {i}/{len(ids)} | 已完成: {len(progress['completed'])} | 无附件: {len(progress['missing'])} ---")

    total_elapsed = time.perf_counter() - total_start
    avg = len(ids) / total_elapsed if total_elapsed > 0 else 0
    print(f"\n 全部完成！")
    print(f"总耗时: {total_elapsed:.2f}s | 平均速度: {avg:.2f} 条/秒")
    print(f"附件保存目录: {Path(OUTPUT_DIR).resolve()}")
    print(f"无附件列表: {Path(MISSING_FILE).resolve()}")


if __name__ == "__main__":
    main()
