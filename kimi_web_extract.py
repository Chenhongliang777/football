#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通过 Kimi 网页（Playwright + 本机 Chrome）批量解析公告包附件，输出 notice__<ID>.json。

首次使用请先登录（会打开 Chrome 并保存会话）：
  python kimi_web_extract.py --login

单条调试（显示浏览器 + 保存 kimi_debug）：
  python kimi_web_extract.py --ids 910 --headed --debug

批量（908–1904，默认无头、不写 debug，断点续跑）：
  python kimi_web_extract.py --from-id 908 --to-id 1904
  python kimi_web_extract.py --from-id 908 --to-id 1904 --retry-failed

与本地解析配合：
  python submit_table_extract.py
  python kimi_web_extract.py --merge-only
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

from extraction_common import (
    NOTICE_ATHLETE_FIELDS,
    NOTICE_JSON_FIELDS,
    NOTICE_META_FIELDS,
    PROJECT_ROOT,
    TEMP_JSON_DIR,
    merge_jsons_to_csv,
)

NOTICE_DATA_DIR = PROJECT_ROOT / "notice_attachment_crawler" / "data"
PROMPT_PATH = PROJECT_ROOT / "prompts" / "notice_kimi_prompt.txt"
PROFILE_DIR = PROJECT_ROOT / ".kimi_chrome_profile"
FAIL_LOG = PROJECT_ROOT / "kimi_agent_failures.txt"
FAILED_IDS_FILE = PROJECT_ROOT / "kimi_extract_failed_ids.txt"
# 首页即快速模型（与手动一致）；仅当误入 /agent 时再跳回
KIMI_HOME = "https://www.kimi.com/"

UPLOAD_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".docx", ".doc", ".txt"}

DEFAULT_CHROME_PATHS = (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
)


def _sort_key_numeric(path: Path) -> tuple:
    name = path.name
    return (0, int(name)) if name.isdigit() else (1, name)


def resolve_chrome_executable() -> Optional[str]:
    for key in ("KIMI_CHROME_PATH", "CHROME_PATH", "GOOGLE_CHROME_SHIM"):
        val = os.environ.get(key, "").strip().strip('"')
        if val and Path(val).is_file():
            return val
    for p in DEFAULT_CHROME_PATHS:
        if Path(p).is_file():
            return p
    return None


def list_notice_folders(
    *,
    limit: Optional[int] = None,
    ids: Optional[Set[str]] = None,
    from_id: Optional[int] = None,
    to_id: Optional[int] = None,
) -> List[Path]:
    if not NOTICE_DATA_DIR.is_dir():
        return []
    folders = [
        p
        for p in sorted(NOTICE_DATA_DIR.iterdir(), key=_sort_key_numeric)
        if p.is_dir() and p.name.isdigit()
    ]
    if from_id is not None or to_id is not None:
        lo = from_id if from_id is not None else 0
        hi = to_id if to_id is not None else 10**9
        folders = [p for p in folders if lo <= int(p.name) <= hi]
    if ids:
        folders = [p for p in folders if p.name in ids]
    if limit is not None and limit > 0:
        folders = folders[:limit]
    return folders


def read_source_url(folder: Path) -> str:
    url_file = folder / "url.txt"
    if url_file.is_file():
        return url_file.read_text(encoding="utf-8").strip()
    return ""


def collect_upload_files(folder: Path) -> List[Path]:
    files: List[Path] = []
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.name.startswith("."):
            continue
        if p.suffix.lower() not in UPLOAD_EXTENSIONS:
            continue
        files.append(p)
    return files


def build_prompt(file_id: str, folder: Path) -> str:
    template = PROMPT_PATH.read_text(encoding="utf-8")
    return template.format(file_id=file_id)


def _sanitize_for_json(text: str) -> str:
    return (
        text.replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2018", "'")
        .replace("\u2019", "'")
    )


def parse_agent_response(
    text: str,
) -> Tuple[List[Dict[str, Any]], Optional[int], Dict[str, str]]:
    """解析 meta 对象 + 运动员 JSON 数组（数组以 [ 开头）+ COUNT。"""
    count: Optional[int] = None
    m = re.search(r"COUNT\s*=\s*(\d+)", text, re.IGNORECASE)
    if m:
        count = int(m.group(1))

    text = _sanitize_for_json(text)
    meta: Dict[str, str] = {f: "" for f in NOTICE_META_FIELDS}
    best: List[Dict[str, Any]] = []
    decoder = json.JSONDecoder()

    for match in re.finditer(r"\[", text):
        start = match.start()
        try:
            obj, _ = decoder.raw_decode(text, start)
        except json.JSONDecodeError:
            continue
        if not isinstance(obj, list):
            continue
        athletes = [
            x for x in obj if isinstance(x, dict) and (x.get("姓名") or "").strip()
        ]
        if len(athletes) > len(best):
            best = athletes

    got_array = bool(best)

    for match in re.finditer(r"\{", text):
        start = match.start()
        try:
            obj, _ = decoder.raw_decode(text, start)
        except json.JSONDecodeError:
            continue
        if not isinstance(obj, dict):
            continue
        if any(k in obj for k in NOTICE_META_FIELDS):
            for k in NOTICE_META_FIELDS:
                if k in obj and obj[k] is not None:
                    meta[k] = str(obj[k]).strip()
        if got_array:
            continue
        if (obj.get("姓名") or "").strip():
            best.append(obj)

    return best, count, meta


def normalize_records(
    records: List[Dict[str, Any]],
    file_id: str,
    source_url: str,
    meta: Optional[Dict[str, str]] = None,
) -> List[Dict[str, str]]:
    """合并公示级 meta + 本地 url.txt，写入每条运动员（与 final CSV 字段一致）。"""
    meta = meta or {}
    fid = (meta.get("文件ID") or file_id or "").strip()
    src = (meta.get("来源文件路径") or source_url or "").strip()
    unit = (meta.get("授予单位") or "").strip()

    out: List[Dict[str, str]] = []
    for rec in records:
        if not isinstance(rec, dict):
            continue
        row = {f: "" for f in NOTICE_JSON_FIELDS}
        for k in NOTICE_ATHLETE_FIELDS:
            if k in rec and rec[k] is not None:
                row[k] = str(rec[k]).strip()
        for k, v in rec.items():
            if k in row and v is not None and not row[k]:
                row[k] = str(v).strip()
        row["文件ID"] = fid
        row["来源文件路径"] = src
        row["授予单位"] = unit
        if row.get("姓名"):
            out.append(row)
    return out


def output_json_path(file_id: str) -> Path:
    return PROJECT_ROOT / TEMP_JSON_DIR / f"notice__{file_id}.json"


def existing_ok(path: Path, min_rows: int = 1) -> bool:
    if not path.is_file():
        return False
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return False
    return isinstance(data, list) and len(data) >= min_rows


def append_fail_log(file_id: str, reason: str) -> None:
    line = f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{file_id}\t{reason}\n"
    with FAIL_LOG.open("a", encoding="utf-8") as f:
        f.write(line)


def read_failed_ids() -> Set[str]:
    """待重跑 ID 列表（一行一个）。"""
    if not FAILED_IDS_FILE.is_file():
        return set()
    out: Set[str] = set()
    for line in FAILED_IDS_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            out.add(line.split()[0])
    return out


def write_failed_ids(ids: Set[str]) -> None:
    FAILED_IDS_FILE.parent.mkdir(parents=True, exist_ok=True)
    ordered = sorted(ids, key=lambda x: (0, int(x)) if x.isdigit() else (1, x))
    FAILED_IDS_FILE.write_text(
        "\n".join(ordered) + ("\n" if ordered else ""),
        encoding="utf-8",
    )


def _failed_ids_from_log() -> Set[str]:
    if not FAIL_LOG.is_file():
        return set()
    out: Set[str] = set()
    for line in FAIL_LOG.read_text(encoding="utf-8").splitlines():
        parts = line.split("\t")
        if len(parts) >= 2:
            fid = parts[1].strip()
            if fid.isdigit():
                out.add(fid)
    return out


def sync_failed_ids_registry() -> Set[str]:
    """合并失败日志中的 ID，并去掉已有成功 JSON 的项。"""
    ids = read_failed_ids() | _failed_ids_from_log()
    for fid in list(ids):
        if existing_ok(output_json_path(fid)):
            ids.discard(fid)
    write_failed_ids(ids)
    return ids


def register_failure(file_id: str, reason: str) -> None:
    append_fail_log(file_id, reason)
    ids = read_failed_ids()
    ids.add(file_id)
    write_failed_ids(ids)


def register_success(file_id: str) -> None:
    ids = read_failed_ids()
    if file_id in ids:
        ids.discard(file_id)
        write_failed_ids(ids)


def remove_invalid_output(file_id: str) -> None:
    out = output_json_path(file_id)
    if out.is_file() and not existing_ok(out):
        try:
            out.unlink()
        except OSError:
            pass


def plan_batch_folders(
    *,
    from_id: Optional[int],
    to_id: Optional[int],
    ids: Optional[Set[str]],
    limit: Optional[int],
    retry_failed: bool,
) -> List[Path]:
    folders = list_notice_folders(
        limit=limit, ids=ids, from_id=from_id, to_id=to_id
    )
    if retry_failed:
        failed = sync_failed_ids_registry()
        if not failed:
            print("无待重跑 ID（kimi_extract_failed_ids.txt 为空）", flush=True)
            return []
        folders = [p for p in folders if p.name in failed]
        print(f"--retry-failed: 范围内待重跑 {len(folders)} 个", flush=True)
    return folders


def print_batch_plan(
    folders: Sequence[Path],
    *,
    skip_existing: bool,
    failed_ids: Set[str],
) -> Tuple[int, int]:
    """返回 (将跳过数, 将处理数)。"""
    skip_n = 0
    for folder in folders:
        if skip_existing and existing_ok(output_json_path(folder.name)):
            skip_n += 1
    run_n = len(folders) - skip_n
    in_fail = sum(1 for p in folders if p.name in failed_ids)
    print(
        f"批处理计划: 共 {len(folders)} 个文件夹，"
        f"跳过已成功 {skip_n} 个，待执行 {run_n} 个"
        f"（历史失败登记 {in_fail} 个）",
        flush=True,
    )
    return skip_n, run_n


def launch_browser(playwright, *, headless: bool):
    """使用本机 Chrome（不下载 Playwright 自带 Chromium）。"""
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    chrome = resolve_chrome_executable()
    kwargs: Dict[str, Any] = {
        "headless": headless,
        "locale": "zh-CN",
        "accept_downloads": True,
        "viewport": {"width": 1400, "height": 900},
    }
    if chrome:
        kwargs["executable_path"] = chrome
        print(f"Chrome: {chrome}", flush=True)
    else:
        kwargs["channel"] = "chrome"
        print("Chrome: channel=chrome（系统已安装）", flush=True)

    return playwright.chromium.launch_persistent_context(
        str(PROFILE_DIR),
        **kwargs,
    )


def dump_ui_candidates(page, out_path: Path) -> None:
    """列出页面上可能与聊天/上传相关的控件，供人工对照或发给开发者。"""
    script = r"""
() => {
  const visible = (el) => {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none';
  };
  const short = (s, n = 120) => (s && s.length > n ? s.slice(0, n) + '…' : s || '');
  const describe = (el) => {
    const r = el.getBoundingClientRect();
    return {
      tag: el.tagName.toLowerCase(),
      type: el.getAttribute('type') || '',
      role: el.getAttribute('role') || '',
      ariaLabel: el.getAttribute('aria-label') || '',
      placeholder: el.getAttribute('placeholder') || '',
      name: el.getAttribute('name') || '',
      id: el.id || '',
      className: short(el.className?.toString() || '', 80),
      text: short((el.innerText || el.textContent || '').trim().replace(/\s+/g, ' '), 40),
      rect: { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height) },
    };
  };
  const selectors = [
    'textarea',
    '[contenteditable="true"]',
    '[role="textbox"]',
    'input[type="file"]',
    'button',
    '[role="button"]',
    'a[href]',
    '[aria-label]',
    '[class*="upload" i]',
    '[class*="attach" i]',
    '[class*="composer" i]',
    '[class*="input" i]',
    '[class*="chat" i]',
  ];
  const seen = new Set();
  const items = [];
  for (const sel of selectors) {
    for (const el of document.querySelectorAll(sel)) {
      if (seen.has(el)) continue;
      seen.add(el);
      if (!visible(el) && el.tagName !== 'INPUT') continue;
      items.push({ selectorHint: sel, ...describe(el) });
    }
  }
  items.sort((a, b) => a.rect.y - b.rect.y || a.rect.x - b.rect.x);
  return {
    url: location.href,
    title: document.title,
    viewport: { w: innerWidth, h: innerHeight },
    count: items.length,
    items,
  };
}
"""
    data = page.evaluate(script)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已写入控件清单: {out_path}", flush=True)
    print("请把该文件内容（或其中 textarea / file / 上传 相关几段）发给我。", flush=True)


def leave_agent_page_if_needed(page) -> None:
    """仅当 URL 在 /agent 时回到首页；不重复跳转 /chat。"""
    if "/agent" in page.url.lower():
        page.goto(KIMI_HOME, wait_until="domcontentloaded", timeout=120_000)
        page.wait_for_timeout(1500)


def _sidebar_is_expanded(page) -> bool:
    """侧栏展开时 #chat-box 会明显右移（比宽侧栏占用空间）。"""
    return bool(
        page.evaluate(
            """
            () => {
              const box = document.querySelector('#chat-box');
              if (!box) return false;
              return box.getBoundingClientRect().x > 300;
            }
            """
        )
    )


def ensure_sidebar_collapsed(page) -> None:
    """每个会话只尝试收起侧栏一次，避免重复乱点。"""
    if getattr(page, "_kimi_sidebar_handled", False):
        return
    setattr(page, "_kimi_sidebar_handled", True)
    if not _sidebar_is_expanded(page):
        return
    print("  侧栏展开，收起一次…", flush=True)
    for sel in (
        '[aria-label*="收起"]',
        '[aria-label*="折叠"]',
        '[class*="sider-trigger"]',
        '[class*="sidebar-toggle"]',
    ):
        try:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible():
                loc.click(timeout=3000)
                page.wait_for_timeout(800)
                break
        except Exception:
            continue
    if not _sidebar_is_expanded(page):
        print("  侧栏已收起", flush=True)
    else:
        print("  侧栏仍展开（不再重复点击，请手动收起）", flush=True)


def wait_for_composer(page) -> None:
    """等待中央输入区（ui_candidates: #chat-box）。"""
    page.locator("#chat-box").wait_for(state="visible", timeout=60_000)
    page.locator("#chat-box .chat-input-editor").wait_for(state="visible", timeout=30_000)


def open_new_chat(page) -> None:
    setattr(page, "_kimi_sidebar_handled", False)
    page.goto(KIMI_HOME, wait_until="domcontentloaded", timeout=120_000)
    page.wait_for_timeout(2000)
    ensure_sidebar_collapsed(page)
    try:
        page.locator("a.new-chat-btn").first.click(timeout=5000)
        page.wait_for_timeout(1200)
    except Exception:
        pass
    leave_agent_page_if_needed(page)
    wait_for_composer(page)
    if "/agent" in page.url.lower():
        raise RuntimeError(
            f"当前仍在 Agent 页: {page.url}，请从侧栏回到首页 {KIMI_HOME} 后重试"
        )


def _set_files_on_hidden_input(page, str_paths: List[str]) -> bool:
    """ui_candidates: input.hidden-input（点「+」后出现，0×0 不可见但可 set_input_files）。"""
    for sel in ("#app input.hidden-input", "#app input[type='file']"):
        inputs = page.locator(sel)
        n = inputs.count()
        if n == 0:
            continue
        inputs.nth(n - 1).set_input_files(str_paths)
        return True
    return False


def _upload_menu_visible(page) -> bool:
    for loc in (
        page.get_by_text("文件和图片", exact=True),
        page.get_by_text("常用语", exact=True),
        page.locator(".v-binder-follower-container").get_by_text("文件和图片"),
    ):
        try:
            if loc.first.is_visible():
                return True
        except Exception:
            continue
    return False


_PLUS_FIND_JS = r"""
() => {
  const box = document.querySelector('#chat-box');
  if (!box) return null;
  const br = box.getBoundingClientRect();
  let best = null, minX = 1e9;
  for (const el of box.querySelectorAll('button, [role="button"], div, span')) {
    const t = (el.innerText || el.textContent || '').trim();
    if (/Agent|K2|快速|文件|图片/i.test(t)) continue;
    const r = el.getBoundingClientRect();
    if (r.width < 14 || r.width > 56 || r.height < 14) continue;
    if (r.x < br.x + 8 || r.x > br.x + br.width * 0.42) continue;
    if (r.y < br.y + br.height * 0.45) continue;
    if (r.x < minX) { minX = r.x; best = el; }
  }
  return best;
}
"""


def _click_element_handle(page, element_handle) -> bool:
    try:
        el = element_handle.as_element()
        if el:
            el.scroll_into_view_if_needed()
            el.click(timeout=5000)
            return True
    except Exception:
        pass
    return False


def _click_plus_button(page) -> bool:
    """仅在需要时点击 #chat-box 内「+」。"""
    page.locator("#chat-box").first.scroll_into_view_if_needed(timeout=5000)
    page.wait_for_timeout(300)
    handle = page.evaluate_handle(_PLUS_FIND_JS)
    if handle and _click_element_handle(page, handle):
        page.wait_for_timeout(900)
        return _upload_menu_visible(page)
    return False


def _attachments_visible(page, paths: Sequence[Path]) -> bool:
    """确认页面上出现附件名（避免误用上次残留的 hidden-input）。"""
    if not paths:
        return False
    found = 0
    for p in paths:
        name = p.name
        for probe in (name, name[:20], p.stem[:20]):
            if not probe:
                continue
            try:
                if page.get_by_text(probe, exact=False).first.is_visible():
                    found += 1
                    break
            except Exception:
                continue
    return found >= max(1, len(paths) // 2)


def _click_files_and_images(page) -> bool:
    """「+」弹出菜单 →「文件和图片」（菜单可能在浮层里）。"""
    candidates = (
        page.locator(".v-binder-follower-container").get_by_text("文件和图片", exact=True),
        page.get_by_text("文件和图片", exact=True),
    )
    for loc in candidates:
        try:
            item = loc.first
            item.wait_for(state="visible", timeout=6000)
            item.click(timeout=5000)
            page.wait_for_timeout(500)
            return True
        except Exception:
            continue
    return False


def _debug_shot(page, debug_dir: Optional[Path], name: str) -> None:
    if not debug_dir:
        return
    debug_dir.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(debug_dir / name), full_page=True)
    except Exception:
        pass


def upload_files(
    page, paths: Sequence[Path], *, debug_dir: Optional[Path] = None
) -> None:
    if not paths:
        return
    str_paths = [str(p.resolve()) for p in paths]
    _debug_shot(page, debug_dir, "upload_01_composer.png")

    if _set_files_on_hidden_input(page, str_paths):
        page.wait_for_timeout(2000 + 500 * len(str_paths))
        if _attachments_visible(page, paths):
            print(f"  附件已上传（hidden-input，未点 +），共 {len(paths)} 个", flush=True)
            _debug_shot(page, debug_dir, "upload_03_done.png")
            return

    print("  点击 + 打开上传菜单…", flush=True)
    if not _click_plus_button(page):
        _debug_shot(page, debug_dir, "upload_02_plus_failed.png")
        raise RuntimeError(
            "未点开「+」菜单：请加 --debug --headed 查看 kimi_debug/upload_02_plus_failed.png，"
            "或运行 --inspect-ui 校准 + 位置"
        )
    _debug_shot(page, debug_dir, "upload_02_menu_open.png")

    if _set_files_on_hidden_input(page, str_paths):
        page.wait_for_timeout(2000 + 500 * len(str_paths))
        if _attachments_visible(page, paths):
            print(f"  附件已出现在页面（{len(paths)} 个）", flush=True)
            _debug_shot(page, debug_dir, "upload_03_done.png")
            return

    try:
        with page.expect_file_chooser(timeout=12_000) as fc_info:
            if not _click_files_and_images(page):
                _debug_shot(page, debug_dir, "upload_03_no_menu_item.png")
                raise RuntimeError("未找到菜单项「文件和图片」")
        fc_info.value.set_files(str_paths)
        page.wait_for_timeout(2000 + 500 * len(str_paths))
        if _attachments_visible(page, paths):
            print(f"  附件已出现在页面（{len(paths)} 个）", flush=True)
            _debug_shot(page, debug_dir, "upload_03_done.png")
            return
    except RuntimeError:
        raise
    except Exception:
        pass

    if not _click_files_and_images(page):
        _debug_shot(page, debug_dir, "upload_03_no_menu_item.png")
        raise RuntimeError("未找到菜单项「文件和图片」")
    page.wait_for_timeout(500)
    if _set_files_on_hidden_input(page, str_paths):
        page.wait_for_timeout(2000 + 500 * len(str_paths))
        if _attachments_visible(page, paths):
            print(f"  附件已出现在页面（{len(paths)} 个）", flush=True)
            _debug_shot(page, debug_dir, "upload_03_done.png")
            return

    _debug_shot(page, debug_dir, "upload_03_input_failed.png")
    raise RuntimeError(
        "上传失败：已点「+」和「文件和图片」，但未写入文件；"
        f"当前 URL: {page.url}"
    )


def _ensure_agent_toggle_off(page) -> None:
    """尽量关闭 @ Agent，避免走代码/工具循环（非 /agent 页面）。"""
    for loc in (
        page.locator("#chat-box .chat-editor-action").get_by_text("Agent"),
        page.locator("#chat-box").get_by_text("@ Agent"),
    ):
        try:
            loc = loc.first
            if not loc.count() or not loc.is_visible():
                continue
            cls = loc.evaluate(
                "(el) => (el.className && el.className.toString()) || ''"
            )
            aria = loc.get_attribute("aria-pressed") or ""
            if re.search(r"active|selected|checked|on", cls, re.I) or aria == "true":
                loc.click(timeout=3000)
                page.wait_for_timeout(500)
        except Exception:
            continue


def fill_prompt(page, prompt: str) -> None:
    editor = page.locator("#chat-box .chat-input-editor").first
    editor.wait_for(state="visible", timeout=15_000)
    editor.scroll_into_view_if_needed(timeout=5000)
    editor.click(timeout=5000)
    try:
        editor.fill(prompt, timeout=15_000)
        return
    except Exception:
        page.keyboard.press("Control+a")
        page.keyboard.type(prompt, delay=2)


_SEND_ARROW_JS = r"""
() => {
  const lum = (rgb) => {
    const m = (rgb || '').match(/[\d.]+/g);
    if (!m || m.length < 3) return 0.5;
    const [r, g, b] = m.map(Number);
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  };
  const bar = document.querySelector('#chat-box .chat-editor-action')
    || document.querySelector('#chat-box [class*="editor-action"]');
  if (!bar) return { state: 'missing', x: 0, y: 0 };
  const ar = bar.getBoundingClientRect();
  const minX = ar.x + ar.width * 0.65;
  let best = null, maxX = -1;
  const kids = bar.querySelectorAll(':scope > *');
  const pick = kids.length ? kids[kids.length - 1] : null;
  if (pick) {
    const pt = (pick.innerText || pick.textContent || '').trim();
    if (/K2|快速/i.test(pt) && kids.length >= 2) best = kids[kids.length - 2];
    else if (!/Agent/i.test(pt)) best = pick;
  }
  if (!best) {
    for (const el of bar.querySelectorAll(':scope > *')) {
      const t = (el.innerText || el.textContent || '').trim();
      if (/Agent|K2|快速/i.test(t)) continue;
      const r = el.getBoundingClientRect();
      if (r.width < 12 || r.width > 72 || r.height < 12) continue;
      if (r.x < minX) continue;
      if (r.x > maxX) { maxX = r.x; best = el; }
    }
  }
  if (!best) return { state: 'missing', x: 0, y: 0 };
  const s = getComputedStyle(best);
  const op = parseFloat(s.opacity || '1');
  const bgLum = lum(s.backgroundColor);
  const cls = (best.className || '').toString();
  const hasAnim = (el) => {
    const s = getComputedStyle(el);
    if (s.animationName && s.animationName !== 'none') return true;
    for (const c of el.querySelectorAll('*')) {
      const cs = getComputedStyle(c);
      if (cs.animationName && cs.animationName !== 'none') return true;
      const cn = (c.className || '').toString();
      if (/loading|spinning|generat/i.test(cn)) return true;
    }
    return false;
  };
  const busy =
    best.getAttribute('aria-busy') === 'true' ||
    /loading|spinning|generat/i.test(cls) ||
    hasAnim(best);
  const r = best.getBoundingClientRect();
  const pos = { x: r.x + r.width / 2, y: r.y + r.height / 2 };
  // 灰：生成完毕、不可点
  if (op < 0.55 || bgLum > 0.68) return { state: 'done', ...pos };
  // 黑点：生成中（有动画/aria-busy，不是静态箭头 svg）
  if (busy && op >= 0.75) return { state: 'generating', ...pos };
  // 黑：可发送
  if (op >= 0.82 && bgLum < 0.52) return { state: 'ready', ...pos };
  if (op < 0.72) return { state: 'done', ...pos };
  return { state: 'unknown', ...pos };
}
"""

_SEND_STATE_LABEL = {
    "ready": "黑色（可发送）",
    "generating": "黑点（生成中）",
    "done": "灰色（已结束）",
    "missing": "未找到",
    "unknown": "未知",
}


def get_send_arrow_state(page) -> Dict[str, Any]:
    """发送箭头三态：ready=黑 / generating=黑点 / done=灰。"""
    return page.evaluate(_SEND_ARROW_JS)


def wait_for_send_ready(page, timeout_sec: int = 120) -> None:
    """仅当箭头为黑色（ready）时才算可发送；灰/黑点继续等，不提前退出。"""
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        info = get_send_arrow_state(page)
        st = info.get("state", "missing")
        if st == "ready":
            return
        page.wait_for_timeout(500)
    info = get_send_arrow_state(page)
    raise RuntimeError(
        f"等待发送箭头变黑超时，当前：{_SEND_STATE_LABEL.get(info.get('state','unknown'), info.get('state'))}"
    )


def _message_list_child_count(page) -> int:
    return int(
        page.evaluate(
            """
            () => {
              const list = document.querySelector('.message-list');
              return list ? list.children.length : 0;
            }
            """
        )
    )


def _find_send_button_locator(page):
    """发送钮在工具栏最右侧：黑底白箭头圆钮；K2.6 在其左侧。"""
    action = page.locator("#chat-box .chat-editor-action").first
    n = action.locator(":scope > *").count()
    if n == 0:
        return None
    last = action.locator(":scope > *").nth(n - 1)
    try:
        label = (last.inner_text(timeout=1500) or "").strip()
    except Exception:
        label = ""
    if any(k in label for k in ("K2", "快速", "Agent")):
        if n >= 2:
            return action.locator(":scope > *").nth(n - 2)
        return None
    return last


def _dispatch_click_send_in_page(page) -> bool:
    return bool(
        page.evaluate(
            """
            () => {
              const bar = document.querySelector('#chat-box .chat-editor-action');
              if (!bar) return false;
              const ar = bar.getBoundingClientRect();
              const minX = ar.x + ar.width * 0.78;
              let best = null, maxX = -1;
              for (const el of bar.querySelectorAll('button, [role="button"], div, span')) {
                const t = (el.innerText || el.textContent || '').trim();
                if (/Agent|K2|快速/i.test(t)) continue;
                const r = el.getBoundingClientRect();
                if (r.width < 22 || r.width > 52 || r.height < 22 || r.height > 52) continue;
                if (r.x < minX) continue;
                if (r.x > maxX) { maxX = r.x; best = el; }
              }
              if (!best) return false;
              best.click();
              return true;
            }
            """
        )
    )


def _click_send_button_dom(page) -> bool:
    """点击右下角黑色圆形发送钮。"""
    loc = _find_send_button_locator(page)
    if not loc:
        return False
    try:
        loc.scroll_into_view_if_needed(timeout=3000)
        loc.click(timeout=5000, force=True)
        return True
    except Exception:
        return False


def _send_took_effect(page, *, before_msgs: int) -> bool:
    """发送成功：新消息气泡或黑点生成中；箭头仍为黑色则视为未发送。"""
    if _message_list_child_count(page) > before_msgs:
        return True
    st = get_send_arrow_state(page).get("state")
    if st == "ready":
        return False
    if st == "generating":
        return True
    if _is_generating(page):
        return True
    return False


def click_send(page) -> None:
    """必须真实点到黑色圆钮发送；箭头仍为黑色则继续重试。"""
    wait_for_send_ready(page)
    before_msgs = _message_list_child_count(page)
    info = get_send_arrow_state(page)
    if info.get("state") != "ready":
        label = _SEND_STATE_LABEL.get(info.get("state", ""), info.get("state"))
        raise RuntimeError(f"发送箭头不是黑色，当前：{label}")

    for attempt in range(5):
        print(f"  点击发送箭头（第 {attempt + 1} 次）…", flush=True)
        _click_send_button_dom(page)
        page.wait_for_timeout(500)
        _dispatch_click_send_in_page(page)
        page.wait_for_timeout(1200)
        if _send_took_effect(page, before_msgs=before_msgs):
            print("  发送已生效（新消息或生成中）", flush=True)
            return
        loc = _find_send_button_locator(page)
        if loc:
            try:
                box = loc.bounding_box()
                if box:
                    page.mouse.click(
                        box["x"] + box["width"] / 2,
                        box["y"] + box["height"] / 2,
                    )
                    page.wait_for_timeout(1200)
                    if _send_took_effect(page, before_msgs=before_msgs):
                        print("  发送已生效（坐标点击）", flush=True)
                        return
            except Exception:
                pass
        info = get_send_arrow_state(page)
        if info.get("state") != "ready":
            continue
        editor = page.locator("#chat-box .chat-input-editor").first
        editor.click(timeout=3000)
        page.keyboard.press("Enter")
        page.wait_for_timeout(1200)
        if _send_took_effect(page, before_msgs=before_msgs):
            print("  发送已生效（Enter）", flush=True)
            return
        info = get_send_arrow_state(page)

    st = get_send_arrow_state(page).get("state", "missing")
    raise RuntimeError(
        f"发送未生效（箭头仍为 {_SEND_STATE_LABEL.get(st, st)}），"
        "请 --headed 看右下角黑圆钮是否被点击"
    )


def _is_generating(page) -> bool:
    st = get_send_arrow_state(page).get("state")
    if st == "generating":
        return True
    for sel in (
        '[aria-label*="停止"]',
        "text=停止",
        '[class*="stop-generat" i]',
    ):
        try:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible():
                return True
        except Exception:
            continue
    return False


def _is_send_done(page) -> bool:
    return get_send_arrow_state(page).get("state") == "done"


def _scroll_messages_to_bottom(page) -> None:
    page.evaluate(
        """
        () => {
          const el = document.querySelector('.message-list')
            || document.querySelector('.message-list-container');
          if (el) el.scrollTop = el.scrollHeight;
          window.scrollTo(0, document.body.scrollHeight);
        }
        """
    )


def get_latest_reply_text(page) -> str:
    """读 .message-list 全文（发送后输入区会重置，不能依赖箭头 DOM）。"""
    text = page.evaluate(
        """
        () => {
          const list = document.querySelector('.message-list')
            || document.querySelector('.message-list-container');
          if (list) return list.innerText || '';
          return '';
        }
        """
    )
    text = (text or "").strip()
    if len(text) > 20:
        return text
    for sel in (".message-list", ".message-list-container", "#app"):
        loc = page.locator(sel).first
        if loc.count():
            try:
                t = loc.inner_text(timeout=10_000).strip()
                if len(t) > len(text):
                    text = t
            except Exception:
                pass
    return text


def wait_for_agent_reply(page, timeout_sec: int) -> str:
    """见到回复末尾 COUNT=N 再结束；不因部分 JSON 或箭头变灰提前退出。"""
    page.wait_for_timeout(800)
    deadline = time.time() + timeout_sec
    last_log = 0.0
    saw_generating = False
    started = time.time()
    while time.time() < deadline:
        page.wait_for_timeout(1500)
        _scroll_messages_to_bottom(page)
        text = get_latest_reply_text(page)
        records, count, _ = parse_agent_response(text)
        arrow = get_send_arrow_state(page)
        ast = arrow.get("state", "missing")
        if ast == "generating":
            saw_generating = True
        if _is_generating(page):
            saw_generating = True

        if count is not None:
            n = len(records)
            print(
                f"  看到 COUNT={count}（已解析 {n} 条），结束等待",
                flush=True,
            )
            return text

        if (
            ast == "ready"
            and not saw_generating
            and time.time() - started > 45
        ):
            raise RuntimeError(
                "等待超时：未见 COUNT= 且发送可能未成功，请 --headed 重试"
            )

        if time.time() - last_log >= 20:
            hint = f"COUNT={count}" if count else "无COUNT"
            print(
                f"  等待中… 箭头={_SEND_STATE_LABEL.get(ast, ast)} "
                f"已解析={len(records)} 条 {hint}",
                flush=True,
            )
            last_log = time.time()

    _scroll_messages_to_bottom(page)
    text = get_latest_reply_text(page)
    _, count, _ = parse_agent_response(text)
    if count is not None:
        print(f"  超时前见到 COUNT={count}，结束等待", flush=True)
        return text
    return text


def submit_prompt_and_collect(
    page,
    prompt: str,
    *,
    response_timeout: int,
    debug_dir: Optional[Path],
    file_id: str,
) -> str:
    """填入提示词 → 点黑色发送箭头 → 等待 Kimi 输出。"""
    _ensure_agent_toggle_off(page)
    fill_prompt(page, prompt)
    page.wait_for_timeout(1000)
    _debug_shot(page, debug_dir, f"{file_id}_submit_01_filled.png")
    _ensure_agent_toggle_off(page)
    print("  等待发送箭头变黑（可发送）…", flush=True)
    click_send(page)
    _debug_shot(page, debug_dir, f"{file_id}_submit_02_sent.png")
    print("  等待 Kimi 生成回复（以 COUNT= 为准）…", flush=True)
    raw = wait_for_agent_reply(page, response_timeout)
    if not raw.strip():
        raise RuntimeError("未抓取到回复内容")
    _debug_shot(page, debug_dir, f"{file_id}_submit_03_reply.png")
    return raw


def save_reply_and_json(
    file_id: str,
    folder: Path,
    raw: str,
    *,
    debug_dir: Optional[Path],
) -> Tuple[bool, str, int]:
    """解析回复并写入 temp_jsons/notice__<ID>.json。"""
    if debug_dir:
        debug_dir.mkdir(parents=True, exist_ok=True)
        (debug_dir / f"{file_id}_reply.txt").write_text(raw, encoding="utf-8")

    records, count, meta = parse_agent_response(raw)
    records = normalize_records(
        records, file_id, read_source_url(folder), meta=meta
    )
    if count is not None and len(records) > count:
        seen: set[str] = set()
        deduped: List[Dict[str, str]] = []
        for row in records:
            key = (row.get("身份证号") or row.get("姓名") or "").strip()
            if key and key in seen:
                continue
            if key:
                seen.add(key)
            deduped.append(row)
        records = deduped[:count] if len(deduped) > count else deduped
    if not records:
        return False, "未能从回复中解析 JSON 数组", 0

    out_path = output_json_path(file_id)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    hint = f"保存 {len(records)} 条 -> {out_path.name}"
    if count is not None:
        hint += f"（COUNT={count}）"
    return True, hint, len(records)


def process_folder(
    page,
    folder: Path,
    *,
    response_timeout: int,
    debug_dir: Optional[Path],
    skip_upload: bool = False,
) -> Tuple[bool, str, int]:
    file_id = folder.name
    upload_list = collect_upload_files(folder)
    if not upload_list and not skip_upload:
        return False, "文件夹无可用附件", 0

    prompt = build_prompt(file_id, folder)
    if not skip_upload:
        print(f"  上传 {len(upload_list)} 个: {', '.join(p.name for p in upload_list)}", flush=True)
        open_new_chat(page)
        print(f"  页面: {page.url}", flush=True)
        upload_files(page, upload_list, debug_dir=debug_dir)
        page.wait_for_timeout(2000)
    else:
        wait_for_composer(page)
        print("  跳过上传（--skip-upload），仅提交并抓取回复", flush=True)

    raw = submit_prompt_and_collect(
        page,
        prompt,
        response_timeout=response_timeout,
        debug_dir=debug_dir,
        file_id=file_id,
    )
    return save_reply_and_json(file_id, folder, raw, debug_dir=debug_dir)


def cmd_inspect_ui(headless: bool) -> None:
    """打开 Kimi，暂停后导出页面控件清单（用于校准选择器）。"""
    from playwright.sync_api import sync_playwright

    out = PROJECT_ROOT / "kimi_debug" / "ui_candidates.json"
    print("将打开 Kimi；请在页面中手动进入「可上传附件的新对话」界面。", flush=True)
    print("准备好后回到本终端按 Enter，会扫描当前页面并保存控件列表。", flush=True)
    with sync_playwright() as p:
        ctx = launch_browser(p, headless=False)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(KIMI_HOME, wait_until="domcontentloaded", timeout=120_000)
        try:
            input("\n>>> 在首页新对话、可点 + 后，按 Enter 扫描页面…\n")
        except KeyboardInterrupt:
            print("已取消", flush=True)
            ctx.close()
            return
        dump_ui_candidates(page, out)
        try:
            page.screenshot(path=str(out.with_suffix(".png")), full_page=True)
            print(f"整页截图: {out.with_suffix('.png')}", flush=True)
        except Exception as err:
            print(f"截图失败: {err}", flush=True)
        ctx.close()


def cmd_login(headless: bool) -> None:
    from playwright.sync_api import sync_playwright

    print("将打开本机 Chrome，请在页面中登录 Kimi。", flush=True)
    print(f"会话目录: {PROFILE_DIR}", flush=True)
    with sync_playwright() as p:
        ctx = launch_browser(p, headless=False)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(KIMI_HOME, wait_until="domcontentloaded")
        print("登录完成后按 Enter 关闭浏览器并保存会话…", flush=True)
        try:
            input()
        except KeyboardInterrupt:
            print("\n已取消", flush=True)
        ctx.close()
    print("会话已保存，后续批处理无需重复登录（除非过期）。", flush=True)


def cmd_extract(
    folders: List[Path],
    *,
    headless: bool,
    skip_existing: bool,
    skip_upload: bool,
    delay_sec: float,
    response_timeout: int,
    debug: bool,
) -> None:
    from playwright.sync_api import sync_playwright

    if not PROMPT_PATH.is_file():
        print(f"缺少提示词: {PROMPT_PATH}", flush=True)
        sys.exit(1)

    if not folders:
        print("没有待处理文件夹", flush=True)
        return

    failed_ids = sync_failed_ids_registry()
    print_batch_plan(folders, skip_existing=skip_existing, failed_ids=failed_ids)
    if headless:
        print("Chrome 无头模式（加 --headed 可显示窗口）", flush=True)
    else:
        print("Chrome 有界面模式", flush=True)
    if debug:
        print(f"调试输出目录: {PROJECT_ROOT / 'kimi_debug'}", flush=True)

    ok = fail = skip = 0
    debug_dir = PROJECT_ROOT / "kimi_debug" if debug else None

    with sync_playwright() as p:
        ctx = launch_browser(p, headless=headless)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        for i, folder in enumerate(folders, 1):
            fid = folder.name
            out = output_json_path(fid)
            print(f"[{i}/{len(folders)}] {fid}", flush=True)

            if skip_existing and existing_ok(out):
                print(f"  跳过（已有 {out.name}）", flush=True)
                register_success(fid)
                skip += 1
                continue

            try:
                success, msg, n = process_folder(
                    page,
                    folder,
                    response_timeout=response_timeout,
                    debug_dir=debug_dir,
                    skip_upload=skip_upload,
                )
            except Exception as err:
                success, msg, n = False, str(err), 0
                if debug_dir:
                    debug_dir.mkdir(parents=True, exist_ok=True)
                    try:
                        page.screenshot(
                            path=str(debug_dir / f"{fid}_error.png"), full_page=True
                        )
                    except Exception:
                        pass

            if success:
                print(f"  OK: {msg} -> {out.name}", flush=True)
                register_success(fid)
                ok += 1
            else:
                print(f"  失败: {msg}", flush=True)
                register_failure(fid, msg)
                remove_invalid_output(fid)
                fail += 1

            if i < len(folders) and delay_sec > 0:
                time.sleep(delay_sec)

        ctx.close()

    pending = read_failed_ids()
    print(f"\n完成: 成功 {ok} | 失败 {fail} | 跳过 {skip}", flush=True)
    if fail or pending:
        print(f"失败明细: {FAIL_LOG}", flush=True)
        print(f"待重跑 ID: {FAILED_IDS_FILE}（共 {len(pending)} 个）", flush=True)
        if pending:
            preview = ", ".join(sorted(pending, key=lambda x: int(x))[:12])
            suffix = " …" if len(pending) > 12 else ""
            print(f"  {preview}{suffix}", flush=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Kimi 网页批量抽取公告名单（Playwright + 本机 Chrome）"
    )
    parser.add_argument(
        "--inspect-ui",
        action="store_true",
        help="打开 Kimi，手动进入新对话后导出控件清单到 kimi_debug/ui_candidates.json",
    )
    parser.add_argument("--login", action="store_true", help="打开 Chrome 登录 Kimi")
    parser.add_argument(
        "--headed",
        action="store_true",
        help="显示 Chrome 窗口（默认无头后台运行）",
    )
    parser.add_argument("--ids", type=str, default=None, help="逗号分隔公告 ID")
    parser.add_argument(
        "--from-id",
        type=int,
        default=None,
        help="批量起始公告 ID（含，需 data 下存在对应文件夹）",
    )
    parser.add_argument(
        "--to-id",
        type=int,
        default=None,
        help="批量结束公告 ID（含）",
    )
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        default=True,
        help="跳过已有有效 notice__<ID>.json（批处理默认开启）",
    )
    parser.add_argument(
        "--no-skip-existing",
        action="store_false",
        dest="skip_existing",
        help="不跳过，全部重新跑",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="同 --no-skip-existing，覆盖已有 JSON",
    )
    parser.add_argument(
        "--retry-failed",
        action="store_true",
        help="仅重跑 kimi_extract_failed_ids.txt 中的 ID（仍受 --from-id/--to-id 限制）",
    )
    parser.add_argument("--delay", type=float, default=3.0, help="每个公告间隔秒数")
    parser.add_argument(
        "--timeout",
        type=int,
        default=300,
        help="等待 Kimi 回复的最长时间（秒）",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="写入 kimi_debug/（回复 txt、各步截图；默认不生成）",
    )
    parser.add_argument(
        "--skip-upload",
        action="store_true",
        help="跳过上传（当前对话已手动放好附件时用，仅测提交+保存结果）",
    )
    parser.add_argument("--merge-only", action="store_true", help="仅合并 temp_jsons 到 CSV")
    args = parser.parse_args()

    if args.merge_only:
        merge_jsons_to_csv()
        return

    if args.inspect_ui:
        if not PROFILE_DIR.exists():
            print("请先: python kimi_web_extract.py --login", flush=True)
            return
        cmd_inspect_ui(headless=False)
        return

    if args.login:
        cmd_login(headless=False)
        return

    target_ids = None
    if args.ids:
        target_ids = {x.strip() for x in args.ids.split(",") if x.strip()}

    if args.from_id is not None and args.to_id is not None and args.from_id > args.to_id:
        print("--from-id 不能大于 --to-id", flush=True)
        sys.exit(1)

    folders = plan_batch_folders(
        from_id=args.from_id,
        to_id=args.to_id,
        ids=target_ids,
        limit=args.limit,
        retry_failed=args.retry_failed,
    )
    if not folders and not args.retry_failed:
        print(f"未找到目录: {NOTICE_DATA_DIR}", flush=True)
        return

    headless = not args.headed
    if not PROFILE_DIR.exists():
        print("首次运行请先: python kimi_web_extract.py --login", flush=True)
        return

    skip_existing = args.skip_existing and not args.force
    cmd_extract(
        folders,
        headless=headless,
        skip_existing=skip_existing,
        skip_upload=args.skip_upload,
        delay_sec=args.delay,
        response_timeout=args.timeout,
        debug=args.debug,
    )


if __name__ == "__main__":
    main()
