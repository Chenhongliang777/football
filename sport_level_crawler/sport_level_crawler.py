"""
国家体育总局运动员技术等级查询 
基于 requests + BeautifulSoup

依赖安装：
    pip install requests beautifulsoup4 lxml

使用方式：
    python sport_level_crawler_fast.py
"""

import json
import os
import sys
import time
from pathlib import Path

import requests

# ==================== 配置区域（已根据抓包配置） ====================
INPUT_FILE = "input.json"
OUTPUT_DIR = "data"

# 查询系统实际地址
QUERY_PAGE_URL = "https://zwfw.sport.gov.cn/level.html"

# 实际查询接口（抓包确认）
QUERY_API_URL = "https://zwfw.sport.gov.cn/level.do?m=getLevelList"

# 请求超时（秒）
TIMEOUT = 10

# 反爬很弱，采用标准浏览器 Headers 即可
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://zwfw.sport.gov.cn",
    "Referer": "https://zwfw.sport.gov.cn/level.html",
    "Connection": "keep-alive",
}

# 抓包确认的参数名
PARAM_CERT_NO = "certificate_num"   # 证书编号参数名
PARAM_NAME = "name"                 # 姓名参数名

# 结果解析配置（响应可能是 JSON 或 HTML）
# ==================================================


def ensure_dir(path: str) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def read_input(filepath: str) -> list[dict]:
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        return [data]
    elif isinstance(data, list):
        return data
    else:
        raise ValueError("input.json 应为对象或对象数组")


def safe_filename(name: str) -> str:
    import re
    return re.sub(r'[\\/:*?"<>|]', "_", name)

def query_one(session: requests.Session, cert_no: str, name: str, idx: int) -> dict:
    result = {
        "cert_no": cert_no,
        "name": name,
        "query_time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "success": False,
        "data": {},
        "error": None,
    }
    payload = {PARAM_CERT_NO: cert_no, PARAM_NAME: name}

    print(f"\n[{idx}] 查询: {name} ({cert_no})")
    start = time.perf_counter()

    try:
        resp = session.post(QUERY_API_URL, data=payload, timeout=TIMEOUT)
        resp.raise_for_status()
        resp.encoding = 'utf-8'
        elapsed = time.perf_counter() - start
        print(f"  请求耗时: {elapsed:.3f}s")

        # 直接解析 JSON，已知接口稳定返回 JSON
        result["success"] = True
        result["data"] = resp.json()
        print("  结果: JSON 已提取")

    except requests.exceptions.RequestException as e:
        result["error"] = f"网络请求失败: {e}"
        print(f"  [失败] {result['error']}")
    except Exception as e:
        result["error"] = f"解析异常: {e}"
        print(f"  [失败] {result['error']}")

    return result


def main():
    ensure_dir(OUTPUT_DIR)
    if not Path(INPUT_FILE).exists():
        sample = [{"证书编号": "CFA201900050", "姓名": "胡睿宝"}]
        with open(INPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(sample, f, ensure_ascii=False, indent=2)
        print(f"已创建示例输入文件: {INPUT_FILE}，请按需修改后继续。")
        sys.exit(0)

    records = read_input(INPUT_FILE)
    print(f"共读取 {len(records)} 条查询记录")

    # Session 复用 TCP 连接
    session = requests.Session()
    session.headers.update(HEADERS)

    # 访问一次查询页面获取 Cookie（PHP 系统通常需要 session cookie）
    try:
        print(f"初始化会话: {QUERY_PAGE_URL}")
        session.get(QUERY_PAGE_URL, timeout=TIMEOUT)
        print("会话初始化完成\n")
    except Exception as e:
        print(f"[警告] 初始化页面访问失败: {e}，继续尝试直接查询...")

    all_results = []
    total_start = time.perf_counter()

    for i, rec in enumerate(records, start=1):
        cert_no = rec.get("证书编号") or rec.get("cert_no") or rec.get("编号") or ""
        name = rec.get("姓名") or rec.get("name") or rec.get("运动员姓名") or ""
        if not cert_no or not name:
            print(f"\n[{i}] 跳过无效记录: {rec}")
            continue

        res = query_one(session, cert_no, name, i)
        all_results.append(res)

        # ========== 字段清洗 ==========
        if res["success"] and isinstance(res.get("data"), dict):
            # 挖到最内层 data
            raw = res["data"]
            for _ in range(3):  # 最多剥 3 层 data 嵌套
                if isinstance(raw, dict) and "data" in raw:
                    raw = raw["data"]
                else:
                    break
            keep = {k: raw[k] for k in [
                "certificateNo","athleteRealName", "rankTitle", "subItemName", "smallItemName",
                "grantUnitName", "grantTime", "eventGrade", "eventName",
                "eventPlace", "eventTime","orderBookUrl","scoreBookUrl"
            ] if k in raw}
            
            res["data"] = keep
        # =====================================================

        # 保存单条结果
        single_path = Path(OUTPUT_DIR) / f"result_{safe_filename(name)}_{cert_no}.json"
        with open(single_path, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)

    # 保存汇总
    summary_path = Path(OUTPUT_DIR) / f"summary_{time.strftime('%Y%m%d_%H%M%S')}.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)

    total_elapsed = time.perf_counter() - total_start
    avg_speed = len(records) / total_elapsed if total_elapsed > 0 else 0
    print(f"\n 全部完成！汇总结果: {summary_path}")
    print(f"总耗时: {total_elapsed:.2f}s | 平均速度: {avg_speed:.2f} 条/秒")


if __name__ == "__main__":
    main()
