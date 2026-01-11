import os
import re
import csv
import sys
import traceback
from datetime import datetime
from urllib.parse import urlparse
from email.utils import parsedate_to_datetime

from docx import Document
from docx.shared import Pt

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

CSV_PATH = r"F:\F\AI\web\web.csv"
OUT_DIR  = r"F:\F\AI\web"
NAV_TIMEOUT_MS = 30000


def sanitize_filename(name: str) -> str:
    name = (name or "").strip()
    if not name:
        return ""
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120].strip()


def derive_name_from_url(url: str) -> str:
    try:
        u = urlparse(url)
        host = u.netloc or "site"
        path_last = (u.path or "").rstrip("/").split("/")[-1] or "index"
        q = (u.query or "").strip()
        if q:
            q = sanitize_filename(q)[:40]
            return sanitize_filename(f"{host}_{path_last}_{q}")
        return sanitize_filename(f"{host}_{path_last}")
    except Exception:
        return sanitize_filename(url)[:120] or "site_index"


def read_urls_from_csv(csv_path: str):
    rows = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for r in reader:
            if not r:
                continue
            url = (r[0] if len(r) > 0 else "").strip()
            if not url:
                break
            name = (r[1] if len(r) > 1 else "").strip()
            rows.append((url, name))
    return rows


def parse_date_loose(s: str):
    if not s:
        return None
    s = s.strip()

    m = re.search(r"(20\d{2})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    m2 = re.search(r"\b(20\d{2})(\d{2})(\d{2})\b", s)
    if m2:
        try:
            return datetime(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except ValueError:
            return None

    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None


def extract_date_from_url(url: str):
    return parse_date_loose(url)


def extract_publish_date(page, response, url: str):
    meta_selectors = [
        'meta[property="article:published_time"]',
        'meta[property="og:published_time"]',
        'meta[name="pubdate"]',
        'meta[name="publish-date"]',
        'meta[name="date"]',
        'meta[itemprop="datePublished"]',
        'meta[name="datePublished"]',
    ]

    # 1) meta（先 count() 再取 attribute，避免 locator 不存在直接炸）
    for sel in meta_selectors:
        try:
            loc = page.locator(sel)
            if loc.count() > 0:
                val = loc.first.get_attribute("content")
                d = parse_date_loose(val or "")
                if d:
                    return d
        except Exception:
            pass

    # 2) <time datetime="">
    try:
        loc = page.locator("time[datetime]")
        if loc.count() > 0:
            val = loc.first.get_attribute("datetime")
            d = parse_date_loose(val or "")
            if d:
                return d
    except Exception:
        pass

    # 3) HTTP Last-Modified
    try:
        if response is not None:
            lm = response.headers.get("last-modified")
            if lm:
                try:
                    return parsedate_to_datetime(lm)
                except Exception:
                    d = parse_date_loose(lm)
                    if d:
                        return d
    except Exception:
        pass

    # 4) URL
    return extract_date_from_url(url)


def fetch_page_all(page, url: str):
    resp = page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)

    try:
        title = (page.title() or "").strip()
    except Exception:
        title = ""

    try:
        text = page.evaluate("() => document.body ? document.body.innerText : ''") or ""
        text = text.strip()
    except Exception:
        text = ""

    pub_date = extract_publish_date(page, resp, url)
    return title, text, pub_date


def write_docx(title: str, url: str, text: str, out_path: str):
    doc = Document()

    if title:
        p = doc.add_paragraph(title)
        if p.runs:
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(16)

    p2 = doc.add_paragraph(url)
    if p2.runs:
        p2.runs[0].font.size = Pt(10)

    doc.add_paragraph("")

    for line in (text or "").splitlines():
        line = line.strip()
        if line:
            doc.add_paragraph(line)

    doc.save(out_path)


def main():
    print("[START] web_csv_to_docx")
    print(f"[INFO] Python: {sys.executable}")
    print(f"[INFO] CSV_PATH: {CSV_PATH}")
    print(f"[INFO] OUT_DIR : {OUT_DIR}")

    if not os.path.isfile(CSV_PATH):
        print(f"[ERROR] 找不到 CSV：{CSV_PATH}")
        print("        你可以先確認檔案是否真的存在，或路徑有沒有打錯。")
        return

    os.makedirs(OUT_DIR, exist_ok=True)

    items = read_urls_from_csv(CSV_PATH)
    print(f"[INFO] CSV 讀到 {len(items)} 筆網址")

    if not items:
        print("[WARN] CSV 沒有讀到任何 URL（A欄可能是空的，或第一列就空）")
        return

    # ✅ 這行一定要有冒號
    with sync_playwright() as p:
        print("[INFO] Playwright 啟動 OK")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        ok = 0
        fail = 0

        for idx, (url, name_in_csv) in enumerate(items, start=1):
            url = (url or "").strip()
            if not url:
                continue

            print(f"\n[DO] ({idx}) {url}")

            try:
                title, text, pub_date = fetch_page_all(page, url)
            except PWTimeoutError:
                print(f"[ERR] ({idx}) 逾時：{url}")
                fail += 1
                continue
            except Exception as e:
                print(f"[ERR] ({idx}) 失敗：{url}\n      {e}")
                fail += 1
                continue

            if not pub_date:
                print(f"[WARN] ({idx}) 抓不到日期，使用 00000000")
                date_str = "00000000"
            else:
                try:
                    date_str = pub_date.strftime("%Y%m%d")
                except Exception:
                    date_str = "00000000"

            name = sanitize_filename(name_in_csv)
            if not name:
                name = sanitize_filename(title)
            if not name:
                name = derive_name_from_url(url)
            if not name:
                name = "unnamed"

            out_path = os.path.join(OUT_DIR, f"{date_str}_{name}.docx")

            if os.path.exists(out_path):
                print(f"[SKIP] 已存在：{os.path.basename(out_path)}")
                continue

            try:
                write_docx(title=title, url=url, text=text, out_path=out_path)
                print(f"[OK]  輸出：{os.path.basename(out_path)}")
                ok += 1
            except Exception as e:
                print(f"[ERR] 寫入 DOCX 失敗：{e}")
                fail += 1

        context.close()
        browser.close()

    print(f"\n[DONE] OK={ok}, FAIL={fail}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("[FATAL] 程式發生未捕捉例外：")
        traceback.print_exc()
