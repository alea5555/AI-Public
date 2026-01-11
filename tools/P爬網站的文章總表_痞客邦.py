import os
import re
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import pandas as pd


OUT_DIR = r"F:\F\AI"
SLEEP_SEC = 0.8  # 爬取間隔


def safe_filename(name: str, max_len: int = 120) -> str:
    # Windows 不可用字元: <>:"/\|?*
    name = re.sub(r'[<>:"/\\|?*]', "_", name).strip()
    name = re.sub(r"\s+", " ", name)
    if len(name) > max_len:
        name = name[:max_len].rstrip()
    return name or "output"


def fetch_soup(session: requests.Session, url: str) -> BeautifulSoup:
    r = session.get(url, timeout=30)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")


def extract_posts(soup: BeautifulSoup):
    rows = []

    articles = soup.select("article")
    if not articles:
        articles = soup.select("h2")

    for a in articles:
        title_a = a.select_one("h2 a") if a.name != "h2" else a.select_one("a")
        if not title_a or not title_a.get("href"):
            continue

        title = title_a.get_text(strip=True)
        link = title_a["href"].strip()

        dt = ""
        time_el = a.select_one("time[datetime]") if a.name != "h2" else None
        if time_el and time_el.get("datetime"):
            dt = time_el["datetime"][:10]
        else:
            time_any = a.select_one("time") if a.name != "h2" else None
            if time_any:
                dt = time_any.get_text(strip=True)

        rows.append({
            "日期": dt,
            "名稱": title,
            "網址": link
        })

    uniq = {}
    for r in rows:
        uniq[r["網址"]] = r
    return list(uniq.values())


def find_next_page(soup: BeautifulSoup, current_url: str):
    a = soup.select_one('a[rel="next"]')
    if a and a.get("href"):
        return urljoin(current_url, a["href"])

    for cand in soup.select("a"):
        if "next" in cand.get_text(strip=True).lower() and cand.get("href"):
            return urljoin(current_url, cand["href"])
    return None


def main():
    start_url = input("請輸入要爬取的 EZQuant Blog 網址：\n").strip()
    if not start_url:
        print("❌ 未輸入網址，結束程式")
        return

    os.makedirs(OUT_DIR, exist_ok=True)

    with requests.Session() as s:
        s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        })

        first_soup = fetch_soup(s, start_url)

        page_title = first_soup.title.get_text(strip=True) if first_soup.title else "blog"
        file_name = safe_filename(page_title) + ".xlsx"
        out_path = os.path.join(OUT_DIR, file_name)

        all_rows = []
        url = start_url
        seen = set()

        while url and url not in seen:
            seen.add(url)
            soup = first_soup if url == start_url else fetch_soup(s, url)

            all_rows.extend(extract_posts(soup))
            url = find_next_page(soup, url)

            time.sleep(SLEEP_SEC)

        uniq = {}
        for r in all_rows:
            uniq[r["網址"]] = r

        df = pd.DataFrame(
            uniq.values(),
            columns=["日期", "名稱", "網址"]
        )

        df.to_excel(out_path, index=False)
        print(f"✅ 完成，共 {len(df)} 筆")
        print(f"📄 輸出位置：{out_path}")


if __name__ == "__main__":
    main()
