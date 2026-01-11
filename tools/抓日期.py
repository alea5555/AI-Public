import re
import requests
from bs4 import BeautifulSoup

MONTH_MAP = {
    "jan": "01","feb": "02","mar": "03","apr": "04","may": "05","jun": "06",
    "jul": "07","aug": "08","sep": "09","oct": "10","nov": "11","dec": "12"
}

def get_pixnet_date8(url: str) -> str:
    r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or "utf-8"
    soup = BeautifulSoup(r.text, "lxml")

    y = soup.select_one("li.publish span.year")
    m = soup.select_one("li.publish span.month")
    d = soup.select_one("li.publish span.date")

    if not (y and m and d):
        return ""  # 找不到就回空字串

    year = y.get_text(strip=True)
    mon_raw = m.get_text(strip=True).lower()
    day = d.get_text(strip=True)

    mon = MONTH_MAP.get(mon_raw, "")
    day = re.sub(r"\D", "", day).zfill(2)

    if not (year.isdigit() and mon and day.isdigit()):
        return ""

    return f"{year}{mon}{day}"

# 測試
print(get_pixnet_date8("https://ey90223.pixnet.net/blog/posts/10353867503"))
