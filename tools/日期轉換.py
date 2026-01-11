import re

MONTH = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12"
}

def convert_date(s: str) -> str:
    """
    將輸入字串（例如：May 03 Sun 202022:17）
    轉成 YYYY-MM-DD
    """
    s = s.strip().lower()

    m = re.search(
        r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+(\d{1,2}).*?(20\d{2})",
        s
    )
    if not m:
        return "❌ 無法辨識日期格式"

    mon = MONTH[m.group(1)]
    day = m.group(2).zfill(2)
    year = m.group(3)

    return f"{year}-{mon}-{day}"


# ===== 主程式 =====
if __name__ == "__main__":
    while True:
        s = input("請輸入日期字串（Enter 結束）：\n").strip()
        if not s:
            break

        result = convert_date(s)
        print("轉換結果：", result)
