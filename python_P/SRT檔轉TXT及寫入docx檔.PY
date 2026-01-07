import re
import sys
from pathlib import Path
from datetime import datetime

from docx import Document


# -------------------------
# 自動嘗試編碼讀取
# -------------------------
def read_text_auto(path: Path) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp950", "big5", "utf-16"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="ignore")


# -------------------------
# 去掉 SRT 的序號/時間軸，只留字幕
# -------------------------
TIME_RE = re.compile(r"^\d{2}:\d{2}:\d{2}[,\.]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[,\.]\d{3}")

def extract_subtitle_lines_from_srt(srt_path: Path) -> list[str]:
    text = read_text_auto(srt_path)
    lines = text.splitlines()

    out_lines: list[str] = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.isdigit():
            continue
        if TIME_RE.match(line):
            continue
        out_lines.append(line)
    return out_lines


# -------------------------
# SRT -> TXT（同名 TXT 存在可跳過）
# -------------------------
def write_txt_if_needed(srt_path: Path, lines: list[str], *, skip_if_exists: bool = True) -> str:
    txt_path = srt_path.with_suffix(".txt")
    if skip_if_exists and txt_path.exists():
        return "skipped"
    txt_path.write_text("\n".join(lines), encoding="utf-8")
    return "converted"


# -------------------------
# 找到「最可能」的同名 DOCX
# - 先精準找：同 stem 的 .docx
# - 找不到再找：stem*.docx（檔名可能多了後綴）
# - 多個候選：挑最新修改時間
# -------------------------
def find_matching_docx(srt_path: Path) -> Path | None:
    exact = srt_path.with_suffix(".docx")
    if exact.exists():
        return exact

    # 允許同目錄下有後綴版本：xxx*.docx
    candidates = list(srt_path.parent.glob(srt_path.stem + "*.docx"))
    if not candidates:
        return None

    # 取最新修改的那份
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


# -------------------------
# 追加寫入 DOCX（若存在）
# - 若存檔被鎖住，會改存成 *_APPENDED_yyyymmdd_hhmmss.docx
# -------------------------
def append_to_docx_if_exists(srt_path: Path, lines: list[str]) -> tuple[str, str]:
    docx_path = find_matching_docx(srt_path)
    if not docx_path:
        return "no_docx", ""

    try:
        doc = Document(str(docx_path))
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        doc.add_paragraph("")  # 空一行
        doc.add_paragraph(f"【SRT轉入】{srt_path.name}  ({ts})")

        for t in lines:
            doc.add_paragraph(t)

        try:
            doc.save(str(docx_path))
            return "appended", str(docx_path)
        except PermissionError:
            # Word 開著常見：另存新檔避免你以為沒成功
            ts2 = datetime.now().strftime("%Y%m%d_%H%M%S")
            alt = docx_path.with_name(docx_path.stem + f"_APPENDED_{ts2}" + docx_path.suffix)
            doc.save(str(alt))
            return "saved_as_copy", str(alt)

    except Exception as e:
        return f"failed: {e}", str(docx_path)


# -------------------------
# 取得使用者輸入（支援參數 or 互動）
# -------------------------
def get_input_path() -> str:
    if len(sys.argv) >= 2:
        return " ".join(sys.argv[1:]).strip()
    print("請輸入 SRT 檔案或資料夾路徑（可直接拖曳進來）：")
    return input(">> ").strip()

def clean_path(p: str) -> str:
    p = (p or "").strip()
    if (p.startswith('"') and p.endswith('"')) or (p.startswith("'") and p.endswith("'")):
        p = p[1:-1].strip()
    return p


# -------------------------
# 單檔處理
# -------------------------
def process_one_srt(srt: Path) -> tuple[bool, str, str, str, int]:
    """
    回傳: (ok, txt_status, docx_status, docx_written_path, line_count)
    """
    try:
        lines = extract_subtitle_lines_from_srt(srt)
        txt_status = write_txt_if_needed(srt, lines, skip_if_exists=True)
        docx_status, docx_written_path = append_to_docx_if_exists(srt, lines)
        return True, txt_status, docx_status, docx_written_path, len(lines)
    except Exception as e:
        return False, "failed", f"failed: {e}", "", 0


# -------------------------
# main
# -------------------------
def main():
    raw = clean_path(get_input_path())

    print("--------------------------------------------------")
    print("收到的路徑：", raw)
    print("--------------------------------------------------")

    if not raw:
        print("❌ 未輸入任何路徑")
        return

    path = Path(raw)
    if not path.exists():
        print("❌ 路徑不存在：", path)
        return

    # 單檔
    if path.is_file():
        if path.suffix.lower() != ".srt":
            print("❌ 這不是 .srt 檔：", path)
            return

        ok, txt_status, docx_status, docx_written, n = process_one_srt(path)
        if not ok:
            print("❌ 轉換失敗：", path)
            print("TXT:", txt_status, "| DOCX:", docx_status)
            return

        # TXT 狀態
        if txt_status == "skipped":
            print(f"⏭️  TXT跳過（已存在）：{path.with_suffix('.txt')}")
        else:
            print(f"✅ TXT完成：{path.with_suffix('.txt')}（行數 {n}）")

        # DOCX 狀態
        if docx_status == "no_docx":
            print("ℹ️ 同目錄找不到可寫入的 DOCX（需同名或同名開頭 *.docx）。")
        elif docx_status == "appended":
            print(f"✅ DOCX已追加：{docx_written}")
        elif docx_status == "saved_as_copy":
            print("⚠️ DOCX 原檔可能被 Word 開著鎖住，已改存成新檔：")
            print(f"✅ DOCX已追加(另存)：{docx_written}")
        else:
            print(f"❌ DOCX追加失敗：{docx_status}")
            if docx_written:
                print("   目標DOCX：", docx_written)

        return

    # 資料夾（含子目錄）
    if path.is_dir():
        srt_files = list(path.rglob("*.srt"))
        if not srt_files:
            print("⚠️ 資料夾內（含子目錄）找不到任何 .srt")
            return

        converted = 0
        skipped = 0
        failed = 0
        docx_appended = 0
        docx_saved_copy = 0
        docx_missing = 0
        docx_failed = 0

        print(f"🔍 找到 {len(srt_files)} 個 SRT（TXT 同名存在會跳過；DOCX 同名/同名開頭存在會追加）")
        print("開始處理...")

        for srt in srt_files:
            ok, txt_status, docx_status, docx_written, _ = process_one_srt(srt)

            if not ok:
                failed += 1
                print("❌ [失敗]", srt)
                continue

            if txt_status == "skipped":
                skipped += 1
                print("⏭️  [TXT跳過]", srt)
            else:
                converted += 1
                print("✅ [TXT完成]", srt)

            if docx_status == "appended":
                docx_appended += 1
                print("   ↳ ✅ [DOCX追加]", docx_written)
            elif docx_status == "saved_as_copy":
                docx_saved_copy += 1
                print("   ↳ ⚠️ [DOCX被鎖 另存追加]", docx_written)
            elif docx_status == "no_docx":
                docx_missing += 1
            else:
                docx_failed += 1
                print("   ↳ ❌ [DOCX失敗]", docx_status, "|", docx_written or "")

        print("--------------------------------------------------")
        print(f"🎉 TXT：成功 {converted} / 跳過 {skipped} / 失敗 {failed} / 總數 {len(srt_files)}")
        print(f"📄 DOCX：直接追加 {docx_appended} / 被鎖另存 {docx_saved_copy} / 無DOCX {docx_missing} / 失敗 {docx_failed}")
        return

    print("❌ 不支援的路徑型態：", path)


if __name__ == "__main__":
    main()
    input("\n按 Enter 結束...")
