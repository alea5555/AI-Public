import subprocess
from pathlib import Path
import re
import sys
import tempfile

# =========================
# 固定目錄
# =========================
BASE_DIR = Path(r"F:\F\AI\downloads")

# =========================
# SRT 解析
# =========================
SRT_TIME_RANGE = re.compile(
    r"(?P<s>\d{2}:\d{2}:\d{2},\d{3})\s*-->\s*(?P<e>\d{2}:\d{2}:\d{2},\d{3})"
)
TIME_PATTERN = re.compile(r"(\d{2}):(\d{2}):(\d{2}),(\d{3})")

def time_to_ms(h, m, s, ms):
    return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(ms)

def srt_time_str_to_ms(t: str) -> int:
    m = TIME_PATTERN.match(t.strip())
    if not m:
        raise ValueError(f"SRT 時間格式錯誤：{t}")
    return time_to_ms(*m.groups())

def ms_to_srt_time(ms: int) -> str:
    if ms < 0:
        ms = 0
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms %= 1000
    return f"{h:02}:{m:02}:{s:02},{ms:03}"

def parse_time_input(text: str) -> float:
    """
    支援：
    - mm:ss（例如 1:48）
    - 秒數（例如 108 或 108.5）
    - 空白（=0）
    """
    text = text.strip()
    if not text:
        return 0.0
    if ":" in text:
        parts = text.split(":")
        if len(parts) != 2:
            raise ValueError("時間格式錯誤，請用 mm:ss")
        mm, ss = parts
        return int(mm) * 60 + float(ss)
    return float(text)

def shift_srt_all(input_srt: Path, output_srt: Path, shift_seconds: float):
    shift_ms = int(shift_seconds * 1000)
    text = input_srt.read_text(encoding="utf-8-sig")

    def repl(m):
        ms = time_to_ms(*m.groups())
        return ms_to_srt_time(ms - shift_ms)

    output_srt.write_text(TIME_PATTERN.sub(repl, text), encoding="utf-8")

def delete_srt_middle(input_srt: Path, output_srt: Path, a_sec: float, b_sec: float):
    """
    刪掉中間 [A, B] 的字幕區段：
    - 完全落在 [A,B] 的字幕：刪除
    - A 之前：保留不動
    - B 之後：整體往前移 (B-A)
    - 跨界字幕：把落在刪除區的部分裁掉（更合理）
    """
    a_ms = int(a_sec * 1000)
    b_ms = int(b_sec * 1000)
    if b_ms <= a_ms:
        raise ValueError("B 必須大於 A")

    shift_ms = b_ms - a_ms

    raw = input_srt.read_text(encoding="utf-8-sig")
    blocks = re.split(r"\r?\n\r?\n", raw.strip(), flags=re.M)

    kept_blocks = []

    for blk in blocks:
        lines = blk.splitlines()
        if len(lines) < 2:
            continue

        # 找時間行
        tl = None
        m = None
        for i, line in enumerate(lines):
            mm = SRT_TIME_RANGE.search(line)
            if mm:
                tl = i
                m = mm
                break
        if tl is None:
            continue

        s_ms = srt_time_str_to_ms(m.group("s"))
        e_ms = srt_time_str_to_ms(m.group("e"))

        # 情況 1：完全在 A 之前
        if e_ms <= a_ms:
            new_s, new_e = s_ms, e_ms

        # 情況 2：完全在 B 之後 -> 往前 shift
        elif s_ms >= b_ms:
            new_s, new_e = s_ms - shift_ms, e_ms - shift_ms

        # 情況 3：完全落在 [A,B] -> 刪掉
        elif s_ms >= a_ms and e_ms <= b_ms:
            continue

        # 情況 4：跨越 A（s < A < e <= B）-> 裁到 A
        elif s_ms < a_ms < e_ms <= b_ms:
            new_s, new_e = s_ms, a_ms

        # 情況 5：跨越 B（A <= s < B < e）-> 裁到 B，並 shift
        elif a_ms <= s_ms < b_ms < e_ms:
            new_s, new_e = b_ms - shift_ms, e_ms - shift_ms  # new_s = A
            # 說明：原本從 s 到 e，其中 s 在刪除區，等於從 B 之後才開始顯示，顯示時間對應到 A
        # 情況 6：兩邊都跨（s < A 且 e > B）-> 變成兩段？這裡選擇保留 A 前那段
        elif s_ms < a_ms and e_ms > b_ms:
            new_s, new_e = s_ms, a_ms
        else:
            continue

        if new_e <= new_s:
            continue

        lines[tl] = f"{ms_to_srt_time(new_s)} --> {ms_to_srt_time(new_e)}"
        kept_blocks.append(lines)

    # 重新編號輸出
    out_lines = []
    idx = 1
    for lines in kept_blocks:
        tl = None
        for i, line in enumerate(lines):
            if SRT_TIME_RANGE.search(line):
                tl = i
                break
        if tl is None:
            continue
        text_lines = lines[tl + 1:]

        out_lines.append(str(idx))
        out_lines.append(lines[tl])
        out_lines.extend(text_lines)
        out_lines.append("")
        idx += 1

    output_srt.write_text("\n".join(out_lines).rstrip() + "\n", encoding="utf-8")


# =========================
# 影片工具
# =========================
def get_duration(video: Path) -> float:
    cmd = [
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        str(video)
    ]
    return float(subprocess.check_output(cmd).decode().strip())

def cut_video_front_back(input_video: Path, output_video: Path, cut_front: float, cut_back: float):
    duration = get_duration(input_video)
    keep_start = cut_front
    keep_len = duration - cut_front - cut_back
    if keep_len <= 0:
        raise ValueError("修剪後影片長度 <= 0，請檢查輸入時間")

    cmd = [
        "ffmpeg", "-y",
        "-ss", str(keep_start),
        "-i", str(input_video),
        "-t", str(keep_len),
        "-c", "copy",
        "-movflags", "+faststart",
        str(output_video)
    ]
    subprocess.run(cmd, check=True)

def delete_video_middle_concat(input_video: Path, output_video: Path, a_sec: float, b_sec: float):
    """
    刪掉中間段 [A,B]：輸出 = [0,A] + [B,END]（無重編碼）
    """
    duration = get_duration(input_video)
    if a_sec < 0: a_sec = 0
    if b_sec <= a_sec:
        raise ValueError("B 必須大於 A")
    if a_sec >= duration:
        raise ValueError("A 超過影片長度")
    if b_sec > duration:
        b_sec = duration

    part1_len = a_sec
    part2_start = b_sec
    part2_len = duration - b_sec

    if part1_len <= 0 and part2_len <= 0:
        raise ValueError("刪除後影片長度 <= 0")

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        p1 = td / "part1.mp4"
        p2 = td / "part2.mp4"
        lst = td / "list.txt"

        # part1: 0 -> A
        if part1_len > 0:
            subprocess.run([
                "ffmpeg", "-y",
                "-i", str(input_video),
                "-t", str(part1_len),
                "-c", "copy",
                str(p1)
            ], check=True)

        # part2: B -> end
        if part2_len > 0:
            subprocess.run([
                "ffmpeg", "-y",
                "-ss", str(part2_start),
                "-i", str(input_video),
                "-c", "copy",
                str(p2)
            ], check=True)

        lines = []
        if part1_len > 0:
            lines.append(f"file '{p1.as_posix()}'")
        if part2_len > 0:
            lines.append(f"file '{p2.as_posix()}'")
        lst.write_text("\n".join(lines) + "\n", encoding="utf-8")

        # concat
        subprocess.run([
            "ffmpeg", "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", str(lst),
            "-c", "copy",
            "-movflags", "+faststart",
            str(output_video)
        ], check=True)

# =========================
# 主流程
# =========================
def main():
    videos = list(BASE_DIR.glob("*.mp4"))
    if not videos:
        print("❌ 目錄中找不到 MP4")
        sys.exit(1)

    video = videos[0]
    srt = video.with_suffix(".srt")

    print("📂 目錄：", BASE_DIR)
    print("🎬 影片：", video.name)
    print("📝 字幕：", (srt.name if srt.exists() else "無（將略過字幕處理）"))

    print("\n模式選擇：")
    print("  1) 前後修剪（各自輸入，預設 0）")
    print("  2) 刪掉中間段（從 A 到 B 不要）  ✅你要這個")
    mode = input("請輸入 1 或 2（預設 1）：").strip() or "1"
    if mode not in ("1", "2"):
        print("❌ 只能輸入 1 或 2")
        sys.exit(1)

    out_video = video.with_name(video.stem + "_cut.mp4")
    out_srt = srt.with_name(srt.stem + "_cut.srt")  # 若沒字幕就不會用到

    try:
        if mode == "1":
            front_in = input("請輸入【前面】要修剪的時間（mm:ss 或 秒，預設 0）：")
            back_in  = input("請輸入【後面】要修剪的時間（mm:ss 或 秒，預設 0）：")
            cut_front = parse_time_input(front_in)
            cut_back  = parse_time_input(back_in)

            cut_video_front_back(video, out_video, cut_front, cut_back)

            if srt.exists():
                if cut_front > 0:
                    shift_srt_all(srt, out_srt, cut_front)
                else:
                    out_srt.write_text(srt.read_text(encoding="utf-8-sig"), encoding="utf-8")
                print("📝 輸出字幕：", out_srt.name)

        else:  # mode == "2" 刪中間段
            a_in = input("請輸入【開始 A】（mm:ss 或 秒）：")
            b_in = input("請輸入【結束 B】（mm:ss 或 秒）：")
            a = parse_time_input(a_in)
            b = parse_time_input(b_in)

            delete_video_middle_concat(video, out_video, a, b)

            if srt.exists():
                delete_srt_middle(srt, out_srt, a, b)
                print("📝 輸出字幕：", out_srt.name)

    except Exception as e:
        print("❌ 失敗：", e)
        sys.exit(1)

    print("\n✅ 完成")
    print("🎬 輸出影片：", out_video.name)

if __name__ == "__main__":
    main()
