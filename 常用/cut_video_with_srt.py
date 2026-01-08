import subprocess
from pathlib import Path
import re
import sys

# =========================
# 固定目錄
# =========================
BASE_DIR = Path(r"F:\F\AI\downloads")

# =========================
# SRT 工具
# =========================
TIME_PATTERN = re.compile(r"(\d{2}):(\d{2}):(\d{2}),(\d{3})")

def time_to_ms(h, m, s, ms):
    return int(h)*3600000 + int(m)*60000 + int(s)*1000 + int(ms)

def ms_to_time(ms):
    if ms < 0:
        ms = 0
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms %= 1000
    return f"{h:02}:{m:02}:{s:02},{ms:03}"

def shift_srt(input_srt, output_srt, shift_seconds):
    shift_ms = int(shift_seconds * 1000)
    text = input_srt.read_text(encoding="utf-8-sig")

    def repl(m):
        ms = time_to_ms(*m.groups())
        return ms_to_time(ms - shift_ms)

    output_srt.write_text(
        TIME_PATTERN.sub(repl, text),
        encoding="utf-8"
    )

# =========================
# 影片工具
# =========================
def get_duration(video):
    cmd = [
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        str(video)
    ]
    return float(subprocess.check_output(cmd).decode().strip())

def cut_video(input_video, output_video, cut_front, cut_back):
    duration = get_duration(input_video)
    keep_start = cut_front
    keep_length = duration - cut_front - cut_back

    if keep_length <= 0:
        raise ValueError("修剪後影片長度 <= 0，請檢查秒數")

    cmd = [
        "ffmpeg", "-y",
        "-ss", str(keep_start),
        "-i", str(input_video),
        "-t", str(keep_length),
        "-c", "copy",
        "-movflags", "+faststart",
        str(output_video)
    ]

    subprocess.run(cmd, check=True)

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

    if srt.exists():
        print("📝 字幕：", srt.name)
    else:
        print("📝 字幕：無（將略過字幕處理）")

    front = input("請輸入【前面】要修剪的秒數（預設 0）：").strip()
    back  = input("請輸入【後面】要修剪的秒數（預設 0）：").strip()

    cut_front = float(front) if front else 0.0
    cut_back  = float(back) if back else 0.0

    out_video = video.with_name(video.stem + "_cut.mp4")

    # 1️⃣ 剪影片
    cut_video(video, out_video, cut_front, cut_back)

    # 2️⃣ 有 SRT 才處理字幕
    if srt.exists():
        out_srt = srt.with_name(srt.stem + "_cut.srt")
        if cut_front > 0:
            shift_srt(srt, out_srt, cut_front)
        else:
            out_srt.write_text(
                srt.read_text(encoding="utf-8-sig"),
                encoding="utf-8"
            )
        print("📝 輸出字幕：", out_srt.name)

    print("\n✅ 完成")
    print("🎬 輸出影片：", out_video.name)

if __name__ == "__main__":
    main()
