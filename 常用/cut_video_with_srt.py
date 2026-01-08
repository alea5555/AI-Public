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

def cut_video(input_video, output_video, cut_seconds, mode):
    duration = get_duration(input_video)

    if cut_seconds <= 0 or cut_seconds >= duration:
        raise ValueError("修剪秒數不合理")

    if mode == "front":
        cmd = [
            "ffmpeg", "-y",
            "-ss", str(cut_seconds),
            "-i", str(input_video),
            "-c", "copy",
            "-movflags", "+faststart",
            str(output_video)
        ]
    else:  # back
        keep = duration - cut_seconds
        cmd = [
            "ffmpeg", "-y",
            "-i", str(input_video),
            "-t", str(keep),
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

    if not srt.exists():
        print("❌ 找不到對應 SRT：", srt.name)
        sys.exit(1)

    print("📂 目錄：", BASE_DIR)
    print("🎬 影片：", video.name)
    print("📝 字幕：", srt.name)

    mode = input("請輸入修剪方向（front / back）：").strip().lower()
    if mode not in ("front", "back"):
        print("❌ 只能輸入 front 或 back")
        sys.exit(1)

    cut_seconds = float(input("請輸入修剪秒數（例如 125）："))

    out_video = video.with_name(video.stem + "_cut.mp4")
    out_srt   = srt.with_name(srt.stem + "_cut.srt")

    # 1️⃣ 剪影片
    cut_video(video, out_video, cut_seconds, mode)

    # 2️⃣ 處理字幕
    if mode == "front":
        shift_srt(srt, out_srt, cut_seconds)
    else:
        out_srt.write_text(srt.read_text(encoding="utf-8-sig"), encoding="utf-8")

    print("\n✅ 完成")
    print("🎬 輸出影片：", out_video.name)
    print("📝 輸出字幕：", out_srt.name)

if __name__ == "__main__":
    main()
