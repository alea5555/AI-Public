import os
import time
import subprocess
from datetime import datetime

import numpy as np
import mss

AUDIO_DEVICE_NAME = "CABLE Output (VB-Audio Virtual Cable)"


def pick_monitor(monitors):
    count = len(monitors) - 1
    print(f"偵測到 {count} 個螢幕：")
    for i in range(1, len(monitors)):
        m = monitors[i]
        print(f"  螢幕 {i}: {m['width']}x{m['height']}")
    while True:
        s = input(f"選擇螢幕 (1~{count})：").strip()
        if s.isdigit() and 1 <= int(s) <= count:
            return int(s)
        print("輸入不正確，請重試。")


def pick_resolution(w, h):
    print("\n選擇錄製解析度：")
    print(" 1) 原始解析度")
    print(" 2) 1920 x 1080")
    print(" 3) 1280 x 720")
    while True:
        s = input("輸入 1 / 2 / 3：").strip()
        if s == "1":
            return w, h
        if s == "2":
            return 1920, 1080
        if s == "3":
            return 1280, 720
        print("輸入不正確，請重試。")


def main():
    print("=== Screen Recorder (FFmpeg / MP4 H.264 + VB-Cable) ===")

    fps = int(input("FPS (Enter=30)：") or 30)
    dur_str = input("錄影秒數 (Enter=手動按 q)：").strip()
    duration = float(dur_str) if dur_str else None

    with mss.mss() as sct:
        idx = pick_monitor(sct.monitors)
        mon = sct.monitors[idx]
        src_w, src_h = mon["width"], mon["height"]
        out_w, out_h = pick_resolution(src_w, src_h)

        out_file = f"screen{idx}_{out_w}x{out_h}_{datetime.now():%Y%m%d_%H%M%S}.mp4"
        out_path = os.path.join(os.getcwd(), out_file)

        print("\n====================")
        print(f"輸出檔案：{out_path}")
        print(f"來源解析度：{src_w}x{src_h}  →  輸出：{out_w}x{out_h}")
        print(f"FPS：{fps}")
        print(f"系統聲音：{AUDIO_DEVICE_NAME}（Windows 預設輸出 = CABLE Input）")
        print("CMD 按 q 停止（或 Ctrl+C）")
        print("====================\n")

        cmd = [
            "ffmpeg", "-y",
            "-hide_banner", "-loglevel", "warning",

            # ★ 音訊：加大 queue + buffer，避免 too full / dropped
            "-thread_queue_size", "4096",
            "-rtbufsize", "512M",
            "-f", "dshow",
            "-i", f"audio={AUDIO_DEVICE_NAME}",

            # 影像：stdin rawvideo
            "-f", "rawvideo",
            "-pix_fmt", "bgr24",
            "-s", f"{src_w}x{src_h}",
            "-r", str(fps),
            "-i", "-",

            "-map", "1:v:0",
            "-map", "0:a:0",

            "-vf", f"scale={out_w}:{out_h}",
            "-c:v", "libx264",
            "-preset", "veryfast",
            "-crf", "23",
            "-pix_fmt", "yuv420p",

            # ★ 音訊：async 平滑，避免抖動/不同步
            "-af", "aresample=async=1:first_pts=0",
            "-c:a", "aac",
            "-b:a", "192k",

            "-shortest",
            out_path
        ]

        proc = subprocess.Popen(cmd, stdin=subprocess.PIPE)
        import msvcrt

        start = time.time()
        next_t = start
        frame_dt = 1.0 / fps
        last_print_sec = -1

        try:
            while True:
                now = time.time()
                elapsed = now - start

                if duration is not None and elapsed >= duration:
                    break

                if msvcrt.kbhit():
                    ch = msvcrt.getwch()
                    if ch.lower() == "q":
                        break

                if now < next_t:
                    time.sleep(max(0, next_t - now))
                next_t += frame_dt

                img = sct.grab(mon)
                frame = np.asarray(img, dtype=np.uint8)[:, :, :3]
                proc.stdin.write(frame.tobytes())

                sec = int(elapsed)
                if sec != last_print_sec:
                    last_print_sec = sec
                    print(f"錄影中… {sec} 秒", end="\r", flush=True)

        except KeyboardInterrupt:
            pass
        finally:
            print("\n停止錄影，封裝 MP4 中…")
            try:
                proc.stdin.close()
            except Exception:
                pass
            proc.wait()

        print("完成。")


if __name__ == "__main__":
    # 依賴：pip install mss numpy
    main()
