import os
import time
import subprocess
from datetime import datetime

import numpy as np
import mss


def pick_monitor(monitors):
    """
    mss.monitors:
      index 0 = all monitors combined (virtual screen)
      index 1..N = each monitor
    """
    count = len(monitors) - 1
    if count <= 0:
        raise RuntimeError("沒有偵測到任何螢幕。")

    print(f"偵測到 {count} 個螢幕：")
    for i in range(1, len(monitors)):
        m = monitors[i]
        print(f"  螢幕 {i}: {m['width']}x{m['height']} (left={m['left']}, top={m['top']})")

    while True:
        s = input(f"請選擇要錄影的螢幕編號 (1~{count})：").strip()
        if s.isdigit() and 1 <= int(s) <= count:
            return int(s)
        print("輸入不正確，請重試。")


def main():
    print("=== Screen Recorder (FFmpeg → MP4 H.264) ===")

    fps_str = input("FPS？(直接 Enter=30)：").strip()
    fps = 30 if fps_str == "" else int(fps_str)

    dur_str = input("要錄幾秒？(直接 Enter=手動按 q 停止)：").strip()
    duration = None
    if dur_str:
        duration = float(dur_str)
        if duration <= 0:
            duration = None

    # 輸出位置：目前執行的資料夾
    out_dir = os.getcwd()

    with mss.mss() as sct:
        idx = pick_monitor(sct.monitors)
        mon = sct.monitors[idx]
        w, h = mon["width"], mon["height"]

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(out_dir, f"screen{idx}_{ts}.mp4")

        print("\n==========")
        print(f"輸出檔案：{out_path}")
        print(f"解析度：{w}x{h}  FPS={fps}")
        print("操作：在這個 CMD 視窗按 q 停止（或 Ctrl+C）")
        if duration is not None:
            print(f"錄影時間：{duration} 秒")
        print("==========\n")

        # ffmpeg 從 stdin 吃 rawvideo (bgr24) → H.264 MP4
        cmd = [
            "ffmpeg",
            "-y",
            "-hide_banner",
            "-loglevel", "warning",
            "-f", "rawvideo",
            "-pix_fmt", "bgr24",
            "-s", f"{w}x{h}",
            "-r", str(fps),
            "-i", "-",
            "-an",
            "-c:v", "libx264",
            "-preset", "veryfast",
            "-crf", "23",
            "-pix_fmt", "yuv420p",
            out_path
        ]

        proc = subprocess.Popen(cmd, stdin=subprocess.PIPE)

        # Windows CMD 讀鍵盤：按 q 停止
        import msvcrt

        start = time.time()
        frame_interval = 1.0 / fps
        next_t = start
        last_print = -1

        try:
            while True:
                now = time.time()
                elapsed = now - start

                if duration is not None and elapsed >= duration:
                    break

                # 按 q 停止
                if msvcrt.kbhit():
                    ch = msvcrt.getwch()
                    if ch.lower() == "q":
                        break

                # 控 FPS
                if now < next_t:
                    time.sleep(max(0, next_t - now))
                next_t += frame_interval

                img = sct.grab(mon)  # BGRA
                frame = np.asarray(img, dtype=np.uint8)[:, :, :3]  # BGR
                proc.stdin.write(frame.tobytes())

                # 每秒印一次狀態（可確認有在錄）
                sec = int(elapsed)
                if sec != last_print:
                    last_print = sec
                    print(f"錄影中… {sec} 秒", end="\r", flush=True)

        except KeyboardInterrupt:
            pass
        finally:
            print("\n停止中，正在封裝 MP4…")
            try:
                proc.stdin.close()
            except Exception:
                pass
            proc.wait()

        print("完成。")

if __name__ == "__main__":
    # 依賴：pip install mss numpy
    main()
