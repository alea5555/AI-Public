@echo off
setlocal EnableExtensions EnableDelayedExpansion
REM --- 設置編碼為 UTF-8 (Code Page 65001) ---
chcp 65001 >nul
rem 在DOS下執行(檔案存在不轉換) cmd /c "chcp 65001 >nul && call "f:\F\AI\public\Whisper_指定.bat""
REM ===================================================
REM Whisper_指定.bat  (最終交付版)
REM 功能：
REM - 讓你輸入/拖曳「要辨識的目錄」
REM - 遞迴掃描 mp4/mp3
REM - 轉 WAV(16kHz/mono) -> whisper 產生 SRT
REM - 若同名 .srt 已存在則跳過（不轉WAV、不跑Whisper）
REM - 轉完刪除暫存 wav
REM 注意：
REM - 若你用啟動器 run_whisper.bat 啟動，會先 chcp 65001，中文路徑更穩
REM ===================================================

REM ===================================================
REM 你原本的工具路徑（不動）
REM ===================================================
set "WHISPER=C:\_install\Whispertool\main.exe"
set "MODEL=C:\_install\Whispertool\ggml-medium.bin"
set "FFMPEG=C:\_install\Whispertool\ffmpeg.exe"

REM ===================================================
REM 讓你輸入要辨識的目錄（可拖曳資料夾進來）
REM 預設：f:\F\作者
REM ===================================================
set "DEFAULT_DIR=f:\F\AI\downloads"

echo.
echo ===================================================
echo 請輸入要辨識的目錄（可直接拖曳資料夾進來）
echo 直接按 Enter 使用預設：%DEFAULT_DIR%
echo ===================================================
set /p "INPUT_DIR=> "

if "%INPUT_DIR%"=="" (
  set "ROOT=%DEFAULT_DIR%"
) else (
  set "ROOT=%INPUT_DIR%"
)

REM 去掉可能的雙引號（拖曳進來通常會有引號）
set "ROOT=%ROOT:"=%"

REM 轉到該目錄（確保存在）
cd /d "%ROOT%" 2>nul
if errorlevel 1 (
  echo [錯誤] 目錄不存在或無法進入：%ROOT%
  pause
  exit /b 1
)
set "ROOT=%CD%"

echo.
echo ===================================================
echo === Whisper 批量轉 SRT 字幕 (中文) 開始 ===
echo 檢查機制：若同名 .srt 檔案已存在，則跳過（不轉WAV、不跑Whisper）
echo 處理目錄：%ROOT%
echo ===================================================

REM ===================================================
REM 檢查工具是否存在
REM ===================================================
if not exist "%FFMPEG%" (
  echo [錯誤] 找不到 FFmpeg：%FFMPEG%
  pause
  exit /b 1
)
if not exist "%WHISPER%" (
  echo [錯誤] 找不到 Whisper main.exe：%WHISPER%
  pause
  exit /b 1
)
if not exist "%MODEL%" (
  echo [錯誤] 找不到模型檔：%MODEL%
  pause
  exit /b 1
)

REM ===================================================
REM 掃描檔案（遞迴）
REM ===================================================
set /a COUNT=0

for /r "%ROOT%" %%F in (*.mp4 *.mp3) do (
  set "SRT_TARGET=%%~dpnF.srt"
  if exist "!SRT_TARGET!" (
    echo [跳過] 已存在：!SRT_TARGET!
  ) else (
    set /a COUNT+=1
    call :PROCESS "%%~fF"
  )
)

echo.
echo ==========================================
echo === 全部完成，實際轉檔 %COUNT% 個檔案 ===
echo ==========================================
pause
exit /b


REM ===================================================
REM 子程序：處理單一檔案
REM ===================================================
:PROCESS
set "IN=%~1"
set "DIR=%~dp1"
set "BASE=%~n1"

set "WAV=%DIR%%BASE%.wav"
set "SRT_TARGET=%DIR%%BASE%.srt"

echo.
echo -------------------------------------------------------
echo [轉檔 %COUNT%] 輸入檔: %IN%
echo -------------------------------------------------------

REM [1/2] 轉 WAV（16kHz, mono, PCM s16le）
echo [1/2] 轉換 WAV 檔案中...
"%FFMPEG%" -y -i "%IN%" -ar 16000 -ac 1 -c:a pcm_s16le "%WAV%" >nul 2>&1
if errorlevel 1 (
  echo [失敗] FFmpeg 轉 WAV 失敗。
  if exist "%WAV%" del /q "%WAV%"
  goto :eof
)

REM [2/2] Whisper 轉錄輸出 SRT
echo [2/2] 執行 Whisper 轉錄中文 SRT 中...
pushd "%DIR%"
"%WHISPER%" -m "%MODEL%" -l zh -t 4 -osrt -f "%WAV%" >nul 2>&1
popd

if errorlevel 1 (
  echo [失敗] Whisper 轉錄失敗。
  if exist "%WAV%" del /q "%WAV%"
  goto :eof
)

REM 確認 SRT 是否產生
if exist "%SRT_TARGET%" (
  echo [完成] SRT 已生成: "%SRT_TARGET%"
) else (
  echo [警告] 未找到輸出的 SRT: "%SRT_TARGET%"
)

REM 清理暫存 WAV
if exist "%WAV%" (
  del /q "%WAV%"
  echo [清理] 暫存 WAV 已刪除。
)

goto :eof
