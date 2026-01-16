@echo off
REM --- 自動修正 VS Code 播放鍵產生的路徑跳脫錯誤 ---
set "SCRIPT_PATH=%~f0"
set "SCRIPT_PATH=%SCRIPT_PATH:\"=%"

REM --- 自動定位到腳本所在目錄 ---
cd /d "%~dp0"
setlocal EnableExtensions EnableDelayedExpansion
[cite_start]chcp 65001 >nul [cite: 1]

REM ===================================================
REM Whisper_指定.bat (VS Code 強效修正版)
REM ===================================================

REM --- 工具路徑 ---
[cite_start]set "WHISPER=C:\_install\Whispertool\main.exe" [cite: 1]
[cite_start]set "MODEL=C:\_install\Whispertool\ggml-medium.bin" [cite: 1]
[cite_start]set "FFMPEG=C:\_install\Whispertool\ffmpeg.exe" [cite: 1]

REM --- 預設目錄 ---
[cite_start]set "DEFAULT_DIR=f:\F\AI\downloads" [cite: 1]

echo ===================================================
echo 請輸入要辨識的目錄（可直接拖曳資料夾進來）
echo 直接按 Enter 使用預設：%DEFAULT_DIR%
echo ===================================================
[cite_start]set /p "INPUT_DIR=> " [cite: 1]

if "%INPUT_DIR%"=="" (
  [cite_start]set "ROOT=%DEFAULT_DIR%" [cite: 1]
) else (
  [cite_start]set "ROOT=%INPUT_DIR%" [cite: 1]
)

[cite_start]REM 標準化路徑並支援 m4a [cite: 1, 2]
[cite_start]set "ROOT=%ROOT:"=%" [cite: 1]
[cite_start]cd /d "%ROOT%" 2>nul [cite: 1]
[cite_start]set "ROOT=%CD%" [cite: 1]

echo.
echo 處理目錄：%ROOT%
echo ===================================================

REM --- 掃描檔案 (支援 .m4a 且解決特殊字元問題) ---
[cite_start]set /a COUNT=0 [cite: 1]
for /f "delims=" %%F in ('dir /b /s "%ROOT%\*.mp4" "%ROOT%\*.mp3" "%ROOT%\*.m4a" 2^>nul') do (
  [cite_start]set "SRT_TARGET=%%~dpnF.srt" [cite: 1]
  if exist "!SRT_TARGET!" (
    [cite_start]echo [跳過] 已存在：!SRT_TARGET! [cite: 1]
  ) else (
    [cite_start]set /a COUNT+=1 [cite: 1]
    [cite_start]call :PROCESS "%%F" [cite: 1]
  )
)

echo.
[cite_start]echo === 全部完成，實際轉檔 %COUNT% 個檔案 === [cite: 1]
pause
exit /b

:PROCESS
[cite_start]set "IN=%~1" [cite: 1]
[cite_start]set "DIR=%~dp1" [cite: 1]
[cite_start]set "BASE=%~n1" [cite: 1]
[cite_start]set "WAV=%DIR%%BASE%.wav" [cite: 1]
echo.
[cite_start]echo [轉檔 %COUNT%] 輸入檔: %IN% [cite: 1]
[cite_start]"%FFMPEG%" -y -i "%IN%" -ar 16000 -ac 1 -c:a pcm_s16le "%WAV%" >nul 2>&1 [cite: 1]
[cite_start]pushd "%DIR%" [cite: 1]
[cite_start]"%WHISPER%" -m "%MODEL%" -l zh -t 4 -osrt -f "%WAV%" >nul 2>&1 [cite: 1]
[cite_start]popd [cite: 1]
[cite_start]if exist "%WAV%" del /q "%WAV%" [cite: 1]
[cite_start]goto :eof [cite: 1]