@echo off
REM 自動切換到批次檔所在的磁碟與目錄
cd /d "%~dp0"
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

REM ===================================================
REM Whisper_指定.bat (強效相容版)
REM ===================================================

REM --- 工具路徑 ---
set "WHISPER=C:\_install\Whispertool\main.exe"
set "MODEL=C:\_install\Whispertool\ggml-medium.bin"
set "FFMPEG=C:\_install\Whispertool\ffmpeg.exe"

REM --- 預設目錄 ---
set "DEFAULT_DIR=f:\F\AI\downloads"

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

REM 去掉雙引號並轉換路徑
set "ROOT=%ROOT:"=%"
cd /d "%ROOT%" 2>nul
set "ROOT=%CD%"

echo.
echo 處理目錄：%ROOT%
echo ===================================================

REM --- 檢查工具 ---
if not exist "%FFMPEG%" ( echo [錯誤] 找不到 FFmpeg & pause & exit /b 1 )
if not exist "%WHISPER%" ( echo [錯誤] 找不到 Whisper & pause & exit /b 1 )
if not exist "%MODEL%" ( echo [錯誤] 找不到模型檔 & pause & exit /b 1 )

REM --- 掃描檔案 (修正後的 for 迴圈) ---
set /a COUNT=0

for /f "delims=" %%F in ('dir /b /s "%ROOT%\*.mp4" "%ROOT%\*.mp3" "%ROOT%\*.m4a" 2^>nul') do (
  set "SRT_TARGET=%%~dpnF.srt"
  if exist "!SRT_TARGET!" (
    echo [跳過] 已存在：!SRT_TARGET!
  ) else (
    set /a COUNT+=1
    call :PROCESS "%%F"
  )
)

echo.
echo === 全部完成，實際轉檔 %COUNT% 個檔案 ===
pause
exit /b

:PROCESS
set "IN=%~1"
set "DIR=%~dp1"
set "BASE=%~n1"
set "WAV=%DIR%%BASE%.wav"
set "SRT_TARGET=%DIR%%BASE%.srt"

echo.
echo [轉檔 %COUNT%] 輸入檔: %IN%

REM [1/2] 轉 WAV
"%FFMPEG%" -y -i "%IN%" -ar 16000 -ac 1 -c:a pcm_s16le "%WAV%" >nul 2>&1

REM [2/2] Whisper 轉錄
pushd "%DIR%"
"%WHISPER%" -m "%MODEL%" -l zh -t 4 -osrt -f "%WAV%" >nul 2>&1
popd

if exist "%WAV%" del /q "%WAV%"
goto :eof