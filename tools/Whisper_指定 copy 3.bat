@echo off
REM --- 自動修正 VS Code 可能產生的錯誤引號轉義 ---
set "FIXED_PATH=%~f0"
set "FIXED_PATH=%FIXED_PATH:\"=%"

REM --- 切換到腳本所在目錄 ---
cd /d "%~dp0"
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

REM ===================================================
REM Whisper_指定.bat (VS Code 強效相容版)
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

REM 去掉輸入路徑中的雙引號
set "ROOT=%ROOT:"=%"
cd /d "%ROOT%" 2>nul
set "ROOT=%CD%"

echo.
echo 處理目錄：%ROOT%
echo ===================================================

REM --- 掃描檔案 (支援 .m4a) ---
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
echo.
echo [轉檔 %COUNT%] 輸入檔: %IN%
"%FFMPEG%" -y -i "%IN%" -ar 16000 -ac 1 -c:a pcm_s16le "%WAV%" >nul 2>&1
pushd "%DIR%"
"%WHISPER%" -m "%MODEL%" -l zh -t 4 -osrt -f "%WAV%" >nul 2>&1
popd
if exist "%WAV%" del /q "%WAV%"
goto :eof