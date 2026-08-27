@echo off
chcp 65001 >nul
setlocal

echo ==========================================================
echo  项目运行依赖安装脚本
echo ==========================================================
echo.

rem ---------- 1. 定位 Python ----------
set "PYTHON="
if exist "%~dp0python\python.exe" (
    set "PYTHON=%~dp0python\python.exe"
    echo [1/3] 使用项目自带 Python: %PYTHON%
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON=python"
        echo [1/3] 使用系统 Python
    ) else (
        echo [错误] 未找到 Python，请先安装 Python 3.10+ 或放置嵌入式运行时到 python\ 目录。
        pause
        exit /b 1
    )
)

rem ---------- 2. 安装 Python 依赖 ----------
echo.
echo [2/3] 安装 Python 依赖 (pip install -r requirements.txt) ...
"%PYTHON%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo [警告] 依赖安装可能未成功，请检查网络或 pip 配置。
)

rem ---------- 3. 检查外部运行时 ----------
echo.
echo [3/3] 检查外部运行时 (Tesseract OCR / Poppler) ...
set "MISSING="

set "TESS_CMD=%~dp0Tesseract-OCR\tesseract.exe"
if not exist "%TESS_CMD%" (
    set "MISSING=1"
    echo   [缺失] Tesseract OCR 未找到。
    echo          请从官方下载并放置到: %~dp0Tesseract-OCR\
    echo          官网: https://github.com/UB-Mannheim/tesseract/wiki
)

set "POPPLER_BIN=%~dp0poppler-25.12.0\Library\bin\pdftoppm.exe"
if not exist "%POPPLER_BIN%" (
    set "MISSING=1"
    echo   [缺失] Poppler 未找到 (任意 poppler-* 版本目录)。
    echo          请下载 Poppler for Windows 并解压到项目根目录。
    echo          参考: https://github.com/oschwartz10612/poppler-windows/releases
)

echo.
if defined MISSING (
    echo 依赖检查完成，但存在缺失项，请按上述提示补齐后重试。
) else (
    echo 所有依赖已就绪。
)
echo.
pause
endlocal