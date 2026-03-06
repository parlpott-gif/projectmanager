@echo off
chcp 65001 >nul
title Hardware Project Manager
cd /d "%~dp0"

REM 检查并安装必要的依赖
python -c "import flask" 2>nul
if errorlevel 1 (
    echo Installing dependencies...
    pip install flask xlrd xlwt xlutils pandas python-docx -q
)

REM 验证新增的重命名模块是否可用
python -c "from server import get_auto_renamed_file, get_numbered_image_name" 2>nul
if errorlevel 1 (
    echo ERROR: Auto-rename functions not found in server.py
    echo Please check if server.py has been updated correctly.
    pause
    exit /b 1
)

REM 自动打开浏览器
start "" cmd /c "timeout /t 3 >nul && start http://localhost:5000"

echo Starting Hardware Project Manager with Auto-Rename Feature...
echo.
echo [INFO] Server starting at http://127.0.0.1:5000
echo [INFO] File auto-rename feature is active
echo [INFO] Press Ctrl+C to stop the server
echo.

REM 启动应用
python app_main.py
pause
