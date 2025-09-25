@echo off
chcp 65001 >nul
echo ========================================
echo Excel Function Maker / Excel函数制作器
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found!
    echo 错误：未找到Python！
    echo.
    echo Please install Python 3.6+ from https://python.org
    echo 请从 https://python.org 安装Python 3.6+
    pause
    exit /b 1
)

echo Python found! Checking dependencies...
echo 找到Python！正在检查依赖...

REM Check if pyperclip is installed
pip show pyperclip >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing dependencies...
    echo 正在安装依赖...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo Failed to install dependencies!
        echo 依赖安装失败！
        pause
        exit /b 1
    )
)

echo Starting Excel Function Maker...
echo 正在启动Excel函数制作器...
echo.

python excel_function_maker.py

if %errorlevel% neq 0 (
    echo.
    echo Application exited with error. Press any key to close.
    echo 应用程序出错退出。按任意键关闭。
    pause
)
