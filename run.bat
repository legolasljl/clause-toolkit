@echo off
chcp 65001 >nul 2>&1
title 智能条款工具箱 V18.9

:: 切换到脚本所在目录
cd /d "%~dp0"

python Clause_Comparison_Assistant_windows.py
if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出，请检查是否已安装依赖
    echo 运行 install_dependencies.bat 安装依赖
    pause
)
