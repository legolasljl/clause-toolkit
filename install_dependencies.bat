@echo off
chcp 65001 >nul 2>&1
title 智能条款工具箱 - 安装依赖

echo ============================================
echo   智能条款工具箱 V18.9 - 依赖安装脚本
echo ============================================
echo.

:: 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo [1/3] 升级 pip...
python -m pip install --upgrade pip

echo.
echo [2/3] 安装核心依赖...
pip install PyQt5>=5.15 ^
            pandas>=1.5 ^
            openpyxl>=3.0 ^
            python-docx>=0.8

echo.
echo [3/3] 安装可选依赖（增强功能）...
pip install jieba ^
            scikit-learn ^
            numpy ^
            pdfplumber ^
            PyPDF2 ^
            deep-translator

echo.
echo ============================================
echo   安装完成！
echo ============================================
echo.
echo 核心依赖: PyQt5, pandas, openpyxl, python-docx
echo 可选依赖: jieba(分词), scikit-learn(智能匹配),
echo           pdfplumber(PDF), deep-translator(翻译)
echo.
echo 如需转换 .doc 文件，请安装 LibreOffice:
echo https://www.libreoffice.org/download/
echo.
pause
