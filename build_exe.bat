@echo off
chcp 65001 >nul 2>&1
title 智能条款工具箱 - 打包为 EXE

echo ============================================
echo   智能条款工具箱 V18.9 - PyInstaller 打包
echo ============================================
echo.

:: 切换到脚本所在目录
cd /d "%~dp0"

:: 检查 PyInstaller
pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo [提示] 正在安装 PyInstaller...
    pip install pyinstaller
)

:: 检查图标文件
set ICON_FILE=app.ico
if not exist "%ICON_FILE%" (
    echo [警告] 未找到 %ICON_FILE%，将使用默认图标
    echo 请将 .ico 图标文件放在当前目录并命名为 app.ico
    set ICON_PARAM=
) else (
    set ICON_PARAM=--icon=%ICON_FILE%
)

echo.
echo [1/2] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [2/2] 开始打包...
echo.

pyinstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name "智能条款工具箱" ^
    %ICON_PARAM% ^
    --add-data "Property.json;." ^
    --add-data "Liability.json;." ^
    --add-data "clause_mapping_manager.py;." ^
    --add-data "clause_mapping_dialog.py;." ^
    --add-data "insurance_calculator.py;." ^
    --add-data "wx.jpg;." ^
    --add-data "zfb.jpg;." ^
    --hidden-import=jieba ^
    --hidden-import=sklearn ^
    --hidden-import=sklearn.feature_extraction.text ^
    --hidden-import=sklearn.metrics.pairwise ^
    --hidden-import=pdfplumber ^
    --hidden-import=PyPDF2 ^
    --hidden-import=deep_translator ^
    --hidden-import=numpy ^
    --collect-data=jieba ^
    Clause_Comparison_Assistant_windows.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请检查错误信息
    pause
    exit /b 1
)

echo.
echo ============================================
echo   打包完成！
echo ============================================
echo.
echo 输出目录: dist\智能条款工具箱\
echo 主程序:   dist\智能条款工具箱\智能条款工具箱.exe
echo.
echo 分发时请将整个 "智能条款工具箱" 文件夹打包为 ZIP
echo.

:: 复制额外资源到dist目录
echo 复制额外资源文件...
if exist "excluded_titles.json" copy /y "excluded_titles.json" "dist\智能条款工具箱\" >nul
if exist "clause_mappings.json" copy /y "clause_mappings.json" "dist\智能条款工具箱\" >nul

echo.
echo 全部完成！按任意键打开输出目录...
pause
explorer "dist\智能条款工具箱"
