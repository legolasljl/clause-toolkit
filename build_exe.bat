@echo off
chcp 65001 >nul
title Smart Clause Toolbox - Build EXE

echo ============================================
echo   Smart Clause Toolbox V18.9 - PyInstaller
echo ============================================
echo.

:: Switch to script directory
cd /d "%~dp0"

:: Check PyInstaller
pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing PyInstaller...
    pip install pyinstaller
)

:: Check icon file
set ICON_FILE=app.ico
if not exist "%ICON_FILE%" (
    echo [WARN] %ICON_FILE% not found, using default icon
    echo Place .ico file in current directory as app.ico
    set ICON_PARAM=
) else (
    set ICON_PARAM=--icon=%ICON_FILE%
)

echo.
echo [1/2] Cleaning old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

setlocal enabledelayedexpansion

:: Build --add-data list dynamically (skip missing files)
echo.
echo Checking data files...
set DATA_ARGS=
for %%F in (Property.json Liability.json clause_mapping_manager.py clause_mapping_dialog.py insurance_calculator.py wx.jpg zfb.jpg) do (
    if exist "%%F" (
        set "DATA_ARGS=!DATA_ARGS! --add-data %%F;."
        echo   [OK] %%F
    ) else (
        echo   [SKIP] %%F not found
    )
)

echo.
echo [2/2] Building...
echo.

pyinstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --noupx ^
    --name "SmartClauseToolbox" ^
    %ICON_PARAM% ^
    !DATA_ARGS! ^
    --hidden-import=jieba ^
    --hidden-import=sklearn ^
    --hidden-import=sklearn.feature_extraction.text ^
    --hidden-import=sklearn.metrics.pairwise ^
    --hidden-import=pdfplumber ^
    --hidden-import=PyPDF2 ^
    --hidden-import=numpy ^
    --collect-data=jieba ^
    Clause_Comparison_Assistant_windows.py

endlocal

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed, check error messages above
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build complete!
echo ============================================
echo.
echo Output: dist\SmartClauseToolbox\
echo EXE:    dist\SmartClauseToolbox\SmartClauseToolbox.exe
echo.
echo Please zip the entire "SmartClauseToolbox" folder for distribution
echo.

:: Copy extra resources to dist
echo Copying extra resource files...
if exist "excluded_titles.json" copy /y "excluded_titles.json" "dist\SmartClauseToolbox\" >nul
if exist "clause_mappings.json" copy /y "clause_mappings.json" "dist\SmartClauseToolbox\" >nul

echo.
echo Done! Press any key to open output directory...
pause
explorer "dist\SmartClauseToolbox"
