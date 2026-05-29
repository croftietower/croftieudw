@echo off
chcp 65001 >nul
title Bidirectional Translator Packager
color 0A
echo ====================================
echo   Bidirectional Translator 打包工具
echo   Bidirectional Translator v1.0.0
echo ====================================
echo.

REM 先让窗口停留，让用户看到
echo 按任意键开始打包...
pause >nul
cls

REM 获取脚本所在目录
cd /d "%~dp0"
echo 工作目录: %CD%
echo.

REM 检查是否在正确目录
if not exist "translation_app.py" (
    echo.
    echo [错误] 找不到 translation_app.py!
    echo.
    echo 请确保以下两个文件在同一个文件夹：
    echo   - translation_app.py
    echo   - package_windows.bat
    echo.
    echo 当前目录内容：
    dir /b
    echo.
    pause
    exit /b 1
)

echo [OK] 找到 translation_app.py
echo.

REM 检查 Python
echo 检查 Python 环境...
echo.

where python >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :found_python
)

where python3 >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python3
    goto :found_python
)

REM 如果都没找到，尝试常见安装路径
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
    goto :found_python
)
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
    set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python310\python.exe
    goto :found_python
)
if exist "C:\Python311\python.exe" (
    set PYTHON_CMD=C:\Python311\python.exe
    goto :found_python
)
if exist "C:\Python310\python.exe" (
    set PYTHON_CMD=C:\Python310\python.exe
    goto :found_python
)

REM 确实没找到Python
echo.
echo [错误] 未找到 Python!
echo.
echo 请按以下步骤安装 Python：
echo.
echo 1. 访问 https://www.python.org/downloads/
echo 2. 下载并安装 Python 3.10 或更高版本
echo 3. [重要!] 安装时务必勾选 "Add Python to PATH"
echo 4. 安装完成后重新运行此脚本
echo.
pause
exit /b 1

:found_python
echo [OK] 找到 Python
%PYTHON_CMD% --version
echo.

REM 检查依赖
echo 检查依赖...

echo   - PyInstaller...
%PYTHON_CMD% -c "import PyInstaller" 2>nul
if %errorlevel% neq 0 (
    echo     正在安装 PyInstaller...
    %PYTHON_CMD% -m pip install PyInstaller
)

echo   - requests...
%PYTHON_CMD% -c "import requests" 2>nul
if %errorlevel% neq 0 (
    echo     正在安装 requests...
    %PYTHON_CMD% -m pip install requests
)

echo   - python-docx...
%PYTHON_CMD% -c "import docx" 2>nul
if %errorlevel% neq 0 (
    echo     正在安装 python-docx...
    %PYTHON_CMD% -m pip install python-docx
)

echo [OK] 所有依赖已就绪
echo.

REM 清理旧文件
echo 清理旧文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
del /q *.spec 2>nul
echo [OK] 清理完成
echo.

echo.
echo ====================================
echo   开始打包...
echo ====================================
echo.

%PYTHON_CMD% -m PyInstaller --onefile --noconsole --name "Bidirectional Translator v1.0.0" --clean --noconfirm translation_app.py

if %errorlevel% equ 0 (
    if exist "dist\Bidirectional Translator v1.0.0.exe" (
        echo.
        echo ====================================
        echo   打包成功！
        echo ====================================
        echo.
        echo 可执行文件位置: dist\Bidirectional Translator v1.0.0.exe
        echo.
        dir "dist\Bidirectional Translator v1.0.0.exe"
        echo.
        echo 你可以直接运行它或压缩分享！
        echo.
        
        set /p choice="是否打开输出目录？ (Y/N): "
        if /i "%choice%"=="Y" explorer dist
    ) else (
        echo.
        echo [提示] 打包完成，但未找到可执行文件
        echo.
        echo 检查 dist 目录...
        dir dist
        echo.
        echo 如果看到 Bidirectional Translator v1.0.0.app 或其他文件夹，
        echo 说明打包成功了，只是单文件模式失败。试试用文件夹模式。
        echo.
        explorer dist
    )
) else (
    echo.
    echo ====================================
    echo   打包失败，尝试备用方案...
    echo ====================================
    echo.
    
    REM 备用方案 - onedir模式
    %PYTHON_CMD% -m PyInstaller --onedir --noconsole --name "Bidirectional Translator v1.0.0" --clean --noconfirm translation_app.py
    
    if %errorlevel% equ 0 (
        echo.
        echo 备用方案打包成功！
        echo 程序目录: dist\Bidirectional Translator v1.0.0
        echo.
        explorer dist
    )
)

echo.
echo 完成！按任意键退出...
pause >nul
