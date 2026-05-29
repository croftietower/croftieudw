#!/bin/bash
# macOS 一键打包脚本

echo "===================================="
echo "  Bidirectional Translator v1.0.0 打包工具"
echo "===================================="
echo ""

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
echo "工作目录: $SCRIPT_DIR"
cd "$SCRIPT_DIR"
echo ""

# 检查是否在正确目录
if [ ! -f "translation_app.py" ]; then
    echo "错误: 找不到 translation_app.py！"
    ls -la
    read -p "按任意键退出..."
    exit 1
fi

# 检查是否有 miniforge（优先使用）
if [ -d "miniforge" ] && [ -f "miniforge/bin/python" ]; then
    PYTHON_CMD="./miniforge/bin/python"
    echo "使用项目内置的 miniforge"
else
    echo "检查系统 Python..."
    if command -v python3 &>/dev/null; then
        PYTHON_CMD="python3"
    elif command -v python &>/dev/null; then
        PYTHON_CMD="python"
    else
        echo "错误: 未找到 Python！"
        read -p "按任意键退出..."
        exit 1
    fi
fi

echo ""
$PYTHON_CMD --version
echo ""

# 检查和安装依赖
echo "检查依赖..."
$PYTHON_CMD -c "import PyInstaller" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "安装 PyInstaller..."
    $PYTHON_CMD -m pip install PyInstaller
fi

$PYTHON_CMD -c "import requests" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "安装 requests..."
    $PYTHON_CMD -m pip install requests
fi

$PYTHON_CMD -c "import docx" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "安装 python-docx..."
    $PYTHON_CMD -m pip install python-docx
fi

echo ""

# 清理旧文件
echo "清理旧文件..."
rm -rf build dist __pycache__ *.spec 2>/dev/null

# 设置临时缓存目录
export PYINSTALLER_CACHE_DIR="$(pwd)/pyi_temp_cache"
mkdir -p "$PYINSTALLER_CACHE_DIR"

echo ""
echo "开始打包 (single-file + GUI 模式)..."
echo "===================================="

$PYTHON_CMD -m PyInstaller \
    --onefile \
    --windowed \
    --name "Bidirectional Translator v1.0.0" \
    --clean \
    --noconfirm \
    translation_app.py

if [ -f "dist/Bidirectional Translator v1.0.0" ]; then
    echo ""
    echo "===================================="
    echo "  打包成功！"
    echo "===================================="
    echo ""
    echo "可执行文件位置: dist/Bidirectional Translator v1.0.0"
    ls -lh "dist/Bidirectional Translator v1.0.0"
    echo ""
    echo "你可以直接运行它或压缩分享！"
    echo ""
    
    # 清理临时缓存
    rm -rf "$PYINSTALLER_CACHE_DIR"
    
    read -p "是否在 Finder 中打开？(Y/n) " choice
    case "$choice" in 
        y|Y|"" ) open "dist" ;;
        * ) echo "好的，文件在 dist/ 目录下。" ;;
    esac
else
    echo ""
    echo "===================================="
    echo "  打包失败，请检查错误信息。"
    echo "===================================="
    echo ""
    echo "尝试备用打包方案（目录模式）..."
    echo ""
    
    # 备用方案 - 使用 onedir 模式
    $PYTHON_CMD -m PyInstaller \
        --onedir \
        --windowed \
        --name "Bidirectional Translator v1.0.0" \
        --clean \
        --noconfirm \
        translation_app.py
        
    if [ -d "dist/Bidirectional Translator v1.0.0.app" ]; then
        echo ""
        echo "备用方案打包成功！"
        echo "应用位置: dist/Bidirectional Translator v1.0.0.app"
        open "dist"
    fi
fi

echo ""
echo "完成！"
