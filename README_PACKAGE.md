# Bidirectional Translator - 打包说明

## 📦 快速打包方法

### macOS 用户

#### 方法一：使用一键脚本（推荐）

1. 打开终端，进入项目目录
2. 运行：
   ```bash
   chmod +x package_mac.sh
   ./package_mac.sh
   ```

#### 方法二：手动打包

```bash
# 1. 安装依赖
pip3 install pyinstaller requests python-docx

# 2. 打包成单个文件
pyinstaller --onefile --windowed --name "BidirectionalTranslator" translation_app.py

# 3. 查看结果
ls -lh dist/
```

#### 方法三：使用 conda 环境（如果方法一二失败）

```bash
# 使用项目目录下已安装的 miniforge
./miniforge/bin/python -m pip install pyinstaller

# 打包
./miniforge/bin/python -m PyInstaller --onefile --windowed --name "BidirectionalTranslator" translation_app.py
```

---

### Windows 用户

#### 方法一：使用一键脚本（推荐）

1. 双击运行 `package_windows.bat`
2. 等待打包完成

#### 方法二：手动打包

```cmd
REM 安装依赖
pip install pyinstaller requests python-docx

REM 打包
pyinstaller --onefile --noconsole --name "BidirectionalTranslator" translation_app.py
```

---

## 🔧 常见问题

### macOS 权限问题

如果遇到权限错误：

```bash
# 清除 PyInstaller 缓存
rm -rf ~/Library/Application\ Support/pyinstaller

# 或者使用自定义缓存目录
mkdir -p ./pyi_cache
PYINSTALLER_CACHE_DIR="./pyi_cache" pyinstaller ...
```

### Windows 杀毒软件误报

PyInstaller 打包的程序有时会被杀毒软件误报，这是正常的。你可以：
1. 添加信任
2. 使用代码签名证书
3. 提供文件哈希值给用户验证

### 打包文件太大

这是正常的，因为包含了 Python 运行时和所有依赖库。通常在 50-100MB 之间。

---

## 📁 打包后文件

打包成功后，可执行文件在 `dist/` 目录下：

- **macOS**: `dist/BidirectionalTranslator` 或 `dist/BidirectionalTranslator.app`
- **Windows**: `dist/BidirectionalTranslator.exe`

---

## 🚀 分发应用

### macOS

1. 压缩 `dist/BidirectionalTranslator` 或 `.app` 包
2. 分享给其他用户
3. 用户首次打开可能需要右键 → 打开（绕过安全检查）

### Windows

1. 压缩 `dist/BidirectionalTranslator.exe`
2. 分享给其他用户
3. 用户可以直接运行

---

## 📝 更新日志

### v1.1 (最新)
- ✅ 添加了模型切换功能（DeepSeek V4 Pro/V4 Flash）
- ✅ 修复了段落分割问题（支持10-200字随意调节）
- ✅ 修复了 Word 导出功能
- ✅ 保存原始文本，支持反复重新分割

### v1.0
- 初始版本
