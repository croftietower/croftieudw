# Bidirectional Translator 🌐

一个基于 Python/Tkinter 的双向翻译桌面应用，支持中英互译，并可导出为 Word 文档。

## ✨ 功能特性

- 中英文双向翻译
- 段落分割（支持 10-200 字自由调节）
- Word 文档导出
- 模型切换（DeepSeek V4 Pro / V4 Flash）
- 保存原始文本，支持反复重新分割
- 跨平台支持（macOS / Windows）

## 📦 安装与运行

### 安装依赖

```bash
pip install requests python-docx
```

### 运行

```bash
python translation_app.py
```

### 打包为可执行文件

详见 [README_PACKAGE.md](README_PACKAGE.md)

## 🔄 自动构建

项目配置了 GitHub Actions，推送 `v*` 格式的 tag 时会自动构建 macOS 和 Windows 版本：

```bash
git tag v1.0.0
git push origin v1.0.0
```

## 📁 项目结构

```
├── translation_app.py          # 主程序源码
├── Bidirectional Translator v1.0.0.spec  # PyInstaller 打包配置
├── package_mac.sh              # macOS 一键打包脚本
├── package_windows.bat         # Windows 一键打包脚本
├── README_PACKAGE.md           # 打包说明文档
├── .github/workflows/build.yml # GitHub Actions CI/CD
└── .gitignore
```

## 📝 更新日志

### v1.1 (最新)
- ✅ 添加了模型切换功能（DeepSeek V4 Pro/V4 Flash）
- ✅ 修复了段落分割问题（支持10-200字随意调节）
- ✅ 修复了 Word 导出功能
- ✅ 保存原始文本，支持反复重新分割

### v1.0
- 初始版本
