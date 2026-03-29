# shuck-md2docx

Windows 右键菜单工具：一键将 Markdown 转换为格式规范的 Word 文档。

A Windows context-menu tool that converts Markdown files to well-formatted Word (.docx) documents with one click.

---

## 功能特性 / Features

- **右键一键转换 / One-Click Conversion**
  右键点击任意 `.md` 文件，选择「转换为 DOCX」即可。
  Right-click any `.md` file and select "Convert to DOCX".

- **Obsidian 图片支持 / Obsidian Image Support**
  自动识别 `![[image.png|size]]` 语法，在同目录及子目录中查找图片并嵌入 Word。
  Automatically recognizes `![[image.png|size]]` syntax, locates images in the directory tree, and embeds them.

- **引文行内化 / Inline Citations**
  将脚注（`^[1]`、`^{1-5}`、`[^id]`）转换为行内引用 `[1]`，避免 Word 中生成脚注。
  Converts footnotes (`^[1]`, `^{1-5}`, `[^id]`) to inline references `[1]` instead of Word footnotes.

- **三线表 / Three-Line Tables**
  表格自动转为学术论文常用的三线表样式。
  Tables are automatically formatted in the three-line style commonly used in academic papers.

- **学术排版 / Academic Formatting**
  - 中文宋体 + 英文 Times New Roman
  - 字号小四（12pt）
  - 两倍行距、两端对齐
  - 自动移除水平分隔线

## 前置要求 / Prerequisites

- **Windows 10/11**
- **Python 3.8+** — [下载 / Download](https://www.python.org/downloads/)
- **Pandoc** — [下载 / Download](https://pandoc.org/installing.html)

## 安装 / Installation

### 方法一：一键安装（推荐）/ Method 1: One-Click Install (Recommended)

1. 下载或克隆本仓库 / Download or clone this repo:
   ```bash
   git clone https://github.com/Shan-Zhu/shuck-md2docx.git
   ```

2. 双击运行 `install.bat`（会自动请求管理员权限）。
   Double-click `install.bat` (it will request admin privileges automatically).

   该脚本会自动完成：
   The script will automatically:
   - 检查 Python 和 Pandoc 是否已安装 / Check Python and Pandoc installation
   - 安装 python-docx 依赖 / Install python-docx dependency
   - 注册右键菜单 / Register the context menu entry

### 方法二：手动安装 / Method 2: Manual Install

1. 安装依赖 / Install dependency:
   ```bash
   pip install python-docx
   ```

2. 生成注册表文件 / Generate registry files:
   ```bash
   python setup.py
   ```

3. 双击 `install.reg` 导入注册表（需管理员权限）。
   Double-click `install.reg` to import (requires admin privileges).

## 使用 / Usage

1. 在文件资源管理器中，右键点击任意 `.md` 文件。
   In File Explorer, right-click any `.md` file.

2. 选择 **「转换为 DOCX」**。
   Select **"Convert to DOCX"**.

3. 转换完成后会弹窗提示，生成的 `.docx` 文件与原文件在同一目录。
   A dialog will confirm success. The `.docx` file is saved alongside the original.

## 卸载 / Uninstall

双击运行 `uninstall.bat`，或手动双击 `uninstall.reg`。

Double-click `uninstall.bat`, or manually double-click `uninstall.reg`.

## 项目结构 / Project Structure

```
shuck-md2docx/
├── md2docx.py      # 核心转换脚本 / Core conversion script
├── setup.py        # 注册表文件生成器 / Registry file generator
├── install.bat     # 一键安装 / One-click installer
├── uninstall.bat   # 一键卸载 / One-click uninstaller
├── LICENSE
└── README.md
```

## 许可证 / License

[MIT License](LICENSE)
