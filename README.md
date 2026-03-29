# shuck-md2docx

Windows 右键菜单工具：一键将 Markdown 转换为格式规范的 Word 文档。

**[English](README_EN.md)** | 中文

---

## 功能特性

- **右键一键转换** — 右键点击任意 `.md` 文件，选择「转换为 DOCX」即可
- **Obsidian 图片支持** — 自动识别 `![[image.png|size]]` 语法，在同目录及子目录中查找图片并嵌入 Word
- **引文行内化** — 将脚注（`^[1]`、`^{1-5}`、`[^id]`）转换为行内引用 `[1]`，避免 Word 中生成脚注
- **三线表** — 表格自动转为学术论文常用的三线表样式
- **学术排版**
  - 中文宋体 + 英文 Times New Roman
  - 字号小四（12pt）
  - 两倍行距、两端对齐
  - 自动移除水平分隔线

## 前置要求

- **Windows 10/11**
- **Python 3.8+** — [下载](https://www.python.org/downloads/)
- **Pandoc** — [下载](https://pandoc.org/installing.html)

## 安装

### 方法一：一键安装（推荐）

1. 从 [Releases](https://github.com/Shan-Zhu/shuck-md2docx/releases) 下载最新版本压缩包并解压，或克隆本仓库：
   ```bash
   git clone https://github.com/Shan-Zhu/shuck-md2docx.git
   ```

2. 双击运行 `install.bat`（会自动请求管理员权��）。

   该脚本会自动完成：
   - 检查 Python 和 Pandoc 是否已安装
   - 安装 python-docx 依赖
   - 注册右键菜单

### 方法二：手动安装

1. 安装依赖：
   ```bash
   pip install python-docx
   ```

2. 生成注册表文件：
   ```bash
   python setup.py
   ```

3. 双击 `install.reg` ���入注册表（需管理员权限）。

## 使用

1. 在文件资源管理器中，右键点击任意 `.md` 文件
2. 选择 **「转换为 DOCX」**
3. 转换完成后会弹窗提示，生成的 `.docx` 文件与原文件在同一目录

## 卸载

双击运行 `uninstall.bat`，或手动双击 `uninstall.reg`。

## 项目结构

```
shuck-md2docx/
├── md2docx.py      # 核心转换脚本
├── setup.py        # 注册表文件生成器
├── install.bat     # 一键安装
├── uninstall.bat   # 一键卸载
├── LICENSE
├── README.md       # 中文说明
└── README_EN.md    # English README
```

## 许可证

[MIT License](LICENSE)
