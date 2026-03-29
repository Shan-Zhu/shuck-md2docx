# shuck-md2docx

A Windows context-menu tool that converts Markdown files to well-formatted Word (.docx) documents with one click.

[中文](README.md) | **English**

---

## Features

- **One-Click Conversion** — Right-click any `.md` file and select "Convert to DOCX"
- **Obsidian Image Support** — Automatically recognizes `![[image.png|size]]` syntax, locates images in the directory tree, and embeds them
- **Inline Citations** — Converts footnotes (`^[1]`, `^{1-5}`, `[^id]`) to inline references `[1]` instead of Word footnotes
- **Three-Line Tables** — Tables are automatically formatted in the three-line style commonly used in academic papers
- **Academic Formatting**
  - Chinese: SimSun (宋体) / English: Times New Roman
  - Font size: 12pt
  - Double line spacing, justified alignment
  - Automatic removal of horizontal rules

## Prerequisites

- **Windows 10/11**
- **Python 3.8+** — [Download](https://www.python.org/downloads/)
- **Pandoc** — [Download](https://pandoc.org/installing.html)

## Installation

### Method 1: One-Click Install (Recommended)

1. Download the latest release from [Releases](https://github.com/Shan-Zhu/shuck-md2docx/releases) and extract, or clone the repo:
   ```bash
   git clone https://github.com/Shan-Zhu/shuck-md2docx.git
   ```

2. Double-click `install.bat` (it will request admin privileges automatically).

   The script will automatically:
   - Check that Python and Pandoc are installed
   - Install the python-docx dependency
   - Register the context menu entry

### Method 2: Manual Install

1. Install dependency:
   ```bash
   pip install python-docx
   ```

2. Generate registry files:
   ```bash
   python setup.py
   ```

3. Double-click `install.reg` to import (requires admin privileges).

## Usage

1. In File Explorer, right-click any `.md` file
2. Select **"Convert to DOCX"**
3. A dialog will confirm success. The `.docx` file is saved alongside the original

## Uninstall

Double-click `uninstall.bat`, or manually double-click `uninstall.reg`.

## Project Structure

```
shuck-md2docx/
├── md2docx.py      # Core conversion script
├── setup.py        # Registry file generator
├── install.bat     # One-click installer
├── uninstall.bat   # One-click uninstaller
├── LICENSE
├── README.md       # 中文说明
└── README_EN.md    # English README
```

## License

[MIT License](LICENSE)
