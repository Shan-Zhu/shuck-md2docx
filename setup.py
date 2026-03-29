"""生成 install.reg 和 uninstall.reg，用于添加/移除右键菜单"""

import os
import sys
import shutil

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MD2DOCX_SCRIPT = os.path.join(SCRIPT_DIR, "md2docx.py")

REG_KEY = r"SystemFileAssociations\.md\shell\md2docx"

INSTALL_TEMPLATE = r"""Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\{key}]
@="转换为 DOCX"
"Icon"="imageres.dll,2"

[HKEY_CLASSES_ROOT\{key}\command]
@="\"{python_exe}\" \"{script}\" \"%1\""
"""

UNINSTALL_TEMPLATE = r"""Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\{key}]
"""


def find_python():
    python_exe = sys.executable
    if not python_exe or "python" not in python_exe.lower():
        python_exe = shutil.which("python") or shutil.which("python3")
    if not python_exe:
        print("错误：找不到 Python 可执行文件")
        sys.exit(1)
    return python_exe


def generate():
    python_exe = find_python()

    # 注册表文件中反斜杠需要转义
    python_escaped = python_exe.replace("\\", "\\\\")
    script_escaped = MD2DOCX_SCRIPT.replace("\\", "\\\\")

    install_content = INSTALL_TEMPLATE.format(
        key=REG_KEY, python_exe=python_escaped, script=script_escaped
    )
    uninstall_content = UNINSTALL_TEMPLATE.format(key=REG_KEY)

    install_path = os.path.join(SCRIPT_DIR, "install.reg")
    uninstall_path = os.path.join(SCRIPT_DIR, "uninstall.reg")

    # .reg 文件需要 UTF-16 LE BOM 编码
    for path, content in [(install_path, install_content), (uninstall_path, uninstall_content)]:
        with open(path, "w", encoding="utf-16-le") as f:
            f.write("\ufeff" + content)

    print(f"已生成：")
    print(f"  {install_path}")
    print(f"  {uninstall_path}")
    print()
    print("双击 install.reg 即可添加右键菜单")
    print("双击 uninstall.reg 可移除右键菜单")


if __name__ == "__main__":
    generate()
