#!/usr/bin/env python3
"""
UMDoc 启动引导器 - 自动管理虚拟环境和所有依赖
只需：python3 launcher.py
"""
import subprocess
import sys
from pathlib import Path

VENV_DIR = "app_env"
REQUIRED_PACKAGES = ["PySide6", "markitdown[all]"]

def get_venv_python():
    base = Path(__file__).resolve().parent
    venv_path = base / VENV_DIR
    if sys.platform == "win32":
        return venv_path / "Scripts" / "python.exe"
    else:
        return venv_path / "bin" / "python"

def is_venv_valid():
    return get_venv_python().is_file()

def create_venv():
    venv_path = Path(__file__).resolve().parent / VENV_DIR
    print(f"创建虚拟环境: {venv_path}")
    subprocess.run([sys.executable, "-m", "venv", str(venv_path)], check=True)

def install_packages():
    python_exe = str(get_venv_python())
    print("升级 pip ...")
    subprocess.run([python_exe, "-m", "pip", "install", "--upgrade", "pip"],
                   capture_output=True)
    for pkg in REQUIRED_PACKAGES:
        print(f"安装 {pkg} ...")
        subprocess.run([python_exe, "-m", "pip", "install", pkg], check=True)

def check_all_dependencies():
    python_exe = str(get_venv_python())
    try:
        # 检查关键库能否导入
        code = "import PySide6, markitdown, openpyxl"
        subprocess.run([python_exe, "-c", code],
                       capture_output=True, check=True)
        return True
    except subprocess.CalledProcessError:
        return False

def main():
    if not is_venv_valid():
        create_venv()

    if not check_all_dependencies():
        print("正在安装所需依赖（首次可能需要几分钟，包含全量文档支持）...")
        install_packages()
        print("安装完成！")
    else:
        print("依赖已就绪，启动应用...")

    main_script = Path(__file__).resolve().parent / "umdoc.py"
    python_exe = str(get_venv_python())
    print(f"启动 {main_script}")
    subprocess.run([python_exe, str(main_script)])

if __name__ == "__main__":
    main()