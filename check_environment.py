#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
环境检查脚本
用于检查打包前的环境是否满足要求
"""

import os
import sys
import platform
import importlib
import subprocess
import traceback

def print_header(title):
    """打印美观的标题"""
    print("\n" + "=" * 60)
    print(f" {title} ".center(60, "="))
    print("=" * 60)

def print_status(name, status, details=None):
    """打印状态信息"""
    status_str = "✓ 已安装" if status else "✗ 未安装"
    print(f"{name:<20} {status_str:<10}", end="")
    if details:
        print(f" - {details}")
    else:
        print()

def check_module(module_name, import_name=None):
    """检查Python模块是否已安装"""
    if import_name is None:
        import_name = module_name
    
    try:
        module = importlib.import_module(import_name)
        version = getattr(module, "__version__", "未知版本")
        return True, version
    except ImportError:
        return False, None

def check_command(command):
    """检查系统命令是否可用"""
    try:
        result = subprocess.run(
            command, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE,
            shell=True,
            text=True
        )
        return result.returncode == 0, result.stdout.strip()
    except:
        return False, None

def main():
    """主函数"""
    print_header("Excel数据处理工具 - 环境检查")
    
    # 检查Python版本
    py_version = platform.python_version()
    py_implementation = platform.python_implementation()
    py_status = True if tuple(map(int, py_version.split('.'))) >= (3, 6) else False
    
    print(f"操作系统: {platform.system()} {platform.release()}")
    print(f"Python版本: {py_version} ({py_implementation})")
    print(f"Python路径: {sys.executable}")
    print("\n")
    
    if not py_status:
        print("警告: Python版本低于3.6，可能会影响程序运行!")
    
    # 检查必要模块
    required_modules = [
        ("pandas", None),
        ("openpyxl", None),
        ("tkinter", "tkinter"),
        ("pyinstaller", None)
    ]
    
    print_header("必要模块检查")
    all_modules_installed = True
    
    for module_name, import_name in required_modules:
        status, version = check_module(module_name, import_name)
        print_status(module_name, status, version)
        if not status:
            all_modules_installed = False
    
    # 检查PyInstaller
    print_header("打包工具检查")
    pyinstaller_cmd_status, pyinstaller_version = check_command("pyinstaller --version")
    print_status("PyInstaller命令", pyinstaller_cmd_status, pyinstaller_version)
    
    # 检查项目文件
    print_header("项目文件检查")
    required_files = [
        "excel_ui.py",
        "excel_processor.py",
        "excel_icon.ico",
        "requirements.txt",
        "excel-app.spec",
    ]
    
    all_files_exist = True
    for file_name in required_files:
        file_exists = os.path.exists(file_name)
        print_status(file_name, file_exists)
        if not file_exists:
            all_files_exist = False
    
    # 总结
    print_header("检查结果")
    
    if all_modules_installed and pyinstaller_cmd_status and all_files_exist:
        print("✓ 所有检查通过! 环境准备就绪，可以开始打包。")
    else:
        print("✗ 检查未全部通过，请解决以上问题后再尝试打包。")
        
        if not all_modules_installed:
            print("\n提示: 请运行以下命令安装缺失的模块:")
            print("pip install -r requirements.txt")
        
        if not pyinstaller_cmd_status:
            print("\n提示: PyInstaller未安装或不可用，请运行:")
            print("pip install pyinstaller")
        
        if not all_files_exist:
            print("\n提示: 有必要的项目文件缺失，请确保所有文件都在当前目录。")
    
    return 0 if (all_modules_installed and pyinstaller_cmd_status and all_files_exist) else 1

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print("\n" + "!" * 60)
        print("检查过程中发生错误:")
        print(str(e))
        print("\n详细错误信息:")
        traceback.print_exc()
        print("!" * 60)
        sys.exit(1) 