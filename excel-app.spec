# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# 收集所有pandas和openpyxl的依赖
pandas_datas = collect_all('pandas')
openpyxl_datas = collect_all('openpyxl')
numpy_datas = collect_all('numpy')
tkinter_datas = collect_all('tkinter')

added_datas = []
added_datas.extend(pandas_datas[0])
added_datas.extend(openpyxl_datas[0])
added_datas.extend(numpy_datas[0])
added_datas.extend(tkinter_datas[0])

added_binaries = []
added_binaries.extend(pandas_datas[1])
added_binaries.extend(openpyxl_datas[1])
added_binaries.extend(numpy_datas[1])
added_binaries.extend(tkinter_datas[1])

added_hiddenimports = []
added_hiddenimports.extend(pandas_datas[2])
added_hiddenimports.extend(openpyxl_datas[2])
added_hiddenimports.extend(numpy_datas[2])
added_hiddenimports.extend(tkinter_datas[2])

# 添加更多可能需要的隐藏导入
added_hiddenimports.extend(collect_submodules('dateutil'))
added_hiddenimports.extend(collect_submodules('pandas._libs'))
added_hiddenimports.extend(collect_submodules('openpyxl.cell'))
added_hiddenimports.extend(collect_submodules('openpyxl.styles'))

# 添加本项目的Python文件
datas = [
    ('*.py', '.'),
    ('*.ico', '.'),
]
datas.extend(added_datas)

a = Analysis(
    ['excel_ui.py'],
    pathex=[],
    binaries=added_binaries,
    datas=datas,
    hiddenimports=added_hiddenimports + [
        'pandas', 
        'openpyxl', 
        'tkinter', 
        'openpyxl.styles.numbers',
        'openpyxl.styles.fonts',
        'openpyxl.styles.alignment',
        'openpyxl.styles.borders',
        'datetime',
        're',
        'numpy',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
        'tkinter.ttk',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 排除多余的DLL文件以减小体积（Windows特有）
if sys.platform == 'win32':
    excluded_dlls = [
        'MSVCP140.dll',  # 系统DLL，无需打包
        'VCRUNTIME140.dll',  # 系统DLL，无需打包
    ]
    a.binaries = [x for x in a.binaries if not any(e.lower() == x[0].lower() for e in excluded_dlls)]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel数据处理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 暂时不使用图标
    version='file_version_info.txt' if os.path.exists('file_version_info.txt') else None,
)

if sys.platform == 'darwin':  # macOS
    app = BUNDLE(
        exe,
        name='Excel数据处理工具.app',
        icon=None,
        bundle_identifier=None,
    ) 