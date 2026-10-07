# -*- mode: python ; coding: utf-8 -*-

import os
import glob


comparison_modules = [
    f"excel_comparator.comparisons.{os.path.basename(path)[:-3]}"
    for path in glob.glob(
        os.path.join("src", "excel_comparator", "comparisons", "*.py")
    )
    if not os.path.basename(path).startswith("__")
]
comparison_datas = [
    (path, os.path.join("excel_comparator", "comparisons"))
    for path in glob.glob(
        os.path.join("src", "excel_comparator", "comparisons", "*.py")
    )
]


a = Analysis(
    ['src/excel_comparator/main.py'],
    pathex=[os.path.abspath('src')],
    binaries=[],
    datas=[
        ('config', 'config'),
    ] + comparison_datas,
    hiddenimports=[
        # 將你根目錄的所有核心 py 模組全部宣告進來
        'excel_comparator',
        'excel_comparator.application.comparison_registry',
        # 確保 GUI 拖放外掛正常運作
        'tkinterdnd2',
    ] + comparison_modules,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ExcelComparator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
