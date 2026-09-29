# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[os.path.abspath('.')],
    binaries=[],
    datas=[
        ('compare', 'compare'),     # 把 compare 資料夾完整帶入
    ],
    hiddenimports=[
        # 將你根目錄的所有核心 py 模組全部宣告進來
        'company_helper',
        'comparelist',
        'resultExport',
        'utils',
        # 將 compare 資料夾及其內部的比對模組宣告進來
        'compare',
        'compare.SAP625_SAN070R1',
        # 確保 GUI 拖放外掛正常運作
        'tkinterdnd2',
    ],
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
