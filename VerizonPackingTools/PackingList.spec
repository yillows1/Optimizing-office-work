# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = [('Excel Bases', 'Excel Bases')]
datas += collect_data_files('win32com')


a = Analysis(
    ['packing_list_gui.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=['win32com', 'win32com.client', 'win32timezone', 'pythoncom', 'pywintypes', 'improvedScanner', 'pypdf', 'docx', 'openpyxl'],
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
    [],
    exclude_binaries=True,
    name='PackingList',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PackingList',
)
