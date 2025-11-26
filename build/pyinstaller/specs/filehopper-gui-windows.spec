# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
hiddenimports += collect_submodules('pandastable')


a = Analysis(
    ['..\\..\\..\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\jeroe\\Documents\\GitHub\\PythonHopper\\PythonHopper\\clients_db.json', '.'), ('C:\\Users\\jeroe\\Documents\\GitHub\\PythonHopper\\PythonHopper\\suppliers_db.json', '.'), ('C:\\Users\\jeroe\\Documents\\GitHub\\PythonHopper\\PythonHopper\\delivery_addresses_db.json', '.'), ('C:\\Users\\jeroe\\Documents\\GitHub\\PythonHopper\\PythonHopper\\app_settings.json', '.')],
    hiddenimports=hiddenimports,
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
    name='filehopper-gui-windows',
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
    name='filehopper-gui-windows',
)
