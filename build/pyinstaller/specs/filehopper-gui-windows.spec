# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
hiddenimports += collect_submodules('pandastable')


a = Analysis(
    ['..\\..\\..\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\jeroe\\Documents\\Visual Studio\\PythonHopper\\clients_db.json', '.'), ('C:\\Users\\jeroe\\Documents\\Visual Studio\\PythonHopper\\suppliers_db.json', '.'), ('C:\\Users\\jeroe\\Documents\\Visual Studio\\PythonHopper\\delivery_addresses_db.json', '.'), ('C:\\Users\\jeroe\\Documents\\Visual Studio\\PythonHopper\\build\\pyinstaller\\data-files\\app_settings.json', '.'), ('C:\\Users\\jeroe\\Documents\\Visual Studio\\PythonHopper\\suppliers_template.csv', '.')],
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
    a.binaries,
    a.datas,
    [],
    name='filehopper-gui-windows',
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
    version='C:\\Users\\jeroe\\Documents\\Visual Studio\\PythonHopper\\build\\pyinstaller\\version-files\\filehopper-gui-windows.version.txt',
)
