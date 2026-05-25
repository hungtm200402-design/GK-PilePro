# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('assets\\tool_kl_logo.png', 'assets'), ('assets\\tool_kl_taskbar.png', 'assets'), ('assets\\tool_kl.ico', 'assets'), ('C:\\Users\\tranh\\AppData\\Local\\Programs\\Python\\Python311\\tcl\\tcl8.6', '_tcl_data'), ('C:\\Users\\tranh\\AppData\\Local\\Programs\\Python\\Python311\\tcl\\tk8.6', '_tk_data')],
    hiddenimports=[],
    hookspath=['pyi_hooks'],
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
    name='Tool_KL_UI8_app',
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
    icon=['assets\\tool_kl.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Tool_KL_UI8_app',
)
