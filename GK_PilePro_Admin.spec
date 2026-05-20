# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('.env', '.'),
        ('assets\\tool_kl_logo.png', 'assets'),
        ('assets\\tool_kl_taskbar.png', 'assets'),
        ('assets\\tool_kl.ico', 'assets'),
        ('C:\\Users\\ADMIN\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tcl8.6', '_tcl_data'),
        ('C:\\Users\\ADMIN\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tk8.6', '_tk_data'),
    ],
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
    a.binaries,
    a.datas,
    [],
    name='GK PilePro Admin',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='assets\\tool_kl.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
