# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all


datas = [
    ('assets\\tool_kl_logo.png', 'assets'),
    ('assets\\tool_kl_taskbar.png', 'assets'),
    ('assets\\tool_kl.ico', 'assets'),
]
binaries = []
hiddenimports = []

pil_datas, pil_binaries, pil_hiddenimports = collect_all('PIL')
datas += pil_datas
binaries += pil_binaries
hiddenimports += pil_hiddenimports

env_file = Path('.env')
if env_file.exists():
    datas.append((str(env_file), '.'))

tcl_root = Path(sys.base_prefix) / 'tcl'
for folder, target in (('tcl8.6', '_tcl_data'), ('tk8.6', '_tk_data')):
    src = tcl_root / folder
    if src.exists():
        datas.append((str(src), target))

python_dlls = Path(sys.base_prefix) / 'DLLs'
for dll_name in ('_tkinter.pyd', 'tcl86t.dll', 'tk86t.dll'):
    dll_path = python_dlls / dll_name
    if dll_path.exists():
        binaries.append((str(dll_path), '.'))

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='Tool_KL_UI8',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
