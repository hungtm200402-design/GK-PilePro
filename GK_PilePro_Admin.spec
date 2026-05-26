# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path


datas = [
    ('.env', '.'),
    ('assets\\tool_kl_logo.png', 'assets'),
    ('assets\\tool_kl_taskbar.png', 'assets'),
    ('assets\\tool_kl.ico', 'assets'),
]

tcl_root = Path(sys.base_prefix) / 'tcl'


def add_tree(datas, src_root, target_root):
    if not src_root.exists():
        return
    for dirpath, _dirnames, filenames in os.walk(src_root):
        dirpath = Path(dirpath)
        rel_dir = dirpath.relative_to(src_root)
        for filename in filenames:
            src_file = dirpath / filename
            target_dir = Path(target_root) / rel_dir
            datas.append((str(src_file), str(target_dir)))


add_tree(datas, tcl_root / 'tcl8.6', '_tcl_data')
add_tree(datas, tcl_root / 'tcl8.6', 'tcl_data')
add_tree(datas, tcl_root / 'tk8.6', '_tk_data')
add_tree(datas, tcl_root / 'tk8.6', 'tk_data')


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
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
    contents_directory='.',
)
