# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all


datas = [
    ('.env', '.'),
    ('assets\\gk_logo.png', 'assets'),
    ('assets\\tool_kl_logo.png', 'assets'),
    ('assets\\gk_app_icon.png', 'assets'),
    ('assets\\gk_icon_emblem.png', 'assets'),
    ('assets\\gk_taskbar_icon.png', 'assets'),
    ('assets\\gk_app_icon.ico', 'assets'),
    ('assets\\gk_splash_logo.png', 'assets'),
    ('assets\\gk_splash_bg.png', 'assets'),
    ('assets\\loading\\loading_bg.png', 'assets\\loading'),
    ('assets\\loading\\video-loading.mp4', 'assets\\loading'),
    ('assets\\gk_ui_decor_bottom_right.png', 'assets'),
    ('assets\\gk_ui_decor_sidebar_bottom.png', 'assets'),
    ('assets\\gk_footer_decoration.png', 'assets'),
    ('assets\\gk_sidebar_decoration.png', 'assets'),
    ('assets\\decor_part1_solid.png', 'assets'),
    ('assets\\decor_part2_solid.png', 'assets'),
    ('assets\\decor_part3_solid.png', 'assets'),
    ('assets\\goc_phai_moi.png.png', 'assets'),
    ('assets\\goc_trai_moi.png.png', 'assets'),
]
binaries = []
hiddenimports = []

pil_datas, pil_binaries, pil_hiddenimports = collect_all('PIL')
datas += pil_datas
binaries += pil_binaries
hiddenimports += pil_hiddenimports

cv2_datas, cv2_binaries, cv2_hiddenimports = collect_all('cv2')
datas += cv2_datas
binaries += cv2_binaries
hiddenimports += cv2_hiddenimports

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
add_tree(datas, Path('assets') / 'GK_PilePro_icon_files_no_bg', Path('assets') / 'GK_PilePro_icon_files_no_bg')
add_tree(
    datas,
    Path('assets') / 'sidebar_icon_white_outline_transparent' / 'canvas_128',
    Path('assets') / 'sidebar_icon_white_outline_transparent' / 'canvas_128',
)

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
    name='GK PilePro Admin',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='assets\\gk_app_icon.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    contents_directory='.',
)
