# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('materials\\icon.ico', '.'), ('materials\\template.docx', '.'), ('materials\\nesterov_sign.png', '.'), ('materials\\gnelitskiy_sign.png', '.'), ('materials\\goncharok_sign.png', '.'), ('materials\\fonts\\timesnrcyrmt.ttf', '.'), ('materials\\Инструкция.pdf', '.')]
binaries = []
hiddenimports = []
tmp_ret = collect_all('transliterate')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
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
    name='ИИшница v0.2',
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
    icon=['materials\\icon.ico'],
)
