# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['fit_images_column_k.py'],
    pathex=[],
    binaries=[],
    datas=[('YuRa - Copy.ico', '.'), ('templates', 'templates'), ('static', 'static')],
    hiddenimports=[],
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
    name='FitImagesColumnK',
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
    icon=['YuRa - Copy.ico'],
)
