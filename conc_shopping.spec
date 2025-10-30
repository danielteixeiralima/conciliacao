# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['conc_shopping.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pyautogui', 'pywinauto', 'cv2', 'PIL', 'numpy', 'openai', 'anthropic', 'holidays', 'pyexcel_xls'],
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
    name='conc_shopping',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
