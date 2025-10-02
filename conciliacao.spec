# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['conciliacao.py'],
    pathex=[],
    binaries=[],
    datas=[('credentials.json', '.')],
    hiddenimports=['pyautogui', 'pygetwindow', 'pyscreeze', 'pytweening', 'mouseinfo', 'pywin32', 'pywinauto', 'cv2', 'holidays', 'anthropic', 'openai', 'pyexcel_xls', 'pyexcel_xlsx', 'pyexcel_io', 'pyexcel_io.readers', 'pyexcel_io.writers', 'pyexcel.plugins.parsers.excel', 'pyexcel.plugins.renderers.excel', 'pyexcel.plugins.sources.file_input', 'pyexcel.plugins.sources.file_output', 'google_auth_oauthlib', 'googleapiclient', 'googleapiclient.discovery', 'googleapiclient.http', 'google.auth', 'google.auth.transport.requests'],
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
    name='conciliacao',
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
