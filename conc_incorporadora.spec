# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['conc_incorporadora.py'],
    pathex=[],
    binaries=[],
    datas=[('credentials.json', '.')],
    hiddenimports=['pyexcel_xls', 'pyexcel_xlsx', 'pyexcel_io', 'pyexcel_io.readers', 'pyexcel_io.writers', 'pyexcel.plugins.parsers.excel', 'pyexcel.plugins.renderers.excel', 'pyexcel.plugins.sources.file_input', 'pyexcel.plugins.sources.file_output', 'google_auth_oauthlib', 'googleapiclient', 'googleapiclient.discovery', 'googleapiclient.http', 'google.auth', 'google.auth.transport.requests'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='conc_incorporadora',
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
