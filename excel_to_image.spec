# -*- mode: python ; coding: utf-8 -*-

mydatas = [
('./aspose-cells-24.4.jar', '.'),
('./bcprov-jdk15on-1.68.jar', '.'),
('./bcpkix-jdk15on-1.68.jar', '.'),
('./JavaClassBridge.jar', '.') 
]

a = Analysis(
    ['excel_to_image.py'],
    pathex=[],
    binaries=[],
    datas=mydatas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    name='excel_to_image',
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
)
