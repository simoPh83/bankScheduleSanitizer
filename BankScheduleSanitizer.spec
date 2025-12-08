# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['bank_schedule_sanitizer.py'],
    pathex=[],
    binaries=[],
    datas=[('instructions', 'instructions')],
    hiddenimports=['pandas', 'openpyxl', 'xlsxwriter', 'tkinter', 'datetime'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', '_pytest', 'test'],
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
    name='BankScheduleSanitizer',
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
app = BUNDLE(
    exe,
    name='BankScheduleSanitizer.app',
    icon=None,
    bundle_identifier='com.simoneph.bankschedulesanitizer',
)
