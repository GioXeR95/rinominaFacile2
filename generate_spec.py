#!/usr/bin/env python
"""
Generate PyInstaller spec file that works on Windows, Linux, and macOS
"""
import os
from pathlib import Path

spec_content = '''# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import sys
import os

datas = [
    ('app/translations', 'translations'),
]
binaries = []
hiddenimports = ['ui', 'ui.components', 'ui.toolbar', 'core', 'ai']

tmp_ret = collect_all('ui')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

tmp_ret = collect_all('core')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

tmp_ret = collect_all('ai')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

tmp_ret = collect_all('translations')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

a = Analysis(
    ['app/main.py'],
    pathex=['app'],
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
    a.zipfiles,
    a.datas,
    [],
    name='RinominaFacile2',
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

# macOS app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='RinominaFacile2.app',
        icon=None,
        bundle_identifier='com.rinominafacile.app',
        info_plist={
            'NSPrincipalClass': 'NSApplication',
            'NSHighResolutionCapable': 'True',
        },
    )
'''

# Genera il file spec
spec_path = Path('RinominaFacile2.spec')
spec_path.write_text(spec_content)
print(f"✓ Spec file generato: {spec_path}")
print("\nOra esegui:")
print("  pyinstaller RinominaFacile2.spec")
