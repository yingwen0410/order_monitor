# order-monitor.spec
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# tkinterdnd2 needs manual datas entry so the tcl/tk DnD extension is bundled
import tkinterdnd2 as _tkdnd
_tkdnd_path = os.path.dirname(_tkdnd.__file__)

a = Analysis(
    ['system/main.py'],
    pathex=['system'],          # lets 'import utils', 'import reader', etc. resolve
    binaries=[],
    datas=[
        (_tkdnd_path, 'tkinterdnd2'),       # DnD extension files
        *collect_data_files('openpyxl'),    # openpyxl built-in templates
    ],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32com.shell',
        'pywintypes',
        'win32api',
        'win32con',
        'win32gui',
        *collect_submodules('win32com'),
    ],
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
    [],
    exclude_binaries=True,
    name='訂單未交量報表産生器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,              # no console window for warehouse staff
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='訂單未交量報表産生器',
)
