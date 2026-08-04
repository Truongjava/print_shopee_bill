# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for TTS_Bill — Shopee Seller Automation"""

import sys
from pathlib import Path

_basedir = Path('.')
_bill_dir = _basedir / 'bill_calculate'

# ── Collect bill_calculate module files ──
bill_datas = []
for f in _bill_dir.glob('*.py'):
    bill_datas.append((str(f), 'bill_calculate'))

# ── Collect Excel data files (bundled defaults) ──
excel_datas = []
for xlsx_name in ['mã combo.xlsx', 'sp bán lẻ.xlsx', 'Bảng thống kê hàng.xlsx']:
    p = _basedir / xlsx_name
    if p.exists():
        excel_datas.append((str(p), '.'))

# ── Collect cookie (if exists) ──
cookie_datas = []
for c in _basedir.glob('banhang.shopee.vn_*.json'):
    cookie_datas.append((str(c), '.'))

# ── Bundle SumatraPDF portable (PDF printing engine) ──
sumatra = _basedir / 'SumatraPDF.exe'
sumatra_datas = [(str(sumatra), '.')] if sumatra.exists() else []

datas = bill_datas + excel_datas + cookie_datas + sumatra_datas

a = Analysis(
    ['main_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'openpyxl', 'pdfplumber', 'pypdf', 'playwright', 'fpdf',
        'pythoncom', 'win32com', 'win32com.client', 'win32print',
        'pywintypes', 'win32api',
        'PIL', 'PIL._imagingtk', 'PIL._tkinter_finder',
        'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets',
        'PySide6.QtNetwork', 'PySide6.QtPrintSupport',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', 'tcl', 'tk',
        'matplotlib', 'numpy', 'scipy', 'pandas',
        'notebook', 'jupyter', 'ipykernel',
        'setuptools', 'pip', 'wheel',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TTS_Bill',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # ← Windowed mode, không hiện console
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
