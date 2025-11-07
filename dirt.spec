# DIRT.spec
# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules
import os

block_cipher = None

current_dir = os.path.dirname(__file__) if '__file__' in globals() else os.getcwd()
e_png = os.path.join(current_dir, 'E.png')
bosch_png = os.path.join(current_dir, 'Bosch.png')

datas = []
for path in [e_png, bosch_png]:
    if os.path.exists(path):
        datas.append((path, '.'))

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=collect_submodules('xlsxwriter') + ['tkinter', '_tkinter', 'tcl', 'tk'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DIRT',
    debug=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='DIRT'
)