# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = [
    "PyQt5.sip",
    "PyQt5.QtCore",
    "PyQt5.QtGui",
    "PyQt5.QtWidgets",
    "PyQt5.QtNetwork",
    "win32timezone",
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    "apscheduler.triggers.cron",
    "apscheduler.triggers.date",
    "apscheduler.schedulers.background",
    "apscheduler.jobstores.memory",
    "apscheduler.executors.pool",
    "tzlocal",
    "oletools.olevba",
    "olefile",
]
hiddenimports += collect_submodules("apscheduler")

a = Analysis(
    ["scheduler_app.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pandas", "numpy", "matplotlib", "tkinter"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="CMC스케줄러",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="CMC스케줄러",
)
