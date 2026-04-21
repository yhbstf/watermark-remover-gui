# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec used by the GitHub Actions release workflow."""

import sys

from PyInstaller.utils.hooks import collect_all, collect_submodules


datas = []
binaries = []
hiddenimports = []

# iopaint + torch have heavy dynamic loading; collect everything.
for pkg in ("iopaint", "torch"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# A few modules iopaint pulls in that PyInstaller sometimes misses.
hiddenimports += [
    "iopaint",
    "iopaint.cli",
    "iopaint.api",
    "iopaint.model.lama",
    "iopaint.single_processing",
    "iopaint.batch_processing",
    "safetensors",
    "huggingface_hub",
]
hiddenimports += collect_submodules("iopaint")

block_cipher = None

a = Analysis(
    ["watermark_remover_gui.py"],
    pathex=[],
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

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="watermark-remover-gui",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="watermark-remover-gui",
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="watermark-remover-gui.app",
        icon=None,
        bundle_identifier="com.yhbstf.watermark-remover-gui",
        info_plist={
            "NSHighResolutionCapable": "True",
            "LSBackgroundOnly": "False",
        },
    )
