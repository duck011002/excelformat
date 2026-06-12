# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata


import os
import sys

# Prepend active environment's Library/bin to PATH to ensure correct DLLs are collected
conda_bin = os.path.join(sys.prefix, "Library", "bin")
if os.path.exists(conda_bin):
    os.environ["PATH"] = conda_bin + os.path.pathsep + os.environ["PATH"]

block_cipher = None

datas = [("app.py", ".")]
binaries = []
if os.path.exists(conda_bin):
    for dll in ["libssl-3-x64.dll", "libcrypto-3-x64.dll"]:
        dll_path = os.path.join(conda_bin, dll)
        if os.path.exists(dll_path):
            binaries.append((dll_path, "."))
hiddenimports = list(collect_submodules("grade_excel_cleaner"))
hiddenimports.extend(collect_submodules("streamlit.runtime.scriptrunner"))
hiddenimports.append("streamlit.runtime.scriptrunner.magic_funcs")

for package in [
    "streamlit",
    "altair",
    "pydeck",
    "pandas",
    "openai",
    "pydantic",
    "openpyxl",
    "pyxlsb",
    "xlrd",
    "python_calamine",
]:
    try:
        datas.extend(copy_metadata(package))
    except Exception:
        pass

datas.extend(collect_data_files("streamlit"))

a = Analysis(
    ["windows_launcher.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pandas.tests",
        "numpy.tests",
        "pyarrow.tests",
        "streamlit.testing",
        "IPython",
        "ipykernel",
        "ipywidgets",
        "matplotlib",
        "scipy",
        "pytest",
    ],
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
    name="GradeExcelCleaner",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
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
    upx=True,
    upx_exclude=[],
    name="GradeExcelCleaner",
)
