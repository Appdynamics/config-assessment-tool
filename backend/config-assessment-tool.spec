# -*- mode: python ; coding: utf-8 -*-

import os, sys
from os import path
import pptx
import streamlit
import shutil
import ctypes.util
import subprocess
from PyInstaller.utils.hooks import copy_metadata

pptx_path = path.dirname(pptx.__file__)
streamlit_path = path.dirname(streamlit.__file__)

def find_library_path(name):
    # Try using ctypes first
    path = ctypes.util.find_library(name)
    if path:
        # ctypes returns the filename (e.g., libcrypt.so.1), we need full path
        if os.path.exists(path):
            return path
        try:
             # Run ldconfig to find the path (Linux specific)
            res = subprocess.check_output(f"/sbin/ldconfig -p | grep {path}", shell=True)
            # Parse output: "libcrypt.so.1 (libc6,x86-64) => /lib/x86_64-linux-gnu/libcrypt.so.1"
            for line in res.decode().split('\n'):
                if path in line and "=>" in line:
                    return line.split("=>")[1].strip()
        except Exception:
            pass
    return None

block_cipher = None
bundle_name = "config-assessment-tool"
exec_file_name = bundle_name
platform = ""
platform_binaries=[]

if sys.platform == "win32":
    platform = "-windows"
    exec_file_name = f"{bundle_name}.exe"
elif sys.platform == "linux":
    platform = "-linux"
    # Find libcrypt dynamically instead of hardcoding
    libcrypt_path = find_library_path('crypt')
    if libcrypt_path and os.path.exists(libcrypt_path):
        print(f"Found libcrypt at: {libcrypt_path}")
        platform_binaries=[(libcrypt_path, '.')]
    else:
        print("Warning: libcrypt not found via ctypes/ldconfig. Relying on PyInstaller automatic collection or manual adding later if failed.")
        # Fallback to hardcoded if we want, or empty list. The previous hardcoded value was causing errors.
        # platform_binaries=[('/usr/local/lib/libcrypt.so.2','.')]
elif sys.platform == "darwin":
    platform = "-macosx"
else:
    print(f"Platform not clear. Creating generic bundle {sys.platform}")

version = open("VERSION", "r").read().strip()
bundle_name = f"{bundle_name}{platform}-{version}"

a = Analysis(
    ["../bin/bundle_main.py"],
    pathex=[os.path.abspath(".")],
    binaries=platform_binaries,
    datas=[
        ("../backend/resources/img/splash.txt", "backend/resources/img"),
        ("../VERSION", "."),
        ("../backend/resources/pptAssets/background.jpg", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/background_2.jpg", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/criteria.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/criteria2.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/checkmark.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/xmark.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/HybridApplicationMonitoringUseCase.json", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/HybridApplicationMonitoringUseCase_template.pptx", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/cxPpt_template.pptx", "backend/resources/pptAssets"),
     	(path.join(pptx_path,"templates"), "pptx/templates"), # for pptx
        (path.join(streamlit_path, "static"), "streamlit/static"), # for streamlit

        # Config files previously handled by manual copy
        ("../input/jobs/DefaultJob.json", "input/jobs"),
        ("../input/thresholds/DefaultThresholds.json", "input/thresholds"),
        ("../backend/resources/controllerDefaults/defaultHealthRulesAPM.json", "backend/resources/controllerDefaults"),
        ("../backend/resources/controllerDefaults/defaultHealthRulesBRUM.json", "backend/resources/controllerDefaults"),
        ("../plugins", "plugins"),
        ("../frontend", "frontend"),
    ] + copy_metadata('streamlit'),
    hiddenimports=['streamlit.runtime.scriptrunner.magic_funcs', 'backend.util', 'backend.util.logging_utils', 'tzlocal', 'streamlit_modal', 'backend.core', 'backend.core.Engine', 'backend.api', 'backend.api.Result', 'backend.api.appd', 'backend.api.appd.AppDService', 'backend.api.appd.AppDController', 'backend.api.appd.AuthMethod', 'backend.extractionSteps', 'backend.extractionSteps.general', 'backend.extractionSteps.maturityAssessment'],
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
    name=exec_file_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=True, upx_exclude=[], name=bundle_name, contents_directory='_internal')

# Post-processing: Move 'input' directory out of '_internal' to the bundle root
destination_dir = os.path.join(DISTPATH, bundle_name)
internal_dir = os.path.join(destination_dir, '_internal')
input_source = os.path.join(internal_dir, 'input')
input_dest = os.path.join(destination_dir, 'input')

if os.path.exists(input_source):
    if os.path.exists(input_dest):
        shutil.rmtree(input_dest)
    shutil.move(input_source, input_dest)
    print(f"Moved {input_source} to {input_dest}")

# Post-processing: Move 'plugins' directory out of '_internal' to the bundle root
plugins_source = os.path.join(internal_dir, 'plugins')
plugins_dest = os.path.join(destination_dir, 'plugins')

if os.path.exists(plugins_source):
    if os.path.exists(plugins_dest):
        shutil.rmtree(plugins_dest)
    shutil.move(plugins_source, plugins_dest)
    print(f"Moved {plugins_source} to {plugins_dest}")

# Post-processing: Ensure 'output' directory exists in the bundle root
output_dest = os.path.join(destination_dir, 'output')
if not os.path.exists(output_dest):
    os.makedirs(output_dest)
    print(f"Created {output_dest}")
