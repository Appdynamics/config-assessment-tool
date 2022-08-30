# -*- mode: python ; coding: utf-8 -*-

import os, sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p) # for pptx

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
    platform_binaries=[('/usr/local/lib/libcrypt.so.2','.')]
elif sys.platform == "darwin":
    platform = "-macosx"
else:
    print(f"Platform not clear. Creating generic bundle {sys.platform}")

version = open("VERSION", "r").read().strip()
bundle_name = f"{bundle_name}{platform}-{version}"

a = Analysis(
    ["../backend/backend.py"],
    pathex=["./backend", "."],
    binaries=platform_binaries,
    datas=[
        ("../backend/resources/img/splash.txt", "backend/resources/img"),
        ("../VERSION", "."),
        ("../backend/resources/pptAssets/background.jpg", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/background_2.jpg", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/criteria.png", "backend/resources/pptAssets"),
        ("../backend/resources/pptAssets/criteria2.png", "backend/resources/pptAssets"),
    ],
    hiddenimports=[],
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
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=True, upx_exclude=[], name=bundle_name)

import shutil, sys, os

os.makedirs(f"{DISTPATH}/{bundle_name}/input/jobs")
os.makedirs(f"{DISTPATH}/{bundle_name}/input/thresholds")
os.makedirs(f"{DISTPATH}/{bundle_name}/backend/resources/controllerDefaults")
shutil.copyfile("input/jobs/DefaultJob.json", f"{DISTPATH}/{bundle_name}/input/jobs/DefaultJob.json")
shutil.copyfile("input/thresholds/DefaultThresholds.json", f"{DISTPATH}/{bundle_name}/input/thresholds/DefaultThresholds.json")
shutil.copyfile("backend/resources/controllerDefaults/defaultHealthRulesAPM.json", f"{DISTPATH}/{bundle_name}/backend/resources/controllerDefaults/defaultHealthRulesAPM.json")
shutil.copyfile("backend/resources/controllerDefaults/defaultHealthRulesBRUM.json", f"{DISTPATH}/{bundle_name}/backend/resources/controllerDefaults/defaultHealthRulesBRUM.json")
shutil.copyfile("backend/resources/pptAssets/background.jpg", f"{DISTPATH}/{bundle_name}/backend/resources/pptAssets/background.jpg")
shutil.copyfile("backend/resources/pptAssets/background_2.jpg", f"{DISTPATH}/{bundle_name}/backend/resources/pptAssets/background_2.jpg")
shutil.copyfile("backend/resources/pptAssets/criteria.png", f"{DISTPATH}/{bundle_name}/backend/resources/pptAssets/criteria.png")
shutil.copyfile("backend/resources/pptAssets/criteria2.png", f"{DISTPATH}/{bundle_name}/backend/resources/pptAssets/criteria2.png")
