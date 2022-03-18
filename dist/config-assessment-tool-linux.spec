# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['../backend/backend.py'],
             pathex=['./backend', '.'],
             binaries=[],
             datas=[('../backend/resources/img/splash.txt', 'backend/resources/img'), ('../VERSION', '.')],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts, 
          [],
          exclude_binaries=True,
          name='config-assessment-tool',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas, 
               strip=False,
               upx=True,
               upx_exclude=[],
               name='config-assessment-tool-linux')


import shutil, sys, os
os.makedirs('{0}/config-assessment-tool-linux/input/jobs'.format(DISTPATH))
os.makedirs('{0}/config-assessment-tool-linux/input/thresholds'.format(DISTPATH))
os.makedirs('{0}/config-assessment-tool-linux/backend/resources/controllerDefaults'.format(DISTPATH))
shutil.copyfile('input/jobs/DefaultJob.json', '{0}/config-assessment-tool-linux/input/jobs/DefaultJob.json'.format(DISTPATH))
shutil.copyfile('input/thresholds/DefaultThresholds.json', '{0}/config-assessment-tool-linux/input/thresholds/DefaultThresholds.json'.format(DISTPATH))
shutil.copyfile('backend/resources/controllerDefaults/defaultHealthRulesAPM.json','{0}/config-assessment-tool-linux/backend/resources/controllerDefaults/defaultHealthRulesAPM.json'.format(DISTPATH)),
shutil.copyfile('backend/resources/controllerDefaults/defaultHealthRulesBRUM.json','{0}/config-assessment-tool-linux/backend/resources/controllerDefaults/defaultHealthRulesBRUM.json'.format(DISTPATH)),
