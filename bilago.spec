# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['bilago.py'],
             pathex=['C:\\Users\\Anders\\Dropbox\\Python\\Projects\\bilago'],
             binaries=[],
             datas=[('C:/Users/Anders/Dropbox/Python/Projects/bilago/static/images/gear_loader.png', 'static/images'), ('C:/Users/Anders/Dropbox/Python/Projects/bilago/static/images/header.png', 'static/images')],
             hiddenimports=['pkg_resources.py2_warn'],
             hookspath=[],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='bilago',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True , icon='bilagoicon.ico')
