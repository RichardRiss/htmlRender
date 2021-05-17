# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['htmlRender.py'],
             pathex=['C:\\Users\\rissr\\Projekte\\Skripte\\htmlRender'],
             binaries=[],
             datas=[('C:\\PROGRA~1\\Python38\\venv\\Lib\\site-packages\\pandas\\io\\formats\\templates\\html.tpl', '.')],
             hiddenimports=['jinja2', 'matplotlib'],
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
          name='htmlRender',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True , icon='niles_ori.ico')
