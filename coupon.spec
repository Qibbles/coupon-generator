# -*- mode: python -*-

block_cipher = None


a = Analysis(['coupon.py'],
             pathex=['C:\\Users\\gregory.chua\\odrive\\Projects\\Projects (gregmisc19@gmail.com)\\Projects\\Python\\coupon-generator\\coupon-generator'],
             binaries=[],
             datas=[('coupon-generator.json', '.'), ('mmq.ico', '.')],
             hiddenimports=[],
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
          name='coupon',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='mmq.ico')
