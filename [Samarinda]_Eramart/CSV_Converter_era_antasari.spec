# -*- mode: python -*-

block_cipher = pyi_crypto.PyiBlockCipher(key='68b00c755cef892e512d56621925d836')


a = Analysis(['era_antasari.py'],
             pathex=['D:\\Development\\Python\\CSVConverter\\[Samarinda]_Eramart'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='CSV_Converter_era_antasari',
          debug=False,
          strip=False,
          upx=False,
          runtime_tmpdir=None,
          console=False , version='version.txt', icon='resources\\icon.ico')
