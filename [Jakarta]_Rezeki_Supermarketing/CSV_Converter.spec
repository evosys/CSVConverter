# -*- mode: python -*-

block_cipher = pyi_crypto.PyiBlockCipher(key='68b00c755cef892e512d56621925d836')


a = Analysis(['app-inc-java.py'],
             pathex=['D:\\Development\\Python\\CSVConverter\\[Jakarta]_Rezeki_Supermarketing'],
             binaries=[],
             datas=[('tabula/bin/tabula-1.0.2-jar-with-dependencies.jar', 'tabula/bin/')],
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
          name='CSV_Converter',
          debug=False,
          strip=False,
          upx=False,
          runtime_tmpdir=None,
          console=False , version='version.txt', icon='resources\\icon.ico')
