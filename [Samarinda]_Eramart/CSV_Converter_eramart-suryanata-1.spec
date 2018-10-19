# -*- mode: python -*-

block_cipher = None


a = Analysis(['eramart_suryanata_1.py'],
             pathex=['E:\\Python\\CSVConverter\\[Samarinda]_Eramart'],
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
          exclude_binaries=True,
          name='CSV_Converter_eramart-suryanata-1',
          debug=False,
          strip=False,
          upx=False,
          console=True , version='version.txt', icon='resources\\icon.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               name='CSV_Converter_eramart-suryanata-1')
