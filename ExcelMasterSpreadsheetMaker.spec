# -*- mode: python -*-

block_cipher = None


a = Analysis(['ExcelMasterSpreadsheetMaker.py'],
             pathex=['H:\\backUp_pcwiek\\prywante\\Programming\\projects\\MasterExcelFile'],
             binaries=None,
             datas=None,
             hiddenimports=['openpyxl'],
             hookspath=[],
             runtime_hooks=[],
             excludes=['numpy','scipy','matplotlib','sphinx'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='ExcelMasterSpreadsheetMaker',
          debug=True,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='ExcelMasterSpreadsheetMaker')
