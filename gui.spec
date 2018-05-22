# -*- mode: python -*-

block_cipher = None


a = Analysis(['gui.py'],
             pathex=['C:\\Users\\a.moullard\\Desktop\\xls_python\\Xls_manipulation_python'],
             binaries=[],
             datas=[],
             hiddenimports=['pyexcel_io.database.exporters.sqlalchemy', 'pyexcel_io.database.exporters.django', 'pyexcel_io.database.importers.sqlalchemy', 'pyexcel_io.database.importers.django', 'pyexcel_io.readers.tsvz', 'pyexcel_io.readers.csvz', 'pyexcel_io.readers.tsv', 'pyexcel_io.readers.csvr', 'pyexcel_io.readers.csvz', 'pyexcel_io.readers.tsv', 'pyexcel_io.readers.tsvz', 'pyexcel_io.writers.csvw', 'pyexcel.ext.xls', 'pyexcel.ext.xlsx', 'pyexcel.plugins.sources.file_input', 'pyexcel.plugins.parsers.excel', 'pyexcel_xls', 'pyexcel_xls.xlsr', 'pyexcel.plugins.renderers.sqlalchemy', 'pyexcel.plugins.renderers.django', 'pyexcel.plugins.renderers.excel', 'pyexcel.plugins.renderers._texttable', 'pyexcel.plugins.parsers.excel', 'pyexcel.plugins.parsers.sqlalchemy', 'pyexcel.plugins.sources.http', 'pyexcel.plugins.sources.file_input', 'pyexcel.plugins.sources.memory_input', 'pyexcel.plugins.sources.file_output', 'pyexcel.plugins.sources.output_to_memory', 'pyexcel.plugins.sources.pydata.bookdict', 'pyexcel.plugins.sources.pydata.dictsource', 'pyexcel.plugins.sources.pydata.arraysource', 'pyexcel.plugins.sources.pydata.records', 'pyexcel.plugins.sources.django', 'pyexcel.plugins.sources.sqlalchemy', 'pyexcel.plugins.sources.sqlalchemy'],
             hookspath=['.'],
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
          name='gui',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
