# XLS python
This code is for used to concatenate production data in a specific file for each product.
working with python v3.6.3

Source:
https://github.com/AlexandreMoullard/Xls_manipulation_python.git

Package to install:
anaconda
pyexcel

Installer cmd on pyinstaller:
pyinstaller --onefile --noconsole --additional-hooks-dir=. --hidden-import pyexcel_io.database.exporters.sqlalchemy --hidden-import pyexcel_io.database.exporters.django --hidden-import pyexcel_io.database.importers.sqlalchemy --hidden-import pyexcel_io.database.importers.django --hidden-import pyexcel_io.readers.tsvz --hidden-import pyexcel_io.readers.csvz --hidden-import pyexcel_io.readers.tsv --hidden-import pyexcel_io.readers.csvr --hidden-import pyexcel_io.readers.csvz --hidden-import pyexcel_io.readers.tsv --hidden-import pyexcel_io.readers.tsvz --hidden-import pyexcel_io.writers.csvw --hidden-import pyexcel.ext.xls --hidden-import pyexcel.ext.xlsx  --hidden-import pyexcel.plugins.sources.file_input --hidden-import pyexcel.plugins.parsers.excel --hidden-import pyexcel_xls --hidden-import pyexcel_xls.xlsr --hidden-import pyexcel.plugins.renderers.sqlalchemy --hidden-import pyexcel.plugins.renderers.django --hidden-import pyexcel.plugins.renderers.excel --hidden-import pyexcel.plugins.renderers._texttable --hidden-import pyexcel.plugins.parsers.excel --hidden-import pyexcel.plugins.parsers.sqlalchemy --hidden-import pyexcel.plugins.sources.http --hidden-import pyexcel.plugins.sources.file_input --hidden-import pyexcel.plugins.sources.memory_input --hidden-import pyexcel.plugins.sources.file_output --hidden-import pyexcel.plugins.sources.output_to_memory --hidden-import pyexcel.plugins.sources.pydata.bookdict --hidden-import pyexcel.plugins.sources.pydata.dictsource --hidden-import pyexcel.plugins.sources.pydata.arraysource --hidden-import pyexcel.plugins.sources.pydata.records --hidden-import pyexcel.plugins.sources.django --hidden-import pyexcel.plugins.sources.sqlalchemy --hidden-import pyexcel.plugins.sources.sqlalchemy gui.py
keep images in same directory

Testing:
run test_hums.py with tested_files folder in same directory
16 tests should work