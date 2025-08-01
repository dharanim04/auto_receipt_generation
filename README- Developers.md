$ pip install pandas python-docx tkinter docx2pdf openpyxl

# To convert to exec 
$ pip install pyinstaller

$ pyinstaller --onefile generate_dynamic.py


$ pyinstaller --onefile generate_dynamic.py
# ERROR: The 'pathlib' package is an obsolete backport of a standard library package and is incompatible with PyInstaller.

# to resolve this error 
$ pip uninstall pathlib


# to remove cache and generate exe
$ pyinstaller --clean --onefile generate_dynamic.py


$ pyinstaller --clean --onedir generate_dynamic.py
use "onedir" for faster application run time

# running project in python
python .\generate_dynamic.py


$ pyinstaller --clean --onedir --add-data "data_folder:data_folder" generate_dynamic.py