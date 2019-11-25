from distutils.core import setup
import py2exe
setup(console=['tool.py'], requires=['pyodbc', 'wx', 'openpyxl'])