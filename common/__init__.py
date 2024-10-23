import sqlite3 # Para menejar la base de datos
import ctypes  # Para usar MessageBox en Windows
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import shutil
import os
from datetime import datetime