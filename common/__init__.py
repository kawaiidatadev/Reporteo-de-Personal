import sqlite3  # Para menejar la base de datos
import ctypes  # Para usar MessageBox en Windows
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import shutil
import time
import win32com.client as win32
from datetime import datetime
import tkinter as tk
from PIL import Image, ImageTk
import threading
import subprocess
import pygame
import sys
import win32gui
import win32con
import random