import unittest
from openpyxl import load_workbook
import sys, os
sys.path.append(os.getcwd())
from transform import *

LIMIT = 49647 #This must be set to the size of the data.xlsx file
#data.xlsx must be in .\
#outout.xlsx must be in .\emptyOutput
t = SapDataParser('data.xlsx','output.xlsx')
t.start(limit = LIMIT)