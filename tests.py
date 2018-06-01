import unittest
from openpyxl import load_workbook
import sys, os
sys.path.append(os.getcwd())
from transform import *

class TestTransform(unittest.TestCase):
    def setUp(self):
        self.t = SapDataParser('data.xlsx')
    def testTryMappingColumns(self):
        wb = load_workbook('template.xlsx')
        ws = wb.active
        for col in ws.iter_cols():
            print(col[0].value)
    #def testTransformData(self):
    #    self.t.start()
if __name__ == '__main__':
    unittest.main()