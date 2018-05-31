import unittest
from openpyxl import load_workbook


class TestTransform(unittest.TestCase):
    def testTryMappingColumns(self):
        wb = load_workbook('template.xlsx')
        ws = wb.active
        for col in ws.iter_cols():
            print(col[0].value)
if __name__ == '__main__':
    unittest.main()