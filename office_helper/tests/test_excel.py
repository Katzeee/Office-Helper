import unittest
import tempfile
import os
from utils.excel_op import ExcelOperator
from test_base import BaseTestCase

class ExcelOperatorTestCase(BaseTestCase):
    excel_file = "test.xlsx"
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        print("Create test file...")
        # 创建测试文件
        cls.excel_file = "test.xlsx"
        ExcelOperator.create_excel_file(cls.excel_file)

    @classmethod
    def tearDownClass(cls):
        print("\nDelete test file...")
        os.remove(cls.excel_file)
        super().tearDownClass()


    def test_read_write_cell(self):
        with ExcelOperator(self.excel_file) as excel:
            excel.write_cell_value('Sheet1', 'A1', 'Hello, xlwings!')
            value = excel.read_cell_value('Sheet1', 'A1')
            self.assertEqual(value, 'Hello, xlwings!')
        with ExcelOperator(self.excel_file) as excel:
            value = excel.read_cell_value('Sheet1', 'A1')
            self.assertEqual(value, 'Hello, xlwings!')

    def test_write_read_list(self):
        with ExcelOperator(self.excel_file) as excel:
            data = [['a', 'b', 'c'], [1, 2, 3], [4.5, 6.7, 8.9]]
            excel.write_cell_value('Sheet1' ,'B1', data)
            result = excel.read_cell_value('Sheet1', 'B1:D3')
            self.assertEqual(data, result)

