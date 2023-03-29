import unittest
from excel_project import *


class AppXtest(unittest.TestCase):

    def test_Load_button(self):
        """
        test:Load button(source_file)
        """
        file1_path_split = file1.split('/')
        actual_name_file_1 = file1_path_split[-1]
        expected = 'measurements.xlsx'
        self.assertEqual(expected, actual_name_file_1)

    def test_Report_button(self):
        """
        test:Report file button(final_report)
        """
        file2_path_split = file2.split('/')
        actual_name_file_2 = file2_path_split[-1]
        expected = 'report.xlsx'
        self.assertEqual(expected, actual_name_file_2)

    def test_source_file_record(self):
        """
        test:write data in  measurements.xlsx
        """
        wb = openpyxl.load_workbook(file1)
        sheet = wb.active
        no_value = sheet['A2'].value
        degree_value = sheet['B2'].value
        self.assertIsNotNone(no_value)
        self.assertIsNotNone(degree_value)

    def test_date_hour_file_record(self):
        """
        test:write data in date.xlsx
        """
        wb = openpyxl.load_workbook(file3)
        sheet = wb.active
        no_value=sheet['A2'].value
        hour_value=sheet['B2'].value
        date_value=sheet['C2'].value
        epoch_value=sheet['D2'].value
        cells_values=[no_value,hour_value,date_value,epoch_value]
        for i in cells_values:
            self.assertIsNotNone(i)

    def test_range_letter(self):
        """
        :test: range letter
        """
        start = 'A'
        stop = 'F'
        lista = [i for i in range_letter(start, stop)]
        assert lista == ['A', 'B', 'C', 'D', 'E', 'F']

    def test_report(self):
        """
        :test: write data in report.xlsx
        """
        wb = openpyxl.load_workbook(file2)
        sheet = wb.active
        no_value=sheet['B2'].value
        temp_value=sheet['B3'].value
        hour_value=sheet['B4'].value
        date_value=sheet['B5'].value
        nominal_value=sheet['B8'].value
        ul_value=sheet['B9'].value
        ll_value=sheet['B10'].value

        cells_values=[no_value,temp_value,hour_value,date_value,nominal_value,ul_value,ll_value]
        for i in cells_values:
            self.assertIsNotNone(i)

    def test_date_hour_file_check_values(self):
        """
        :test: check values from date.xlsx
        """
        wb = openpyxl.load_workbook(file3)
        sheet = wb.active
        no=sheet['A2'].value
        expected_no='1.'
        self.assertEqual(no,expected_no)
        hour_value=sheet['B2'].value
        self.assertIn(int(hour_value[:2]),range(24))
        self.assertIn(int(hour_value[3:5]),range(61))
        self.assertIn(int(hour_value[6:]),range(61))
        date_value=sheet['C2'].value
        self.assertIn(int(date_value[:2]), range(32))
        self.assertIn(int(date_value[3:5]), range(13))
        self.assertIn(int(date_value[6:]), range(22,27))

    def test_source_file_check_values(self):
        """
        :test: check values from measurements.xlsx
        """
        wb = openpyxl.load_workbook(file1)
        sheet = wb.active
        no = sheet['A2'].value
        expected_no = '1.'
        self.assertEqual(no, expected_no)
        temp_value=sheet['B2'].value
        self.assertIn(temp_value,range(18,23))


if __name__ == '__main__':
    unittest.main()
