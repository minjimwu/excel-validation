import unittest
import os

from excel_validation import excel_validation

class test_excel_validation_num(unittest.TestCase):

    def test_num_invalidate(self):
        excel = os.path.abspath('tests/samples/test_num.xlsx')
        config = os.path.abspath('tests/samples/test_num-1.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertFalse(result)
        self.assertEqual(errors[0]['cell'], 'A1')
        self.assertEqual(errors[1]['cell'], 'B1')
        self.assertEqual(errors[2]['cell'], 'E1')

    def test_num_range(self):
        excel = os.path.abspath('tests/samples/test_num.xlsx')
        config = os.path.abspath('tests/samples/test_num-2.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertTrue(result)

    def test_num_single(self):
        excel = os.path.abspath('tests/samples/test_num.xlsx')
        config = os.path.abspath('tests/samples/test_num-3.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertTrue(result)
        
    def test_num_firstcellpass_invalidate(self):
        excel = os.path.abspath('tests/samples/test_num.xlsx')
        config = os.path.abspath('tests/samples/test_num-4.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertFalse(result)
        self.assertEqual(errors[0]['cell'], 'B2')
        self.assertEqual(errors[1]['cell'], 'E2')