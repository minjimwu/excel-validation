import unittest
import os

from excel_validation import excel_validation

class test_excel_validation_num(unittest.TestCase):

    def test_string_invalidate(self):
        excel = os.path.abspath('tests/samples/test_string.xlsx')
        config = os.path.abspath('tests/samples/test_string-1.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertFalse(result)
        self.assertEqual(errors[0]['cell'], 'A1')
    
    def test_string_single(self):
        excel = os.path.abspath('tests/samples/test_string.xlsx')
        config = os.path.abspath('tests/samples/test_string-2.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertTrue(result)

    def test_string_range(self):
        excel = os.path.abspath('tests/samples/test_string.xlsx')
        config = os.path.abspath('tests/samples/test_string-3.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertTrue(result)

        
  