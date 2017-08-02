import unittest
import os

from excel_validation import excel_validation

class test_conditional_operator(unittest.TestCase):

    def test_operator_ANY_invalidate(self):
        excel = os.path.abspath('tests/samples/test_operator.xlsx')
        config = os.path.abspath('tests/samples/test_operator-1.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertFalse(result)
        self.assertEqual(errors[0]['cell'], 'A1')
        self.assertEqual(errors[1]['cell'], 'B1')
        self.assertEqual(errors[2]['cell'], 'C1')

    def test_operator_ANY_valid(self):
        excel = os.path.abspath('tests/samples/test_operator.xlsx')
        config = os.path.abspath('tests/samples/test_operator-2.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertTrue(result)

    def test_operator_ANY_range_valid(self):
        excel = os.path.abspath('tests/samples/test_operator.xlsx')
        config = os.path.abspath('tests/samples/test_operator-3.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertTrue(result)

    def test_operator_ANY_multiple_rules_valid(self):
        excel = os.path.abspath('tests/samples/test_operator.xlsx')
        config = os.path.abspath('tests/samples/test_operator-4.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertTrue(result)
    
    def test_operator_ANY_valid_other_rule_invalid(self):
        excel = os.path.abspath('tests/samples/test_operator.xlsx')
        config = os.path.abspath('tests/samples/test_operator-5.json')
        ev = excel_validation(excel, config)
        result, errors = ev.validate()
        self.assertFalse(result)
        self.assertEqual(errors[0]['cell'], 'A1')
        self.assertEqual(errors[1]['cell'], 'B1')
        self.assertEqual(errors[2]['cell'], 'C1')