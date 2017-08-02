# excel-validation
validation for excel cell

Installation
------------
install python 2.7 first

    pip install openpyxl
    pip install argparse
    

Usage (Example)
------------
configs

    [
      {
        "name": "Sheet1", #sheet name
        "rules": [
          {
            "cells": ["A1:E1"], #can be single value or range
            "conds": ["value >= 0.8 and value <= 1"] #result should be True or False
          }
        ]
      }
    ]
    
command

    python excel_validation.py -f tests/samples/test_num.xlsx -c tests/samples/test_num-1.json
    
result (return True or False, Errors)`

    (False,
    [{'cell': 'A1', 'cond': u'value >= 0.8 and value <= 1', 'value': 0.1}, 
    {'cell': 'B1', 'cond': u'value >= 0.8 and value <= 1', 'value': 0.2}, 
    {'cell': 'E1', 'cond': u'value >= 0.8 and value <= 1', 'value': 0.5}])
    
