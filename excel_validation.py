import json
import openpyxl
import argparse
#from openpyxl.formula import Tokenizer 

class excel_validation():

    def __init__(self, file_path, config_path):
        self.wb = openpyxl.load_workbook(filename = file_path, read_only=True, data_only = True)        
        with open(config_path) as data_file:    
            self.sheet_configs = json.load(data_file)
        self.result = True
        self.errors = []

    def validate(self):
        for sheet_config in self.sheet_configs:
            sheet = self.wb[sheet_config['name']]
            for rule in sheet_config['rules']:
                oprtor = None
                satisfy_ANY = False
                if rule.has_key('operator'):
                    oprtor = rule['operator']
                for cells in rule['cells']:
                    if ':' in cells:
                        for x in sheet[cells]:
                            for y in x:                                                        
                                result = self.check_rule(y, rule['conds'])
                                if oprtor == 'ANY' and result:
                                    satisfy_ANY = True
                                    break
                            if satisfy_ANY:
                                break
                    else:                    
                        result = self.check_rule(sheet[cells], rule['conds'])
                        if oprtor == 'ANY' and result:
                            satisfy_ANY = True
                            break
                # check if ANY satisfy for passing this rule
                if satisfy_ANY:
                    self.result = True
                    self.errors = [] # clean any error that have been inserted in check_rule
        return self.result, self.errors

    def check_rule(self, cell, conds): 
        #print '@' + str(cell.column) + '-' + str(cell.row) + ':'
        #string 0.7999999999999 => 0.8
        vtype = type(cell.value).__name__
        exec('value = ' + vtype + '(str(cell.value))')   
        r = True
        for cond in conds:
            exec('result = ' + cond)  
            if not result:                
                self.errors.append({'cell':cell.coordinate, 'value': value, 'cond': cond})
                self.result = False
                r = False

        return r

if __name__ == '__main__':

    parser = argparse.ArgumentParser()        
    parser.add_argument("-f", "--file", help="excel file path (xlsx)", type=str, required=True)
    parser.add_argument("-c", "--config", help="config file path (please see tests)", type=str, required=True)
    args = parser.parse_args()

    ev = excel_validation(args.file, args.config)
    print ev.validate()