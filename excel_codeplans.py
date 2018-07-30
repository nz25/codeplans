from os.path import exists
from openpyxl import load_workbook
from collections import namedtuple
import numbers

CodeplanRow = namedtuple('CodeplanRow', 'index code label')

class ExcelCodeplans():

    def __init__(self, path):
        self.path = path
        self._read_excel()

    def _read_excel(self):
        self.codeplans = {sheet.title:
            ExcelCodeplan(sheet.title, [
                CodeplanEntry(
                    code=row[0].value,
                    label=row[1].value,
                    index=index)
                    for index, row in enumerate(sheet, start=1)])
                for sheet in load_workbook(self.path)}
    
class ExcelCodeplan():

    def __init__(self, name, entries):
        self.name = name
        self.original_entries = entries
        self.entries = self._trim_entries()

        self.errors = []
        self._validate_content()
        self._validate_structure()
        if self.errors:
            print(self.errors)

    def _trim_entries(self):
        first_valid_entry = last_valid_entry = 0
        for entry in self.original_entries:
            if entry.is_valid():
                first_valid_entry = entry.index
                break
        if not first_valid_entry:
            return []
        for entry in reversed(self.original_entries):
            if entry.is_valid():
                last_valid_entry = entry.index
                break
        return self.original_entries[first_valid_entry - 1:last_valid_entry]

    def _validate_content(self):
        for entry in self.entries:
            if not entry.is_valid():
                self.errors.append(f'Error in code "{entry.code}" at row {entry.index}')
        
    def _validate_structure(self):
        for entry in self.entries:
            pass

class CodeplanEntry():

    def __init__(self, code, label, index):
        self.code = str(code).strip()
        self.label = label
        self.index = index

    def __repr__(self):
        return f'CodeplanEntry(code={self.code})'

    def is_valid(self):
        # numeric for simple elements and floats (23.35) for combine()
        if self.code.replace('.', '', 1).isdigit():
            return True
        # start (*) and end tags (#) for net()
        if self.code in '#######' or self.code in '*******':
            return True
        # "142, 43, 5" pattern for combine()
        if ',' in self.code and len(self.code.split(',')) == len(
            [c for c in self.code.split(',') if c.strip().isdigit()]):
            return True
        return False

    def entry_type(self):
        pass

if __name__ == '__main__':
    xl = ExcelCodeplans('Codeplan KTV online 201807_VORAB.xlsx')
    print('ok')