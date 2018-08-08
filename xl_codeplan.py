from openpyxl import load_workbook
from enum import IntEnum
from codeplan import CodeplanNode, CodeplanElement, CodeplanNodeTypes
from settings import CODE_PREFIX
from collections import defaultdict

class XLCodeplanRowTypes(IntEnum):
    Invalid = 0
    Regular = 1
    Combine = 2
    NetStart = 3
    NetEnd = 4


class XLCodeplans():

    def __init__(self, path, *, code_column=1, delete_empty=True):
        self.path = path
        self.code_column = code_column
        self.delete_empty = delete_empty
        self._read_excel()
        self._errors = None
        self._is_valid = None
        if delete_empty:
            self._delete_empty()

    def _read_excel(self):
        self._codeplans = [
            XLCodeplan(
                name=sheet.title,
                excel_rows=[
                    XLCodeplanRow(
                        code=row[self.code_column - 1].value,
                        label=row[self.code_column].value,
                        index=index
                    )
                    for index, row in enumerate(sheet, start=1)
                ]
            )
            for sheet in load_workbook(self.path)
        ]

    def _delete_empty(self):
        for cp in self._codeplans:
            if len(cp) == 0:
                self._codeplans.remove(cp)

    @property
    def errors(self):
        if self._errors is None:
            self._errors = [f'Codeplan "{codeplan.name}": {error}'
                            for codeplan in self._codeplans if not codeplan.is_valid
                            for error in codeplan.errors]
        return self._errors

    @property
    def is_valid(self):
        if self._is_valid is None:
            self._is_valid = False if self.errors else True
        return bool(self._is_valid)

    def __len__(self):
        return len(self._codeplans)

    def __getitem__(self, i):
        return self._codeplans[i]

    def __repr__(self):
        return f"XLCodeplans(path='{self.path}', code_column={self.code_column}, delete_empty={self.delete_empty})"


class XLCodeplan():

    def __init__(self, name, excel_rows):
        self.name = name
        self._rows = self._trim_rows(excel_rows)
        self._elements = None
        self._tree = None
        self._errors = None
        self._is_valid = None

    def _trim_rows(self, rows):
        first_valid_row = last_valid_row = 0
        for row in rows:
            if row.is_valid:
                first_valid_row = row.index
                break
        if not first_valid_row:
            return []
        for row in reversed(rows):
            if row.is_valid:
                last_valid_row = row.index
                break
        return rows[first_valid_row - 1:last_valid_row]

    @property
    def elements(self):
        if self._elements is None:
            elements_with_label_list = defaultdict(list)
            for row in self._rows:
                if row.row_type == XLCodeplanRowTypes.Regular:
                    elements_with_label_list[row.code].append(row.label)
                elif row.row_type == XLCodeplanRowTypes.Combine:
                    for c in row.combine_codes:
                        elements_with_label_list[c].append('')
            self._elements = []
            for code, labels in elements_with_label_list.items():
                self._elements.append(
                    CodeplanElement(
                        code=f'{CODE_PREFIX}{code}',
                        label=labels[0],
                        doubled=True if len(labels) > 1 else False
                    )
                )
        self._elements.sort()
        return self._elements

    @property
    def tree(self):
        if self._tree is None:
            
            self._tree = CodeplanNode(
                name = '',
                label = '',
                node_type=CodeplanNodeTypes.Root,
                parent = None,
                level = 0
            )
            current_parent = self._tree

            base_node = CodeplanNode(
                name='',
                label='',
                node_type=CodeplanNodeTypes.Base,
                parent = current_parent,
                level = current_parent.level
            )
            self._tree.children.append(base_node)
            
            for row in self._rows:
                if row.row_type == XLCodeplanRowTypes.NetStart:
                    node = CodeplanNode(
                        name=f'net{row.index}',
                        label=row.label,
                        node_type=CodeplanNodeTypes.Net,
                        parent=current_parent,
                        level=current_parent.level + 1)
                    current_parent.children.append(node)
                    current_parent = node
                elif row.row_type == XLCodeplanRowTypes.NetEnd:
                    level_difference = current_parent.level - len(row.code) + 1
                    while level_difference:
                        current_parent = current_parent.parent
                        level_difference -= 1
                elif row.row_type == XLCodeplanRowTypes.Regular:
                    node = CodeplanNode(
                        name=f'{CODE_PREFIX}{row.code}',
                        label=row.label,
                        node_type=CodeplanNodeTypes.Regular,
                        parent=current_parent,
                        level=current_parent.level)
                    current_parent.children.append(node)
                elif row.row_type == XLCodeplanRowTypes.Combine:
                    combine_node = CodeplanNode(
                        name=f'comb{row.index}',
                        label=row.label,
                        node_type=CodeplanNodeTypes.Combine,
                        parent=current_parent,
                        level=current_parent.level)
                    for c in row.combine_codes:
                        child = CodeplanNode(
                            name=f'{CODE_PREFIX}{c}',
                            label='',
                            node_type=CodeplanNodeTypes.Regular,
                            parent=combine_node,
                            level=combine_node.level + 1)
                        combine_node.children.append(child)
                    current_parent.children.append(combine_node)

        return self._tree

    @property
    def errors(self):

        if self._errors is None:

            self._errors = []
            # Checks if codeplan is empty
            if len(self._rows) == 0:
                self._errors.append('Empty codeplan')

            # Check if codes are valid
            self._errors.extend([f'Invalid code "{row.code}" in row {row.index}'
                                 for row in self._rows if not row.is_valid])

            # Structural validation
            current_level = 0
            current_elements = []
            last_row_type = XLCodeplanRowTypes.Invalid
            for row in self._rows:
                if row.row_type == XLCodeplanRowTypes.NetStart:
                    current_elements = []
                    if len(row.code) != current_level + 1:
                        self._errors.append(
                            f'Invalid * in row {row.index}({"*"*(current_level + 1)} expected)')
                    current_level = len(row.code)
                elif row.row_type == XLCodeplanRowTypes.NetEnd:
                    if len(row.code) > current_level:
                        self._errors.append(
                            f'Invalid # in row {row.index}. (Less than {"#"*current_level} expected)')
                    if last_row_type == XLCodeplanRowTypes.NetStart:
                        self._errors.append(f'Empty net in row {row.index}')
                    current_level = len(row.code) - 1
                elif row.row_type == XLCodeplanRowTypes.Combine:
                    if len(row.combine_codes) != len(set(row.combine_codes)):
                        self._errors.append(
                            f'Duplicate codes found in row {row.index}')
                elif row.row_type == XLCodeplanRowTypes.Regular:
                    if row.code in current_elements:
                        self._errors.append(
                            f'Duplicate codes found in row {row.index}')
                    else:
                        current_elements.append(row.code)
                last_row_type = row.row_type

        return self._errors

    @property
    def is_valid(self):
        if self._is_valid is None:
            self._is_valid = False if self.errors else True
        return self._is_valid

    def __len__(self):
        return len(self.elements)

    def __getitem__(self, i):
        return self.elements[i]

    def __repr__(self):
        return f"XLCodeplan(name='{self.name}'), len={len(self)}"


class XLCodeplanRow():

    def __init__(self, code, label, index):
        self.code = str(code).strip() if code is not None else ''
        self.label = str(label).strip() if label is not None else ''
        self.index = index

        self._row_type = None
        self._combine_codes = None
        self._is_valid = None

    @property
    def row_type(self):
        if self._row_type is None:
            # regular/numeric rows
            if self.code.isdigit():
                self._row_type = XLCodeplanRowTypes.Regular
            # start (*) and end tags (#) for net()
            elif self.code and self.code in '*******':
                self._row_type = XLCodeplanRowTypes.NetStart
            elif self.code and self.code in '#######':
                self._row_type = XLCodeplanRowTypes.NetEnd
            # floats (23.35) for combine()
            elif '.' in self.code and self.code.replace('.', '', 1).isdigit():
                self._row_type = XLCodeplanRowTypes.Combine
            # "142, 43, 5" pattern for combine()
            elif ',' in self.code and len(self.code.split(',')) == len(
                    [c.strip() for c in self.code.split(',') if c.strip().isdigit()]):
                self._row_type = XLCodeplanRowTypes.Combine
            else:
                self._row_type = XLCodeplanRowTypes.Invalid
        return self._row_type

    @property
    def combine_codes(self):
        if self._combine_codes is None:
            if self.row_type != XLCodeplanRowTypes.Combine:
                self._combine_codes = []
            elif '.' in self.code and self.code.replace('.', '', 1).isdigit():
                self._combine_codes = self.code.split('.')
            elif ',' in self.code:
                self._combine_codes = [
                    c.strip() for c in self.code.split(',') if c.strip().isdigit()]
        return self._combine_codes

    @property
    def is_valid(self):
        if self._is_valid is None:
            self._is_valid = bool(self.row_type)
        return self._is_valid

    def __repr__(self):
        return f"XLCodeplanRow(code='{self.code}', label='{self.label}', index={self.index})"


def test():

    tests = [
        XLCodeplans(f'Examples\\Codeplan KTV online 201807.xlsx'),
        XLCodeplans(f'Examples\\8810xxxx_MOET_CP_nach Coding_MY_2018_2018-05-29_CSS.xlsx', code_column=2),
        XLCodeplans(f'Examples\\88107393_AT_Mueller P2_DE_COFR_2018-07.xlsx'),
        XLCodeplans(f'Examples\\318107382-88C AT_McDonalds_CFrame_KW24-29.xlsx')
    ]

    print('------------ERRORS----------------')
    for t in tests:
        print(t.path)
        if t.errors:
            for e in t.errors:
                print(f'- {e}')
        else:
            print('- OK')
        print('----------------------------------')


    print('OK')

if __name__ == '__main__':
    from timeit import timeit
    print(timeit(test, number=1))
    print('OK')

