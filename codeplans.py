from enum import IntEnum
from win32com import client
from openpyxl import load_workbook
from collections import defaultdict, namedtuple

CODE_PREFIX = 'CB_'
HELPER_FIELD =  '.Coding'

def sort_element(code):
    return int(code[len(CODE_PREFIX):])

############################################################################
#
#                               ENUMERATIONS
#
############################################################################

class CodeplanNodeTypes(IntEnum):
    Root = 0
    Regular = 1
    Combine = 2
    Net = 3


class CodeplanSources(IntEnum):
    XL = 1
    MDD = 2
    Master = 3


class CFileSources(IntEnum):
    Verbaco = 1
    Ascribe = 2


class AxisSeparators(IntEnum):
    Comma = 1
    NetStart = 2
    NetEnd = 3


class XLCodeplanRowTypes(IntEnum):
    Invalid = 0
    Regular = 1
    Combine = 2
    NetStart = 3
    NetEnd = 4


class ObjectTypesConstants(IntEnum):
    mtVariable = 0
    mtArray = 1
    mtGrid = 2
    mtClass = 3
    mtElement = 4
    mtElements = 5
    mtLabel = 6
    mtField = 7
    mtHelperFields = 8
    mtFields = 9
    mtTypes = 10
    mtProperties = 11
    mtRouting = 12
    mtContexts = 13
    mtLanguages = 14
    mtLevelObject = 15
    mtVariableInstance = 16
    mtRoutingItem = 17
    mtCompound = 18
    mtElementInstance = 19
    mtElementInstances = 20
    mtLanguage = 21
    mtRoutingItems = 22
    mtRanges = 23
    mtCategories = 24
    mtCategoryMap = 25
    mtDataSources = 26
    mtDocument = 27
    mtVersion = 29
    mtVersions = 30
    mtVariables = 31
    mtDataSource = 32
    mtAliasMap = 33
    mtIndexElement = 34
    mtIndicesElements = 35
    mtPages = 36
    mtParameters = 37
    mtPage = 38
    mtItems = 39
    mtContext = 40
    mtContextAlternatives = 41
    mtElementList = 42
    mtGoto = 43
    mtTemplate = 44
    mtTemplates = 45
    mtStyle = 46
    mtNote = 47
    mtNotes = 48
    mtIfBlock = 49
    mtConditionalRouting = 50
    mtDBElements = 51
    mtDBQuestionDataProvider = 52
    mtUnknown = 65535


class DataTypeConstants(IntEnum):
    mtNone = 0
    mtLong = 1
    mtText = 2
    mtCategorical = 3
    mtObject = 4
    mtDate = 5
    mtDouble = 6
    mtBoolean = 7

class ElementTypeConstants(IntEnum):
    mtCategory = 0
    mtAnalysisSubheading = 1
    mtAnalysisBase = 2
    mtAnalysisSubtotal = 3
    mtAnalysisSummaryData = 4
    mtAnalysisTotal = 6
    mtAnalysisMean = 7
    mtAnalysisStdDev = 8
    mtAnalysisStdErr = 9
    mtAnalysisSampleVariance = 10
    mtAnalysisMinimum = 11
    mtAnalysisMaximum = 12
    mtAnalysisCategory = 14

class openConstants(IntEnum):
     oREAD = 1
     oREADWRITE = 2
     oNOSAVE = 3



############################################################################
#
#                               CODEPLAN
#
############################################################################


class CodeplanNode:

    def __init__(self, code, label, node_type, parent, level):
        self.code = code
        self.label = label
        self.node_type = node_type
        self.parent = parent
        self.level = level
        self.children = []
        self._flat_children = None
        self._axis = None

    @property
    def axis(self):
        if self._axis is None:
            label = self.label.replace(r"'", r"''")
            if self.node_type == CodeplanNodeTypes.Root:
                self._axis = f'{{base(), {",".join(c.axis for c in self.children)}}}'
            elif self.node_type == CodeplanNodeTypes.Net:
                self._axis = f'{self.code} \'{label}\' net({{{",".join(c.axis for c in self.children)}}})'
            elif self.node_type == CodeplanNodeTypes.Combine:
                self._axis = f'{self.code} \'{label}\' combine({{{",".join(c.axis for c in self.children)}}})'
            elif self.node_type == CodeplanNodeTypes.Regular:
                self._axis = f'{self.code}'
        return self._axis

    @property
    def flat_children(self):
        # deep first traversal
        if self._flat_children is None:
            self._flat_children = []
            stack = [*self.children]
            while stack:
                current = stack[0]
                stack = stack[1:]
                self._flat_children.append(current)
                for child in reversed(current.children):
                    stack.insert(0, child)
        return self._flat_children

    def __repr__(self):
        return f"CodeplanNode(code='{self.code}', label='{self.label}', node_type={self.node_type}, parent={self.parent}, level={self.level}), len={len(self.children)}"


class CodeplanElement:

    def __init__(self, code, label, double=False):
        self.code = code
        self.label = label
        self.double = double

    def __gt__(self, other):
        return sort_element(self.code) > sort_element(other.code)

    def __repr__(self):
        return f"CodeplanElement(code='{self.code}', label='{self.label}', double={self.double})"



############################################################################
#
#                               MDD CODEPLAN
#
############################################################################

class MDDFile():

    def __init__(self, mdd_path,):
        self.mdd_path = mdd_path
        self._read_mdd()
        self._variable_map = None
        self._errors = None
        self._is_valid = None

    def _read_mdd(self):
        mdd = client.Dispatch('MDM.Document')
        mdd.Open(self.mdd_path, mode=openConstants.oREAD)
        self._codeplans = [
            MDDCodeplan(
                name=t.Name,
                elements=[
                    CodeplanElement(
                        code=e.Name,
                        label=e.label
                    )
                    for e in t.Elements
                ],
                axis={f.AxisExpression
                    for f in mdd.Fields
                    if f.ObjectTypeValue == ObjectTypesConstants.mtVariable
                        and f.Elements.Reference.Name == t.Name
                }.pop()
            ) for t in mdd.Types
        ]
        self.variables = [
            MDDVariable(
                name=f.Name,
                label=f.Label,
                type_name=f.Elements.Reference.Name,
                axis=f.AxisExpression
            )
            for f in mdd.Fields
            if f.ObjectTypeValue == ObjectTypesConstants.mtVariable
        ]

        mdd.Close()

    @property
    def variable_map(self):
        
        if self._variable_map is None:

            self._variable_map = {}

            # saves list of variables per field
            field_variables = defaultdict(list)
            for v in self.variables:
                field_variables[v.field_name].append(v)

            # saves list of types per field
            field_types = defaultdict(set)
            for v in self.variables:
                field_types[v.field_name].add(v.type_name)

            # merges both results in temporary fields list
            Field = namedtuple('Field', 'name variables types')
            fields = [
                Field(
                    name=field_name,
                    variables=variables,
                    types=field_types[field_name]
                )
                for field_name, variables
                in field_variables.items()
            ]

            multitype_variables = [
                v
                for f in fields
                if len(f.types) > 1
                for v in f.variables
            ]

            for v in self.variables:
                if v in multitype_variables:
                    self._variable_map[f'{v.label}'] = v.compliant_name

        return self._variable_map

    def append_mdd(self, path):

        mdd = client.Dispatch('MDM.Document')
        mdd.Open(path, mode=openConstants.oREAD)

        new_codeplans = [
            MDDCodeplan(
                name=t.Name,
                elements=[
                    CodeplanElement(
                        code=e.Name,
                        label=e.label)
                    for e in t.Elements
                    ],
                axis={f.AxisExpression
                    for f in mdd.Fields
                    if f.ObjectTypeValue == ObjectTypesConstants.mtVariable
                        and f.Elements.Reference.Name == t.Name
                }.pop())
            for t in mdd.Types
            if t.Name not in [cp.name for cp in self._codeplans]
        ]
        self._codeplans.extend(new_codeplans)

        new_variables = [
            MDDVariable(
                name=f.Name,
                label=f.Label,
                type_name=f.Elements.Reference.Name,
                axis=f.AxisExpression
            )
            for f in mdd.Fields
            if f.ObjectTypeValue == ObjectTypesConstants.mtVariable
                and f.Name not in [v.name for v in self.variables]
        ]
        self.variables.extend(new_variables)

        mdd.Close()
        self._variable_map = None
        self._errors = None
        self._is_valid = None

    def save_mdd(self, path):
        mdd = client.Dispatch('MDM.Document')
        mdd.IncludeSystemVariables = False
        for cp in self._codeplans:
            new_list = mdd.CreateElements(cp.name)
            for element in cp:
                new_element = mdd.CreateElement(element.code, element.label)
                new_element.Type = ElementTypeConstants.mtCategory
                new_list.Add(new_element)
            mdd.Types.Add(new_list)
            
        for v in self.variables:
            new_variable = mdd.CreateVariable(v.name, self.variable_map.get(v.label, v.label + HELPER_FIELD))
            new_variable.DataType = DataTypeConstants.mtCategorical
            new_variable.Elements.ReferenceName = v.type_name
            new_variable.AxisExpression = [cp for cp in self._codeplans if cp.name == v.type_name][0].axis
            mdd.Fields.Add(new_variable)

        mdd.CategoryMap.AutoAssignValues()
        mdd.Save(path)
        mdd.Close()

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
        if isinstance(i, str):
            return [cp for cp in self._codeplans if cp.name == i][0]
        else:
            return self._codeplans[i]

    def __repr__(self):
        return f"MDDFile(path='{self.mdd_path}')"

    def __contains__(self, value):
        return bool([cp for cp in self._codeplans if cp.name == str(value)])

class MDDCodeplan:

    def __init__(self, name, elements, axis):
        self.name = name
        self._elements = elements
        self.axis = axis
        self._errors = None
        self._tree = None
        self._is_valid = None
        self.variable_map = {}
        self.category_map = {}

    @property
    def errors(self):
        if self._errors is None:
            self._errors = []
        return self._errors

    @property
    def tree(self):
        if self._tree is None:
            axis = self.axis[1:-1]
            label_bitmap = self._build_label_bitmap(axis)
            split_axis = self._split_axis(axis, label_bitmap)
            self._tree = self._build_tree(split_axis)

        return self._tree

    @property
    def elements(self):
        return self._elements

    @property
    def flat_tree(self):
        return self.tree.flat_children

    def _build_label_bitmap(self, axis):
        in_label = False
        consecutive_quotes = 0
        label_bitmap = []
        for c in axis:
            if c == "'":
                in_label = True
                consecutive_quotes += 1
            else:
                if in_label and not consecutive_quotes % 2:
                    in_label = not in_label
                    consecutive_quotes = 0
            label_bitmap.append(in_label)
        return label_bitmap

    def _split_axis(self, axis, label_bitmap):
        last_character = 0
        split_axis = []
        for i in range(len(axis)):
            in_label = label_bitmap[i]
            if not in_label:
                current_character = axis[i]
                last_2_characters = axis[i-1:i+1]
                last_5_characters = axis[i-4:i+1]
                if current_character == ',':
                    split_axis.append(
                        (AxisSeparators.Comma, axis[last_character:i].strip()))
                    last_character = i + 1
                elif last_5_characters == 'net({':
                    split_axis.append(
                        (AxisSeparators.NetStart, axis[last_character:i-4].strip()))
                    last_character = i + 1
                elif last_2_characters == '})':
                    split_axis.append(
                        (AxisSeparators.NetEnd, axis[last_character:i-1].strip()))
                    last_character = i + 1

        if last_character < len(axis):
            split_axis.append(axis[last_character:])

        return split_axis

    def _build_tree(self, split_axis):

        root_node = CodeplanNode(
            code='(root)',
            label='',
            node_type=CodeplanNodeTypes.Root,
            parent=None,
            level=0
        )
        current_parent = root_node

        for a in split_axis:
            if a[1] == 'base()':
                pass
            elif a[0] == AxisSeparators.Comma and a[1]:
                # regular element
                code, label = a[1].split(sep=' ', maxsplit=1)
                node = CodeplanNode(
                    code=code.strip(),
                    label=label.strip()[1:-1].replace("''", "'"),
                    node_type=CodeplanNodeTypes.Regular,
                    parent=current_parent,
                    level=current_parent.level + 1
                )
                current_parent.children.append(node)
            elif a[0] == AxisSeparators.NetStart:
                # net element
                code, label = a[1].split(sep=' ', maxsplit=1)
                node = CodeplanNode(
                    code=code.strip(),
                    label=label.strip()[1:-1].replace("''", "'"),
                    node_type=CodeplanNodeTypes.Net,
                    parent=current_parent,
                    level=current_parent.level + 1
                )
                current_parent.children.append(node)
                current_parent = node
            elif a[0] == AxisSeparators.NetEnd:
                if a[1]:
                    code, label = a[1].split(sep=' ', maxsplit=1)
                    node = CodeplanNode(
                        code=code.strip(),
                        label=label.strip()[1:-1].replace("''", "'"),
                        node_type=CodeplanNodeTypes.Regular,
                        parent=current_parent,
                        level=current_parent.level + 1
                    )
                    current_parent.children.append(node)
                current_parent = current_parent.parent
        return root_node

    @property
    def is_valid(self):
        if self._is_valid is None:
            self._is_valid = False if self.errors else True
        return self._is_valid

    @property
    def net_elements(self):
        return [n for n in self.flat_tree if n.node_type == CodeplanNodeTypes.Net]

    def get_element(self, code):
        elements = [e for e in self._elements if e.code == code]
        if elements:
            return elements[0]
        else:
            raise ValueError(f"Element '{code}' not found")

    def print_tree(self):
        for node in self.tree.flat_children:
            print(f'{"    " * (node.level - 1)}{node.code} - {node.label}')

    def inject(self, xl_codeplan):

        if self.errors:
            errors = '\n'.join(self.errors)
            raise ValueError(f'''Errors in Codeplan '{self.name}': {errors}''')
        if xl_codeplan.errors:
            errors = '\n'.join(self.errors)
            raise ValueError(f'''Errors in Excel Codeplan '{xl_codeplan.name}': {errors}''')

        xl_elements = {e.code for e in xl_codeplan.elements}
        mdd_elements = {e.code for e in self.elements}
        missing_in_mdd = sorted(xl_elements - mdd_elements, key=sort_element)
        missing_in_xl = sorted(mdd_elements - xl_elements, key=sort_element)
        
        if missing_in_mdd:
            #raise ValueError(f"Excel elements missing in MDD: {','.join(missing_in_mdd)}")
            print(f'XL elements missing in MDD: {",".join(missing_in_mdd)}')
        if missing_in_xl:
            if xl_codeplan.other_element:
                print(f'MDD elements missing in Excel: {",".join(missing_in_xl)}')
                self.category_map = {e: xl_codeplan.other_element for e in missing_in_xl}
            else:
                raise ValueError(f"'Other element' not set in excel codeplan '{xl_codeplan.name}'")

        exist_in_both = sorted(mdd_elements & xl_elements, key=sort_element)
        for e in exist_in_both:
            mdd_element = self.get_element(e)
            xl_element = xl_codeplan.get_element(e)
            if mdd_element.label != xl_element.label:
                print(f'Overwriting label for {e}: "{mdd_element.label}" -> "{xl_element.label}"')
            
        self._elements = xl_codeplan.elements
        self._tree = xl_codeplan.tree
        self._flat_tree = xl_codeplan.flat_tree
        self._axis = xl_codeplan.axis

    def print_summary(self):
        error_string = '\n'.join(self.errors) if self.errors else '(not found)'
        
        print(f'''Name: {self.name}
# Codes: {len(self.elements)}
# Nets: {len(self.net_elements)}
Errors: {error_string}''')

    def __len__(self):
        return len(self._elements)

    def __getitem__(self, i):
        if isinstance(i, str):
            return [e for e in self._elements if e.code == i][0]
        else:
            return self._elements[i]

    def __repr__(self):
        return f"MDDCodeplan(name='{self.name}'), len={len(self)}"


class MDDVariable:

    def __init__(self, name, label, type_name, axis):
        self.name = name
        self.label = label
        self.type_name = type_name
        self.axis = axis

        self._field_name = None
        self._iterations = None
        self._compliant_name = None

    @property
    def field_name(self):
        '''f4l[{axa}].f4 -> f4l.f4'''
        if self._field_name is None:
            self._field_name = '.'.join(part.split(
                '[')[0] for part in self.label.split('.'))
        return self._field_name

    @property
    def iterations(self):
        '''q7loop[{_12}].q7[_5].slice -> [_12, _5]'''
        if self._iterations is None:
            self._iterations = [part.split(
                '[')[1][1:-2] for part in self.label.split('.') if len(part.split('[')) > 1]
        return self._iterations

    @property
    def compliant_name(self):
        '''f4l[{axa}].f4 -> f4l_f4_axa_o_c'''
        if self._compliant_name is None:
            prefix = '_'.join(self.field_name.split('.'))
            suffix = '_'.join(self.iterations) + '_o_c'
            self._compliant_name = f'{prefix}_{suffix}'
        return self._compliant_name

    def __repr__(self):
        return f"MDDVariable(name='{self.name}'), label='{self.label}', type_name='{self.type_name}'"



############################################################################
#
#                               EXCEL CODEPLAN
#
############################################################################

class XLCodeplan():

    def __init__(self, path, sheet_name, *, code_column=1):
        self.path = path
        self.name = sheet_name
        self.code_column = code_column
        worksheet = load_workbook(self.path, read_only=True, data_only=True)[sheet_name]
        all_rows = [
            XLCodeplanRow(
                code=row[self.code_column - 1].value,
                label=row[self.code_column].value,
                index=index
            )
            for index, row in enumerate(worksheet, start=1)
        ]
        self._rows = self._trim_rows(all_rows)
        self._other_element = None
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
    def other_element(self):
        return self._other_element
    
    @other_element.setter
    def other_element(self, value):
        self._other_element = value

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
                        label=max(labels),
                        double=True if len(labels) > 1 else False
                    )
                )
        self._elements.sort()
        return self._elements

    @property
    def tree(self):
        if self._tree is None:

            self._tree = CodeplanNode(
                code='(root)',
                label='',
                node_type=CodeplanNodeTypes.Root,
                parent=None,
                level=0
            )
            current_parent = self._tree

            for row in self._rows:
                if row.row_type == XLCodeplanRowTypes.NetStart:
                    node = CodeplanNode(
                        code=f'net{row.index}',
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
                        code=f'{CODE_PREFIX}{row.code}',
                        label=row.label,
                        node_type=CodeplanNodeTypes.Regular,
                        parent=current_parent,
                        level=current_parent.level + 1)
                    current_parent.children.append(node)
                elif row.row_type == XLCodeplanRowTypes.Combine:
                    combine_node = CodeplanNode(
                        code=f'comb{row.index}',
                        label=row.label,
                        node_type=CodeplanNodeTypes.Combine,
                        parent=current_parent,
                        level=current_parent.level + 1)
                    for c in row.combine_codes:
                        child = CodeplanNode(
                            code=f'{CODE_PREFIX}{c}',
                            label='',
                            node_type=CodeplanNodeTypes.Regular,
                            parent=combine_node,
                            level=combine_node.level + 1)
                        combine_node.children.append(child)
                    current_parent.children.append(combine_node)

        return self._tree

    @property
    def flat_tree(self):
        return self.tree.flat_children

    @property
    def axis(self):
        return self.tree.axis

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

    @property
    def double_elements(self):
        seen = set()
        duplicates = set()
        for node in self.flat_tree:
            if node.code not in seen:
                seen.add(node.code)
            else:
                duplicates.add(node.code)
        return sorted(duplicates, key=lambda x: int(x[len(CODE_PREFIX):]))

    @property
    def net_elements(self):
        return [n for n in self.flat_tree if n.node_type == CodeplanNodeTypes.Net]

    @property
    def combine_elements(self):
        return [n for n in self.flat_tree if n.node_type == CodeplanNodeTypes.Combine]

    def get_element(self, code):
        elements = [e for e in self._elements if e.code == code]
        if elements:
            return elements[0]
        else:
            raise ValueError(f"Element '{code}' not found")

    def print_tree(self):
        for node in self.flat_tree:
            print(f'{"    "*(node.level - 1)}{node.code} - {node.label}')

    def print_summary(self):
        error_string = '\n'.join(self.errors) if self.errors else '(not found)'
        double_elements = self.double_elements
        double_elements_string = ','.join(double_elements) if double_elements else '(not found)'
        
        print(f'''Name: {self.name}
# Codes: {len(self.elements)}
# Nets: {len(self.net_elements)}
# Combines: {len(self.combine_elements)}
Double codes: {double_elements_string}
Errors: {error_string}''')

    def __len__(self):
        return len(self.elements)

    def __getitem__(self, i):
        if isinstance(i, str):
            return [e for e in self._elements if e.code == i][0]
        else:
            return self._elements[i]

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


class CFileManager:

    def __init__(self, mdd_file, cfile_path):
        self.mdd_file = mdd_file
        self.cfile_path = cfile_path


    def save_cfile(self, new_path, *, cfile_source=CFileSources.Verbaco):

        variable_map = {old + HELPER_FIELD: new for old, new in self.mdd_file.variable_map.items()}
        category_map = {v.label + HELPER_FIELD: {**cp.category_map}
            for cp in self.mdd_file
            for v in self.mdd_file.variables
            if v.type_name == cp.name}

        with open(self.cfile_path, mode='r', encoding='utf-8') as input_file, \
        open(new_path, mode='w', encoding='utf-8') as output_file:
            for input_line in input_file:
                if cfile_source == CFileSources.Verbaco:
                    output_line = self._update_verbaco_line(input_line, variable_map, category_map)
                elif cfile_source == CFileSources.Ascribe:
                    output_line = self._update_ascribe_line(input_line, variable_map, category_map)
                output_file.write(output_line)

    def _update_verbaco_line(self, input_line, variable_map, category_map):
        sql_parts = input_line.split(' ')
        assignment = sql_parts[3]
        variable = assignment.split('=')[0]
        codes = assignment.split('=')[1][1:-1].split(',')
        new_variable = variable_map.get(variable, variable)
        variable_category_map = category_map.get(variable)
        new_codes = [variable_category_map.get(c, c) for c in codes] if variable_category_map else codes
        new_codes_without_duplicates = dict.fromkeys(new_codes)
        new_assignment = f"{new_variable}={{{','.join(new_codes_without_duplicates)}}}"
        sql_parts[3] = new_assignment
        return ' '.join(sql_parts)


    def _update_ascribe_line(self, input_line, variable_map, category_map):
        assignments_string, criteria = input_line[17:].split(' WHERE ')
        assignments = assignments_string.strip().split(', ')
        new_assignments = []
        for a in assignments:
            variable = a.split('=')[0].strip()
            codes = a.split('=')[1].strip()[1:-1].split(',')
            new_variable = variable_map.get(variable, variable)
            variable_category_map = category_map.get(variable)
            new_codes = [variable_category_map.get(c, c) for c in codes] if variable_category_map else codes
            new_codes_without_duplicates = dict.fromkeys(new_codes)
            new_assignment = f"{new_variable} = {{{','.join(new_codes_without_duplicates)}}}"
            new_assignments.append(new_assignment)
        return f"UPDATE vdata SET {', '.join(new_assignments)} WHERE {criteria}"


def update_master_mdd(master_path, verbaco_path, adapter):
    pass

def update_master_ddf(master_path, cfile_path):
    pass
