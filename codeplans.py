from enum import IntEnum
from win32com import client
from openpyxl import load_workbook
from collections import defaultdict, namedtuple
from shutil import copyfile
from adodbapi import connect
from xml.etree import ElementTree

############################################################################
#
#                         CONSTANTS / ENUMERATIONS
#
############################################################################

CODE_PREFIX = 'CB_'
HELPER_FIELD =  '.Coding'

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

class MDDFile:

    def __init__(self, mdd_path, *, parser='xml'):

        self.path = mdd_path
        self.parser = parser

        # reads meta data from mdd
        if parser == 'xml':
            self.types, self.variables = self._read_mdd_from_xml() 
        elif parser == 'com':
           self.types, self.variables = self._read_mdd_from_com() 
 

    def _read_mdd_from_com(self):

        # parser, which uses MDM.Document COM Object
        # to access types and variables

        # opens mdd
        mdd = client.Dispatch('MDM.Document')
        mdd.Open(self.path, mode=openConstants.oREAD)

        # fills types
        types = [
            MDDCodeplan(
                name=t.Name,
                elements=[
                    CodeplanElement(code=e.Name, label=e.label)
                    for e in t.Elements
                ],
                axis={f.AxisExpression
                    for f in mdd.Fields
                    if f.ObjectTypeValue == ObjectTypesConstants.mtVariable
                        and f.Elements.Reference.Name == t.Name
                }.pop(),
                mdd_file=self
            ) for t in mdd.Types]

        # fills variables
        variables = [
            MDDVariable(
                name=f.Name,
                label=f.Label,
                type_name=f.Elements.Reference.Name,
                axis=f.AxisExpression
            )
            for f in mdd.Fields
            if f.ObjectTypeValue == ObjectTypesConstants.mtVariable
        ]

        # closes mdd
        mdd.Close()

        return types, variables


    def _read_mdd_from_xml(self):

        # parser, which uses XML from MDD file
        # to access types and variables

        types = []
        variables = []
        
        # root node for variables and types is xml\mdm:metadata\definition
        tree = ElementTree.parse(self.path)
        root = tree.getroot()[0].find('definition')

        # root node contains elements of 2 types:
        # 'variable' for variables
        # 'categories' for types
        for node in root:
            # parsing logic for variables
            if node.tag == 'variable':
                name = node.get('name')
                var_type = int(node.get('type'))
                # labels element may contain multiple labels 
                # for different label types, contexts and languages (LCL)
                # parser just uses 1st label it encounters
                label = node.find('labels')[0].text
                ref_name = node.find('categories').get('ref_name')
                axis = node.find('axis').get('expression')
                if var_type == DataTypeConstants.mtCategorical and ref_name:
                    variables.append(MDDVariable(name, label, ref_name, axis))
            # parsing logic for types
            elif node.tag == 'categories':
                name = node.get('name')
                elements = [
                    CodeplanElement(
                        code=element.get('name'),
                        label=element.find('labels')[0].text
                    )
                    for element in node
                    if element.tag == 'category'
                ]
                # sets axis expression for type based on the 1st variable
                # which belongs to this type
                # assumes variables collection is completely populated
                axis = {v.axis for v in variables if v.type_name == name}.pop()
                types.append(MDDCodeplan(name, elements, axis, self))

        return types, variables

    @property
    def variable_map(self):

        # populates field list from variable list in 3 stages:
        # - saves list of variables per field
        # - saves list of types per field
        # - merges both results in fields list

        variables_per_field = defaultdict(list)
        for v in self.variables:
            variables_per_field[v.field_name].append(v)

        types_per_field = defaultdict(set)
        for v in self.variables:
            types_per_field[v.field_name].add(v.type_name)

        Field = namedtuple('Field', 'name variables types')
        fields = [
            Field(
                name=k,
                variables=v,
                types=types_per_field[k]
            )
            for k, v in variables_per_field.items()
        ]

        # checks if there are variables, which belong to the same field
        # but use different types.
        # creates variable map, for renaming such variables in cfile

        return {
            f'{v.label}{HELPER_FIELD}': v.compliant_name
                for f in fields
                for v in f.variables
                if len(f.types) > 1
        }

    def save_variable_map(self, path):
        with open(path, mode='w', encoding='utf-8') as f:
            for vm in self.variable_map.items():
                f.write(','.join(vm) + '\n')

    def save_as(self, path):

        self.path = path
        
        # saves types and elements lists in mdd file
        
        mdd = client.Dispatch('MDM.Document')
        mdd.IncludeSystemVariables = False
        for t in self.types:
            new_list = mdd.CreateElements(t.name)
            for element in t:
                new_element = mdd.CreateElement(element.code, element.label)
                new_element.Type = ElementTypeConstants.mtCategory
                new_list.Add(new_element)
            mdd.Types.Add(new_list)
            
        for v in self.variables:
            new_variable = mdd.CreateVariable(v.name, v.label)
            new_variable.DataType = DataTypeConstants.mtCategorical
            new_variable.Elements.ReferenceName = v.type_name
            new_variable.AxisExpression = self[v.type_name].axis
            mdd.Fields.Add(new_variable)

        mdd.CategoryMap.AutoAssignValues()
        mdd.Save(path)
        mdd.Close()

    def __getitem__(self, value):
        if isinstance(value, str):
            return [t for t in self.types if t.name == value][0]
        else:
            return self.types[value]        

    def __contains__(self, value):
        return bool([t for t in self.types if t.name == value])


    def __repr__(self):
        return f"MDDFile(path='{self.path}')"

class MDDCodeplan:

    def __init__(self, name, elements, axis, mdd_file):
        self.name = name
        self.mdd_file = mdd_file
        self.elements = elements
        self.axis = axis
        self._errors = None
        self._tree = None
        self._is_valid = None

    @property
    def errors(self):
        if self._errors is None:
            self._errors = []
        return self._errors

    @property
    def tree(self):
        if self._tree is None and self.axis:
            self.axis = self.axis[1:-1]
            label_bitmap = self._build_label_bitmap(self.axis)
            split_axis = self._split_axis(self.axis, label_bitmap)
            self._tree = self._build_tree(split_axis)

        return self._tree

    @tree.setter
    def tree(self, value):
        self._tree = value

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
        return [n for n in self.tree.flat_children if n.node_type == CodeplanNodeTypes.Net]

    @property
    def variables(self):
        return [v for v in self.mdd_file.variables if v.type_name == self.name]

    def print_tree(self):
        for node in self.tree.flat_children:
            print(f'{"    " * (node.level - 1)}{node.code} - {node.label}')

    def print_summary(self):
        error_string = '\n'.join(self.errors) if self.errors else '(not found)'
        print(f'''Name: {self.name}\n# Codes: {len(self.elements)}\n# Nets: {len(self.net_elements)}\nErrors: {error_string}''')

    def __len__(self):
        return len(self.elements)

    def __getitem__(self, i):
        if isinstance(i, str):
            return [e for e in self.elements if e.code == i][0]
        else:
            return self.elements[i]

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

class XLFile:

    def __init__(self, path, *, code_column=1):
        
        self.path = path
        self.code_column = code_column
        workbook = load_workbook(self.path, read_only=True, data_only=True)
        self.codeplans = [
            XLCodeplan(
                name=sheet.title,
                rows = [
                    XLCodeplanRow(
                        code=row[self.code_column - 1].value,
                        label=row[self.code_column].value,
                        index=index
                    )
                    for index, row in enumerate(sheet, start=1)],
                xl_file = self)
            for sheet in workbook
        ]
        self.category_map = []
        workbook.close()

    def __getitem__(self, value):
        if isinstance(value, str):
            return [t for t in self.codeplans if t.name == value][0]
        else:
            return self.codeplans[value]        

    def __contains__(self, value):
        return bool([t for t in self.codeplans if t.name == value])

    def __repr__(self):
        return f'XLFile(path="{self.path}")'

class XLCodeplan:

    def __init__(self, name, rows, xl_file, *, other_element=''):
        self.name =name
        self.rows = self._trim_rows(rows)
        self.other_element = other_element
        self.xl_file = xl_file
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
            for row in self.rows:
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

            for row in self.rows:
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
    def axis(self):
        return self.tree.axis


    @property
    def errors(self):

        if self._errors is None:

            self._errors = []

            # Checks if codeplan is empty
            if len(self.rows) == 0:
                self._errors.append('Empty codeplan')

            # Check if codes are valid
            self._errors.extend([f'Invalid code "{row.code}" in row {row.index}'
                                 for row in self.rows if not row.is_valid])

            # Structural validation
            current_level = 0
            current_elements = []
            last_row_type = XLCodeplanRowTypes.Invalid
            for row in self.rows:
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
        for node in self.tree.flat_children:
            if node.code not in seen:
                seen.add(node.code)
            else:
                duplicates.add(node.code)
        return sorted(duplicates, key=lambda x: int(x[len(CODE_PREFIX):]))

    @property
    def net_elements(self):
        return [n for n in self.tree.flat_children if n.node_type == CodeplanNodeTypes.Net]

    @property
    def combine_elements(self):
        return [n for n in self.tree.flat_children if n.node_type == CodeplanNodeTypes.Combine]

    def print_tree(self):
        for node in self.tree.flat_children:
            print(f'{"    "*(node.level - 1)}{node.code} - {node.label}')

    def print_summary(self):
        error_string = '\n'.join(self.errors) if self.errors else '(not found)'
        double_elements = self.double_elements
        double_elements_string = ','.join(double_elements) if double_elements else '(not found)'
        
        print(f'''Name: {self.name}\n# Codes: {len(self.elements)}\n# Nets: {len(self.net_elements)}\n# Combines: {len(self.combine_elements)}\nDouble codes: {double_elements_string}\nErrors: {error_string}''')

    def __len__(self):
        return len(self.elements)

    def __getitem__(self, i):
        if isinstance(i, str):
            return [e for e in self.elements if e.code == i][0]
        else:
            return self.elements[i]

    def __repr__(self):
        return f"XLCodeplan(name='{self.name}'), len={len(self)}"

class XLCodeplanRow:

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


############################################################################
#
#                                MERGERS
#
############################################################################

class MDDXLFileMerger:

    def __init__(self, mdd_file, xl_file, mdd_xl_map, *, verbose = False):

        # mdd_xl_map is named_tuple with 3 fields:
        # mdd_name xl_name other_element

        self.mdd_file = mdd_file
        self.xl_file = xl_file
        self.mdd_xl_map = mdd_xl_map
        self.codeplan_mergers = []

        if verbose:
            print('Initializing MDDXLFileMerger...')
            mdd_codeplans_missing_in_map = {t.name for t in self.mdd_file.types if t.name not in {m.mdd_name for m in self.mdd_xl_map}}
            print('MDD types missing in the map:')
            print(','.join(mdd_codeplans_missing_in_map))
            xl_codeplans_missing_in_map = {cp.name for cp in self.xl_file.codeplans if cp.name not in {m.xl_name for m in self.mdd_xl_map}}
            print('XL types missing in the map:')
            print(','.join(xl_codeplans_missing_in_map))

        for m in mdd_xl_map:
            if m.mdd_name and m.xl_name and m.mdd_name in mdd_file and m.xl_name in xl_file:
                mdd_codeplan = mdd_file[m.mdd_name]
                xl_codeplan = xl_file[m.xl_name]
                xl_codeplan.other_element = m.other_element
                self.codeplan_mergers.append(
                    CodeplanMerger(mdd_codeplan, xl_codeplan, self)
                )
        self.category_map = []

    def merge_all(self):
        adjusted_types = []
        for m in self.codeplan_mergers:
            adjusted_types.append(m.merge())
            self.category_map.extend(m.category_map)
        return self.mdd_file

    def save_category_map(self, path):
        with open(path, mode='w', encoding='utf-8') as f:
            for cm in self.category_map:
                f.write(','.join(cm) + '\n')

class MDDFileMerger:

    # merging rules are as follows:
    # - doesn't change types in master mdd
    # - appends types from slave mdd, which don't exist in master
    # - doesn't change variables in master mdd
    # - appends new varaibles from slave mdd
    # - uses axis expression from master for newly added variables
    #   if they use list, which existed in master

    def __init__(self, master_mdd_file, slave_mdd_file):

        # expects 2 MDDFile types parameters
        self.master = master_mdd_file
        self.slave = slave_mdd_file

        
        # types which missing in master
        self.new_types = [t for t in self.slave.types
            if t.name not in [mt.name for mt in self.master.types]
        ]

        # variables which missing in master
        self.new_variables = [v for v in self.slave.variables
            if v.name not in [mv.name for mv in self.master.variables]
        ]

        # adjusts axis expressions in new variables
        for v in self.new_variables:
            if v.type_name in [t.name for t in self.master.types]:
                v.axis = self.master[v.type_name].axis


    @property
    def report(self):
        
        report = ''

        # outputs new types
        report += 'New types:\n'
        for t in self.new_types:
            report += f'{t}\n'

        # outputs new variables
        report += 'New variables:\n'
        for v in self.new_variables:
            report += f'{v}\n'

        # returns report
        return report


    def merge(self):

        # adjusts and returns master
        self.master.types.extend(self.new_types)
        self.master.variables.extend(self.new_variables)
        return self.master

class CodeplanMerger:

    # merging rules are as follows:
    # - excel defines set of elements
    # - excel defines labels for elements
    # - excel defines codeplan structure (axis)
    # - if there are elements in mdd, which are missing in 
    #   excel, category_map is created, which maps, all
    #   missing element to other_element

    def __init__(self, mdd_codeplan, xl_codeplan, file_merger):

        # expects MDDCodeplan and XLCodeplan types parameters
        self.mdd_codeplan = mdd_codeplan
        self.xl_codeplan = xl_codeplan
        self.file_merger = file_merger
        self.other_element = xl_codeplan.other_element

        # comparing mdd and xl elements
        self.xl_elements = {e.code for e in xl_codeplan.elements}
        self.mdd_elements = {e.code for e in mdd_codeplan.elements}
        self.missing_in_mdd = sorted(self.xl_elements - self.mdd_elements, key=sort_element)
        self.missing_in_xl = sorted(self.mdd_elements - self.xl_elements, key=sort_element)
        self.exist_in_both = sorted(self.mdd_elements & self.xl_elements, key=sort_element)

        # check if there are conditions which prohibit merging
        if self.mdd_codeplan.errors or xl_codeplan.errors:
            self.mergeable = False
        elif self.missing_in_mdd or (self.missing_in_xl and not self.other_element):
            self.mergeable = False
        else:
            self.mergeable = True

    @property
    def report(self):

        report = ''
    
        # check for errors in mdd codeplan and xl codeplan
        mdd_errors = '\n'.join(self.mdd_codeplan.errors)
        xl_errors = '\n'.join(self.xl_codeplan.errors)
        if mdd_errors:
            report += f'Errors in MDD Codeplan "{self.mdd_codeplan.name}": {mdd_errors}\n'
        if xl_errors:
            report += f'Errors in XL Codeplan "{self.xl_codeplan.name}": {xl_errors}\n'

        # check for errors in mdd codeplan and xl codeplan
        if self.missing_in_mdd:
            report += f'XL elements missing in MDD: {",".join(self.missing_in_mdd)}\n'
        if self.missing_in_xl:
            if self.other_element:
                report += f'MDD elements missing in Excel: {",".join(self.missing_in_xl)}\n'
            else:
                report += f'"Other element" not set in excel codeplan {self.xl_codeplan.name}\n'

        # checks for differences in labels
        for e in self.exist_in_both:
            mdd_element = self.mdd_codeplan[e]
            xl_element = self.xl_codeplan[e]
            if mdd_element.label != xl_element.label:
                report += f'Label differences for {e}: "{mdd_element.label} (MDD)" -> "{xl_element.label}" (XL)\n'

        # returns report
        return report

    @property
    def category_map(self):

        return [
            (f'{v.label}{HELPER_FIELD}',
                old_code,
                self.other_element)
            for v in self.mdd_codeplan.variables
            for old_code in self.missing_in_xl
        ]
    
    def merge(self):

        # merges codeplans and returns mdd codeplan

        if self.mergeable:
            self.mdd_codeplan.elements = self.xl_codeplan.elements
            self.mdd_codeplan.tree = self.xl_codeplan.tree
            self.mdd_codeplan.axis = self.xl_codeplan.axis
            return self.mdd_codeplan

        else:
            raise ValueError('Not mergeable. See report() for details')


############################################################################
#
#                                 CFILE
#
############################################################################

class CFileManager:

    def __init__(self, cfile_path, variable_map_path, category_map_path, *, cfile_source=CFileSources.Verbaco):
        self.cfile_path = cfile_path

        self.variable_map = {}
        with open(variable_map_path, mode='r', encoding='utf-8') as f:
            for row in f:
                old_var, new_var = row.strip('\n').split(',')
                self.variable_map[old_var] = new_var

        self.category_map = defaultdict(dict)
        with open(category_map_path, mode='r', encoding='utf-8') as f:
            for row in f:
                variable, old_code, new_code = row.strip('\n').split(',')
                self.category_map[variable][old_code] = new_code

        self.cfile_source = cfile_source

    def save_cfile(self, new_path):

        updater = self._update_verbaco_line if self.cfile_source == CFileSources.Verbaco else self._update_ascribe_line
        with open(self.cfile_path, mode='r', encoding='utf-8') as input_file, \
        open(new_path, mode='w', encoding='utf-8') as output_file:
            for input_line in input_file:
                output_line = updater(input_line)
                output_file.write(output_line)

    def _update_verbaco_line(self, input_line):
        sql_parts = input_line.split(' ')
        assignment = sql_parts[3]
        variable = assignment.split('=')[0]
        codes = assignment.split('=')[1][1:-1].split(',')
        new_variable = self.variable_map.get(variable, variable)
        variable_category_map = self.category_map.get(variable)
        new_codes = [variable_category_map.get(c, c) for c in codes] if variable_category_map else codes
        new_codes_without_duplicates = dict.fromkeys(new_codes)
        new_assignment = f"{new_variable}={{{','.join(new_codes_without_duplicates)}}}"
        sql_parts[3] = new_assignment
        return ' '.join(sql_parts)

    def _update_ascribe_line(self, input_line):
        assignments_string, criteria = input_line[17:].split(' WHERE ')
        assignments = assignments_string.strip().split(', ')
        new_assignments = []
        for a in assignments:
            variable = a.split('=')[0].strip()
            codes = a.split('=')[1].strip()[1:-1].split(',')
            new_variable = self.variable_map.get(variable, variable)
            variable_category_map = self.category_map.get(variable)
            new_codes = [variable_category_map.get(c, c) for c in codes] if variable_category_map else codes
            new_codes_without_duplicates = dict.fromkeys(new_codes)
            new_assignment = f"{new_variable} = {{{','.join(new_codes_without_duplicates)}}}"
            new_assignments.append(new_assignment)
        return f"UPDATE vdata SET {', '.join(new_assignments)} WHERE {criteria}"


############################################################################
#
#                               UTILITIES
#
############################################################################

def sort_element(code):
    return int(code[len(CODE_PREFIX):])

def copy_mdd_ddf(input_path, output_path):
    
    # path should include complete directory path and file name without extention
    # e.g. 'C:\Folder\file' for file.mdd in C:\Folder folder

    from os.path import basename

    copyfile(f'{input_path}.mdd', f'{output_path}.mdd')
    copyfile(f'{input_path}.ddf', f'{output_path}.ddf')

    mdd = client.Dispatch('MDM.Document')
    mdd.Open(f'{output_path}.mdd')
    mdd.DataSources.Default.DBLocation = basename(f'{output_path}.ddf')
    mdd.Save()
    mdd.Close()
        
def update_cfile(cfile_path, variable_map, category_map, new_path):
    cfile_manager = CFileManager(cfile_path, variable_map, category_map)
    cfile_manager.save_cfile(new_path)

def update_master_with_mdd_codeplan_with_adapter(master_path, codeplan_path, adapter):

    master_mdd = client.Dispatch('MDM.Document')
    master_mdd.Open(master_path)
    codeplan_file = MDDFile(codeplan_path)

    # checks if all mdd types exist in adapter
    for cp in codeplan_file:
        if cp.name not in [m.mdd_name for m in adapter]:
            print(f"WARNING: {cp.name} doesn't exist in adapter")

    # update types
    for m in adapter:
        if m.mdd_name and m.master_name and m.mdd_name in codeplan_file:
            print(f'Working on Codeplan {m.master_name}')
            codeplan_mdd = codeplan_file[m.mdd_name]
            # creates type if it doesn't exist in the master
            if not master_mdd.Types.Exist(m.master_name):
                create_type(master_mdd, m.master_name, codeplan_mdd)
            master_elements = {e.Name.upper() for e in master_mdd.Types[m.master_name].Elements}
            mdd_elements = {e.code for e in codeplan_mdd.elements}
            
            # adds new elements
            missing_in_master = sorted(mdd_elements - master_elements, key=sort_element)
            for e in missing_in_master:
                print(f'Adding element {e}')
                new_element = master_mdd.CreateElement(e, codeplan_mdd[e].label)
                new_element.Type = ElementTypeConstants.mtCategory
                master_mdd.Types[m.master_name].Add(new_element)

            # updates labels
            exist_in_both = sorted(mdd_elements & master_elements, key=sort_element)
            for e in exist_in_both:
                codeplan_element = codeplan_mdd[e]
                master_element = master_mdd.Types[m.master_name].Elements[e]
                if codeplan_element.label != master_element.Label:
                    print(f'Overwriting label for {e}: "{master_element.Label}" -> "{codeplan_element.label}"')
                    master_element.Label = codeplan_element.label

    # update fields
    original_variables = [v for v in codeplan_file.variables if v.label + HELPER_FIELD not in codeplan_file.variable_map]
    new_variables =  [v for v in codeplan_file.variables if v.label + HELPER_FIELD in codeplan_file.variable_map]

    # creates .Coding variable if doesn't exist
    for v in original_variables:
        if not master_mdd.Fields.Expanded.Exist(v.field_name + HELPER_FIELD):
            master_type = [m.master_name for m in adapter if m.mdd_name == v.type_name][0]
            create_variable(
                mdd = master_mdd,
                parent_collection = master_mdd.Fields[v.field_name].HelperFields,
                var_name = HELPER_FIELD[1:],
                type_name = master_type,
                axis = v.axis
            )

    # checks and creates normal variables if they don't exist
    for v in new_variables:
        new_variable_name = codeplan_file.variable_map[v.label + HELPER_FIELD]
        if not master_mdd.Fields.Expanded.Exist(new_variable_name):
             master_type = [m.master_name for m in adapter if m.mdd_name == v.type_name][0]
             create_variable(
                mdd = master_mdd,
                parent_collection = master_mdd.Fields,
                var_name = new_variable_name,
                type_name = master_type,
                axis = v.axis
            )

    # # updates axis expressions
    for m in adapter:
        if m.mdd_name and m.master_name and m.mdd_name in codeplan_file:
            for f in master_mdd.Fields.Expanded:
                if f.ObjectTypeValue == ObjectTypesConstants.mtVariable and (f.Elements.ReferenceName == m.master_name or (f.Elements.IsReference and f.Elements.Reference.Name == m.master_name) or (f.Elements.Count > 0 and f.Elements[0].ReferenceName == m.master_name)):
                    f.AxisExpression = codeplan_file[m.mdd_name].axis

    master_mdd.CategoryMap.AutoAssignValues()
    master_mdd.Save()
    master_mdd.Close()

def create_type(mdd, name, mdd_codeplan):
    mdd_type = mdd.CreateElements(name)
    for e in mdd_codeplan.elements:
        element = mdd.CreateElement(e.code, e.label)
        element.Type = ElementTypeConstants.mtCategory
        mdd_type.Add(element)
    mdd.Types.Add(mdd_type)


def create_variable(mdd, parent_collection, var_name, type_name, axis):
    new_variable = mdd.CreateVariable(var_name)
    new_variable.DataType = DataTypeConstants.mtCategorical
    new_variable.Elements.ReferenceName = type_name
    new_variable.AxisExpression = axis
    parent_collection.Add(new_variable)

def execute_opens(connection, cfile_path):
    
    ddf = connect(connection).cursor()
    ddf.execute('exec xp_syncdb')
    with open(cfile_path, mode='r', encoding='utf-8') as sql_file:
        line_number = 0
        for sql_line in sql_file:
            ddf.execute(sql_line)
            line_number += 1
            if line_number % 100 == 0:
                print(f'{line_number} rows executed')
    ddf.connection.commit()
    ddf.close()
