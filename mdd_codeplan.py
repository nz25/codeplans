from enum import IntEnum
from win32com import client
from collections import defaultdict, namedtuple
from codeplan import CodeplanNode, CodeplanElement, CodeplanNodeTypes


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

class AxisSeparators(IntEnum):
    Comma = 1
    NetStart = 2
    NetEnd = 3

class MDDCodeplans():

    def __init__(self, path):
        self.path = path
        self._read_mdd()
        self._multitype_variables = None
        self._errors = None
        self._is_valid = None

    def _read_mdd(self):
        mdd = client.Dispatch('MDM.Document')
        mdd.Open(self.path)
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
                variables=[
                    MDDVariable(
                        name=f.Label,
                        name_in_mdd=f.Name,
                        type_name=t.Name,
                        axis=f.AxisExpression
                    )
                    for f in mdd.Fields
                    if f.ObjectTypeValue == ObjectTypesConstants.mtVariable
                        and f.Elements.Reference.Name == t.Name
                ]
            ) for t in mdd.Types
        ]

        mdd.Close()

    @property
    def multitype_variables(self):
        if self._multitype_variables is None:

            # saves list of variables per field
            field_variables = defaultdict(list)
            for cp in self._codeplans:
                for v in cp.variables:
                    field_variables[v.field_name].append(v)

            # saves list of types per field
            field_types = defaultdict(set)
            for cp in self._codeplans:
                for v in cp.variables:
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

            self._multitype_variables = [
                v
                for f in fields
                if len(f.types) > 1
                for v in f.variables
            ]

        return self._multitype_variables

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
        return f"MDDCodeplans(path='{self.path}')"

class MDDCodeplan:

    def __init__(self, name, elements, variables):
        self.name = name
        self._elements = elements
        self.variables = variables
        self._errors = None
        self._tree = None
        self._is_valid = None
        self._axis = None

    @property
    def errors(self):
        if self._errors is None:
            self._errors = []
            if len({v.axis for v in self.variables}) > 1:
                self._errors.append(f'Axis expressions are not unique')
        return self._errors

    @property
    def tree(self):
        if self._tree is None:
            axis = self.axis[1:-1]
            label_bitmap = self._build_label_bitmap(axis)
            split_axis = self._split_axis(axis, label_bitmap)
            self._tree = self._build_tree(split_axis)

        return self._tree

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
                    split_axis.append((AxisSeparators.Comma,axis[last_character:i].strip()))
                    last_character = i + 1
                elif last_5_characters == 'net({':
                    split_axis.append((AxisSeparators.NetStart, axis[last_character:i-4].strip()))
                    last_character = i + 1
                elif last_2_characters == '})':
                    split_axis.append((AxisSeparators.NetEnd, axis[last_character:i-1].strip()))
                    last_character = i + 1

        if last_character < len(axis):
            split_axis.append(axis[last_character:])

        return split_axis

    def _build_tree(self, split_axis):

        root_node = CodeplanNode(
                name = self.name,
                label = '',
                node_type=CodeplanNodeTypes.Root,
                parent = None,
                level = 0
            )
        current_parent = root_node

        for a in split_axis:
            if a[1] == 'base()':
                node = CodeplanNode(
                    name='',
                    label='',
                    node_type=CodeplanNodeTypes.Base,
                    parent=current_parent,
                    level=current_parent.level
                )
                current_parent.children.append(node)
            elif a[0] == AxisSeparators.Comma and a[1]:
                #regular element
                name, label = a[1].split(sep=' ', maxsplit=1)
                node = CodeplanNode(
                    name=name.strip(),
                    label=label.strip()[1:-1].replace("''", "'"),
                    node_type=CodeplanNodeTypes.Regular,
                    parent=current_parent,
                    level=current_parent.level
                )
                current_parent.children.append(node)
            elif a[0] == AxisSeparators.NetStart:
                #net element
                name, label = a[1].split(sep=' ', maxsplit=1)
                node = CodeplanNode(
                    name=name.strip(),
                    label=label.strip()[1:-1].replace("''", "'"),
                    node_type=CodeplanNodeTypes.Net,
                    parent=current_parent,
                    level=current_parent.level + 1
                )
                current_parent.children.append(node)
                current_parent = node
            elif a[0] == AxisSeparators.NetEnd:
                if a[1]:
                    name, label = a[1].split(sep=' ', maxsplit=1)
                    node = CodeplanNode(
                        name=name.strip(),
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
    def axis(self):
        if self._axis is None:
            self._axis = {v.axis for v in self.variables}.pop()
        return self._axis

    def __len__(self):
        return len(self._elements)

    def __getitem__(self, i):
        return self._elements[i]

    def __repr__(self):
        return f"MDDCodeplan(name='{self.name}'), len={len(self)}"

class MDDVariable:

    def __init__(self, name, name_in_mdd, type_name, axis):
        self.name = name
        self.name_in_mdd = name_in_mdd
        self.type_name = type_name
        self.axis = axis

        self._field_name = None
        self._iterations = None
        self._compliant_name = None

    @property
    def field_name(self):
        '''f4l[{axa}].f4 -> f4l.f4'''
        if self._field_name is None:
            self._field_name = '.'.join(part.split('[')[0] for part in self.name.split('.'))
        return self._field_name

    @property
    def iterations(self):
        '''q7loop[{_12}].q7[_5].slice -> [_12, _5]'''
        if self._iterations is None:
            self._iterations = [part.split('[')[1][1:-2] for part in self.name.split('.') if len(part.split('[')) > 1]
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
        return f"MDDVariable(name='{self.name}'), name_in_mdd='{self.name_in_mdd}', type_name='{self.type_name}'"

def test():
    cp = MDDCodeplans('codeplan_1533204901432_2018-08-02.mdd')
    for c in cp:
        print(c.tree)
    for v in cp.multitype_variables:
        print(v.name, v.compliant_name)

if __name__ == '__main__':
    from timeit import timeit

    print(timeit(test,number=1))
    print('OK')