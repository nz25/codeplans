# mdd_codeplans.py

from enums import ObjectTypesConstants
from win32com import client
from os.path import exists
from collections import defaultdict, namedtuple

class MDDCodeplans():

    def __init__(self, path):
        self.path = path
        self._read_mdd()
        self._populate()

    def _read_mdd(self):
        mdd = client.Dispatch('MDM.Document')
        mdd.Open(self.path)
        self.types = {t.Name: {e.Name: e.Label for e in t.Elements} for t in mdd.Types}
        self.variables = {f.Label :
            MDDCodeplanVariable(
                name=f.Label,
                name_in_mdd=f.Name,            
                type_name=f.Elements.Reference.Name,
                axis_expression=f.AxisExpression)
            for f in mdd.Fields
            if f.ObjectTypeValue == ObjectTypesConstants.mtVariable}
        mdd.Close()

    def _populate(self):

        # saves list of variables per field
        field_variables = defaultdict(list)
        for v in self.variables.values():
            field_variables[v.field_name].append(v)

        # saves list of types per field
        field_types = defaultdict(set)
        for v in self.variables.values():
            field_types[v.field_name].add(v.type_name) 

        # merges both results in temporary fields list
        Field = namedtuple('Field', 'name variables types')
        fields = [Field(field_name, variables, field_types[field_name]) for field_name, variables in field_variables.items()]

        # builds collection of helper fields
        # which are required in project mdd
        self.helper_fields = {f'{f.name}.Coding':
            MDDHelperField(
                name='Coding',
                type_name=next(iter(f.types)),
                axis_expression=f.variables[0].axis_expression,
                parent_field_name=f.name)
            for f in fields
            if len(f.types) == 1}

        # following fields have to be created
        # because they don't share common type
        self.new_fields = {v.compliant_name:
            MDDField(
                name=v.compliant_name,
                type_name=v.type_name,
                axis_expression=v.axis_expression)
            for f in fields
                for v in f.variables
            if len(f.types) > 1}    

        ## variable_map dictionary is used for
        ## renaming variables in sql text file
        changed_variables = {
            v.name: v.compliant_name
            for f in fields
                for v in f.variables
            if len(f.types) > 1}
        
        unchanged_variables = {
            f'{v}.Coding': f'{v}.Coding'
            for v in self.variables
            if v not in changed_variables}

        self.variable_map = {**changed_variables, **unchanged_variables}

    def __repr__(self):
        return f'MDDCodeplans(T={len(self.types)}, V={len(self.variables)}, HF={len(self.helper_fields)}, NF={len(self.new_fields)})'

class MDDCodeplanVariable():

    def __init__(self, name, name_in_mdd, type_name, axis_expression):
        self.name = name
        self.name_in_mdd = name_in_mdd
        self.type_name = type_name
        self.axis_expression = axis_expression
        self.field_name = self._get_field_name()
        self.iterations = self._get_iterations()
        self.compliant_name = self._get_compliant_name()

    def _get_field_name(self):
        '''f4l[{axa}].f4 -> f4l.f4'''
        return '.'.join(part.split('[')[0] for part in self.name.split('.'))

    def _get_iterations(self):
        '''q7loop[{_12}].q7[_5].slice -> [_12, _5]'''
        return [part.split('[')[1][1:-2] for part in self.name.split('.') if len(part.split('[')) > 1]

    def _get_compliant_name(self):
        '''f4l[{axa}].f4 -> f4l_f4_axa_o_c'''
        prefix = '_'.join(self.field_name.split('.'))
        suffix = '_'.join(self.iterations) + '_o_c'
        return f'{prefix}_{suffix}'

    def __repr__(self):
        return f'MDDCodeplanVariable(name={self.name})'

class MDDField():

    def __init__(self, name, type_name, axis_expression):
        self.name = name
        self.type_name = type_name
        self.axis_expression = axis_expression

    def __repr__(self):
        return f'MDDField(name={self.name})'

class MDDHelperField(MDDField):

    def __init__(self, name, type_name, axis_expression, parent_field_name):
        super().__init__(name, type_name, axis_expression)
        self.parent_field_name = parent_field_name

    def __repr__(self):
        return f'MDDHelperField(name={self.parent_field_name}.{self.name})'

if __name__ == '__main__':
    cp = MDDCodeplans('codeplan_1531899367053_2018-07-18.mdd')
    print('ok')