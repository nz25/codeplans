from enum import IntEnum
from settings import CODE_PREFIX

class CodeplanNodeTypes(IntEnum):
    Root = 0
    Regular = 1
    Combine = 2
    Net = 3
    Base = 4

class CodeplanSources(IntEnum):
    XL = 1
    MDD = 2
    Master = 3

class Codeplan:

    def __init__(self,
    mdd_codeplan,
    xl_codeplan=None,
    mdd_xl_adapter=None,
    element_source=CodeplanSources.MDD,
    label_source=CodeplanSources.MDD,
    axis_source=CodeplanSources.XL):
        self.xl_codeplan = xl_codeplan
        self.mdd_codeplan = mdd_codeplan
        self.xl_mdd_adapter = mdd_xl_adapter
        self.element_source = element_source
        self.label_source = label_source
        self.axis_source = axis_source

    def __repr__(self):
        return f'Codeplan(mdd_codeplan={self.mdd_codeplan}, xl_codeplan = {self.xl_codeplan}, xl_mdd_adapter={self.xl_mdd_adapter},element_source={self.element_source}, label_source={self.label_source}, axis_source={self.axis_source})'

class CodeplanNode:

    def __init__(self, name, label, node_type, parent, level):
        self.name = name
        self.label = label
        self.node_type = node_type
        self.parent = parent
        self.level = level
        self.children = []
        self._axis = None

    @property
    def axis(self):
        if self._axis is None:
            label = self.label.replace(r"'", r"''")
            if self.node_type == CodeplanNodeTypes.Root:
                self._axis = f'{{{",".join(c.axis for c in self.children)}}}'
            elif self.node_type == CodeplanNodeTypes.Base:
                self._axis = 'base()'
            elif self.node_type == CodeplanNodeTypes.Net:
                self._axis = f'{self.name} \'{label}\' net({{{",".join(c.axis for c in self.children)}}})'
            elif self.node_type == CodeplanNodeTypes.Combine:
                self._axis = f'{self.name} \'{label}\' combine({{{",".join(c.axis for c in self.children)}}})'
            elif self.node_type == CodeplanNodeTypes.Regular:
                self._axis = f'{self.name}'
        return self._axis    

    def __repr__(self):
        return f"CodeplanNode(name='{self.name}', label='{self.label}', node_type={self.node_type}, parent={self.parent}, level={self.level}, children={len(self.children)})"

class CodeplanElement:

    def __init__(self, code, label, doubled=False):
        self.code = code
        self.label = label
        self.doubled = doubled

    def __gt__(self, other):
        return int(self.code[len(CODE_PREFIX):]) > int(other.code[len(CODE_PREFIX):])

    def __repr__(self):
        return f"CodeplanElement(code='{self.code}', label='{self.label}', doubled={self.doubled})"


def test():
    pass



if __name__ == '__main__':
    from timeit import timeit
    print(timeit(test, number=1))
    print('OK')

