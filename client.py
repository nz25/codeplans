from mdd_codeplans import MDDCodeplans
from xl_codeplan import XLCodeplans

def main():
    cp = MDDCodeplans('codeplan_1531899367053_2018-07-18.mdd')
    xl = XLCodeplans('Codeplan KTV online 201807.xlsx')
    print('ok')




def _build_label_bitmap(axis):
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

if __name__ == '__main__':
    x = _build_label_bitmap("n '''hel''lob' x")
    print(x)



        # self.errors = []
        # self.element_list = []
        # self.element_tree = CodeplanElement('', '', CodeplanElements.Root, None, 0)
        # self.axis_expression = ''

        # self._validate_rows()
        # if self.errors:
        #     raise ValueError(f'Error in codeplan {self.name}: /n {self.errors}')

        # self.element_tree = CodeplanElement('', '', CodeplanElements.Root, None, 0)
        # self.flat_elements = []
        # self._populate_elements()

        # self._validate_elements()
        # if self.errors:
        #     raise ValueError(f'Error in codeplan {self.name}: /n {self.errors}')

        # self.axis_expression = self.element_tree._get_axis()
        # print(self.name)
        # print(self.axis_expression)
        # print('')




#     def _populate_elements(self):
#         current_level = 0
#         current_parent = self.element_tree

#         base_element = CodeplanElement('', '', CodeplanElements.Base, current_parent, current_level)
#         self.element_tree.children.append(base_element)

#         for entry in self.rows:
#             if entry.entry_type == CodeplanRows.NetStart:
#                 if len(entry.code) == current_level + 1:
#                     current_level += 1
#                     element = CodeplanElement(f'net{entry.index}', entry.label, CodeplanElements.Net, current_parent, current_level)
#                     current_parent.children.append(element)
#                     self.flat_elements.append(element)
#                     current_parent = element
#                 else:
#                     self.errors.append(f'Invalid opening tag at row {entry.index}')
#                     break                    
#             elif entry.entry_type == CodeplanRows.NetEnd:
#                 current_level = len(entry.code)
#                 level_difference = current_parent.level - current_level + 1
#                 if level_difference > 0:
#                     while level_difference:
#                         current_parent = current_parent.parent
#                         level_difference -= 1
#                     current_level = current_parent.level
#                 else:
#                     self.errors.append(f'Invalid closing tag at row {entry.index}')
#                     break
#             elif entry.entry_type == CodeplanRows.Regular:
#                 element = CodeplanElement(f'CB_{entry.code}', entry.label, CodeplanElements.Regular, current_parent, current_level)
#                 current_parent.children.append(element)
#                 self.flat_elements.append(element)
#             elif entry.entry_type == CodeplanRows.Combine:
#                 combine_element = CodeplanElement(f'comb{entry.index}', entry.label, CodeplanElements.Combine, current_parent, current_level)
#                 for c in entry.combine_codes:
#                     combine_child = CodeplanElement(f'CB_{c}', '', CodeplanElements.Regular, combine_element, current_level + 1)
#                     combine_element.children.append(combine_child)
#                     self.flat_elements.append(combine_child)
#                 current_parent.children.append(combine_element)
#                 self.flat_elements.append(combine_element)

#     def _validate_elements(self):
#         for element in self.flat_elements:
#             if not element.is_valid():
#                 self.errors.append(f'Error in element {element.name}: "{element.label}"')


# class CodeplanElement():

#     def __init__(self, name, label, element_type, parent, level):
#         self.name = name
#         self.label = label
#         self.element_type = element_type
#         self.parent = parent
#         self.level = level
#         self.children = []

#     def _get_axis(self):
#         label = self.label.replace(r"'", r"''")
#         if self.element_type == CodeplanElements.Root:
#             return f'{{{",".join(c._get_axis() for c in self.children)}}}'
#         elif self.element_type == CodeplanElements.Base:
#             return f'base()'
#         elif self.element_type == CodeplanElements.Net:
#             return f'{self.name} \'{label}\' net({{{",".join(c._get_axis() for c in self.children)}}})'
#         elif self.element_type == CodeplanElements.Combine:
#             return f'{self.name} \'{label}\' combine({{{",".join(c._get_axis() for c in self.children)}}})'
#         elif self.element_type == CodeplanElements.Regular:
#             return f'{self.name}'

#     def is_valid(self):


#     def __repr__(self):
#         return f'CodeplanElement(code={self.name})'