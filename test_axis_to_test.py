# pylint: disable-msg=w0614

from settings import *
from codeplans import *
from dimensions_tools import *
from win32com import client

codeplan = MDDFile('test\\axis_to_txt\\codeplan_1542904179786_2018-11-22.mdd')

with open('test\\axis_to_txt\\trees_mdd_11_new.txt', mode='w', encoding='utf-8') as mdd_output:

    for mdd_cp in codeplan:

        mdd_output.write('****************************\n')
        mdd_output.write(mdd_cp.name  + '\n')
        mdd_output.write('****************************\n')
        for c in mdd_cp.tree.xl_children:
            mdd_output.write(c + '\n')



print('OK')
