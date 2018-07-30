from mdd_codeplans import MDDCodeplans
from excel_codeplans import ExcelCodeplans

mdd_excel_codeplan_glue = {

}

def main():
    cp = MDDCodeplans('codeplan_1531899367053_2018-07-18.mdd')
    xl = ExcelCodeplans('Codeplan KTV online 201807_VORAB.xlsx')
    print('ok')

if __name__ == '__main__':
    from timeit import timeit
    print(timeit(main,number=1))