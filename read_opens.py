# read_opens.py
# pylint: disable-msg=w0614

from codeplans import MDDFile, XLCodeplan, CFileManager, update_master_mdd
from settings import *

def merge_mdd_with_xl():
 
    mdd_file = MDDFile(MDD_CODEPLAN)
    for mdd_name, xl_name, _, other_element in ADAPTER:
        if mdd_name and xl_name and mdd_name in mdd_file:
            mdd_codeplan = mdd_file[mdd_name]
            xl_codeplan = XLCodeplan(
                path=EXCEL_CODEPLAN,
                sheet_name=xl_name
            )
            xl_codeplan.other_element = other_element
            xl_codeplan.print_summary()
            mdd_codeplan.print_summary()
            mdd_codeplan.inject(xl_codeplan)

    mdd_file.save_mdd(ADJUSTED_MDD_CODEPLAN)
    cfile_manager = CFileManager(mdd_file, VERBACO_CFILE)
    cfile_manager.save_cfile(ADJUSTED_VERBACO_CFILE)

def main():
    merge_mdd_with_xl()
    update_master_mdd(OUTPUT_PATH, ADJUSTED_MDD_CODEPLAN, ADAPTER)

if __name__ == '__main__':
    from timeit import timeit
    print(timeit(merge_mdd_with_xl,number=1))
