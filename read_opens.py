# read_opens.py
# pylint: disable-msg=w0614

from codeplans import *
from settings import *

def main():

    # create master verbaco mdd by merging previous waves
    # master_verbaco = MDDFile(f'{JOB_ROOT}Data\\Coding\\Raw\\codeplan_1536064682893_2018-09-04.mdd')
    # slave_verbaco = MDDFile(f'{JOB_ROOT}Data\\Coding\\Raw\\codeplan_1533204901432_2018-08-02.mdd')
    # merged_verbaco = MDDFileMerger(master_verbaco, slave_verbaco).merge()
    # merged_verbaco.save_as(MDD_CODEPLAN)

    # merges final verbaco mdd with excel
    verbaco_mdd = MDDFile(MDD_CODEPLAN)
    xl_codeplans = XLFile(EXCEL_CODEPLAN)
    mdd_xl_merger = MDDXLFileMerger(verbaco_mdd, xl_codeplans, ADAPTER, verbose=True)
    adjusted_verbaco_mdd = mdd_xl_merger.merge_all()
    adjusted_verbaco_mdd.save_as(ADJUSTED_MDD_CODEPLAN)

    # produces variable and category maps
    adjusted_verbaco_mdd.save_variable_map(VARIABLE_MAP)
    mdd_xl_merger.save_category_map(CATEGORY_MAP)
    
    # update cfile using maps from above
    cfile_updater = CFileManager(VERBACO_CFILE, VARIABLE_MAP, CATEGORY_MAP)
    cfile_updater.save_cfile(ADJUSTED_VERBACO_CFILE)

    # update master file verbaco mdd
    copy_mdd_ddf(INPUT_PATH, OUTPUT_PATH)
    update_master_with_mdd_codeplan_with_adapter(f'{OUTPUT_PATH}.mdd', ADJUSTED_MDD_CODEPLAN, ADAPTER)

    # executes cfiles
    execute_opens(MROLEDB_CONNECTION_STRING, DB_CFILE)
    execute_opens(MROLEDB_CONNECTION_STRING, ADJUSTED_VERBACO_CFILE)
    execute_opens(MROLEDB_CONNECTION_STRING, DB_CORRECTION_CFILE)

if __name__ == '__main__':
    from timeit import timeit
    print(timeit('main()',globals=globals(),number=1))
