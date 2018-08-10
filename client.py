from codeplans import MDDFile, XLCodeplan

def main():
    mdd_file = MDDFile('codeplan_1533204901432_2018-08-02.mdd', 'verbatims_1533204821119_02.08.2018_2.txt')

    mdd_xl_adapter = {
        'head_11326': 'CP Versicherungen KTV_BEARB',
        'head_201': 'CP Winh CosmosDirekt_BEARB',
        'head_193': 'CP Winh HUK_BEARB',
        'head_10960': 'CP Winh DKV_BEARB',
        'head_6793': 'CP Winh Ergo DV',
        'head_6794': 'CP Winh ERGO',
        'head_174': 'CP Winh D.A.S.',
        'head_292': 'CP Winh Gothaer',
        'head_192': 'CP Winh VHV',
        'head_197': 'CP Winh Württembergische',
        'head_203': 'CP Winh Aachen Münchener',
        'head_205': 'CP Winh Generali',
        'head_204': 'CP Winh Advocard',
        'head_11771': 'Bausparkassen',
        'head_13974': 'CP Winh Europa',
        'head_15880': 'CP Winh Barmenia',
        'head_11416': 'CP HUKINT',
        'head_37121': 'CP Winh DEVK',
        'head_46492': 'CP Adam Riese',
        'head_14437': 'CP ERGSO',
        'head_31792': 'CP Winh Swiss Life',
        'head_46875': 'CP ZGENTESTB1',
        'head_46876': 'CP ZGENTESTB2',
        'head_46877': 'CP ZGENTESTE'
    }


    for mdd_name, xl_name in mdd_xl_adapter.items():
        mdd_codeplan = mdd_file.get_codeplan(mdd_name)
        xl_codeplan = XLCodeplan(
            path='Examples//Codeplan KTV online 201807.xlsx',
            sheet_name=xl_name,
            other_element_name='CB_999'
        )

        xl_codeplan.print_summary()
        mdd_codeplan.print_summary()
        mdd_codeplan.inject(xl_codeplan)

    mdd_file.save_mdd('new.mdd')
    mdd_file.save_cfile('new.txt')
    print('OK')


if __name__ == '__main__':
    from timeit import timeit
    print(timeit(main,number=1))
