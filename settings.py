# settings.py

from collections import namedtuple

ROUND_LABEL = '2018-09'
JOB_ROOT = f'test\\{ROUND_LABEL}\\'

# input 
MDD_CODEPLAN = f'{JOB_ROOT}Data\\Coding\\Raw\\codeplan_1538475273482_2018-10-02.mdd'
#XLS in neuem Format abspeichern:
EXCEL_CODEPLAN = f'{JOB_ROOT}Data\\Coding\\Codeplan KTV online 201809.xlsx'
VERBACO_CFILE = f'{JOB_ROOT}Data\\Coding\\Raw\\verbatims_1538475157942_02.10.2018.txt'
DB_CFILE = f'{JOB_ROOT}Data\\Coding\\Raw\\dw_ktv_cfile_{ROUND_LABEL}.txt'
DB_CORRECTION_CFILE = f'{JOB_ROOT}Data\\Coding\\dw_ktv_cfile_{ROUND_LABEL}_corrections_MaF.txt'

INPUT_PATH = f'{JOB_ROOT}Data\\KTV_Online_FINAL_{ROUND_LABEL}'

# intermediate
ADJUSTED_MDD_CODEPLAN = f'{JOB_ROOT}Data\\Coding\\verbaco_codeplan_adjusted_{ROUND_LABEL}.mdd'
VARIABLE_MAP = f'{JOB_ROOT}Data\\Coding\\variable_map_{ROUND_LABEL}.txt'
CATEGORY_MAP = f'{JOB_ROOT}Data\\Coding\\category_map_{ROUND_LABEL}.txt'
ADJUSTED_VERBACO_CFILE = f'{JOB_ROOT}Data\\Coding\\verbaco_cfile_adjusted_{ROUND_LABEL}.txt'

# output
OUTPUT_PATH = f'{JOB_ROOT}Data\\KTV_Online_FINAL_{ROUND_LABEL}_withOpens'
MROLEDB_CONNECTION_STRING = f'''
    Provider=mrOleDB.Provider.2;
    Data Source=mrDataFileDsc;
    Location={OUTPUT_PATH}.ddf;
    Initial Catalog={OUTPUT_PATH}.mdd;
    MR Init MDM Access=1;
    MR Init Category Names=1;'''

#map
CodeplanMap = namedtuple('CodeplanMap', 'mdd_name xl_name master_name other_element')
ADAPTER = [
    CodeplanMap('head_11326', 'CP Versicherungen KTV_BEARB', 'cp_marken', 'CB_999'),
    CodeplanMap('head_201', 'CP Winh CosmosDirekt_BEARB', 'cp_f4_cosdir', 'CB_99'),
    CodeplanMap('head_193', 'CP Winh HUK_BEARB', 'cp_f4_hukc', 'CB_37'),
    CodeplanMap('head_10960', 'CP Winh DKV_BEARB', 'cp_f4_dkv', 'CB_94'),
    CodeplanMap('head_6793', 'CP Winh Ergo DV', 'cp_f4_ergod', 'CB_99'),
    CodeplanMap('head_6794', 'CP Winh ERGO', 'cp_f4_ergo', 'CB_99'),
    CodeplanMap('head_174', 'CP Winh D.A.S.', 'cp_f4_das', 'CB_73'),
    CodeplanMap('head_292', 'CP Winh Gothaer', 'cp_f4_goth', 'CB_80'),
    CodeplanMap('head_192', 'CP Winh VHV', 'cp_f4_vhv', 'CB_99'),
    CodeplanMap('head_197', 'CP Winh Württembergische', 'cp_f4_wuert', 'CB_6'),
    CodeplanMap('head_203', 'CP Winh Aachen Münchener', 'cp_f4_am', 'CB_99'),
    CodeplanMap('head_205', 'CP Winh Generali', 'cp_f4_gen', 'CB_99'),
    CodeplanMap('head_204', 'CP Winh Advocard', 'cp_f4_advo', 'CB_40'),
    CodeplanMap('head_11771', 'Bausparkassen', 'cp_ww1c1', 'CB_99'),
    CodeplanMap('head_13974', 'CP Winh Europa', 'cp_f4_europa', 'CB_999'),
    CodeplanMap('head_15880', 'CP Winh Barmenia', 'cp_f4_barmenia', 'CB_999'),
    CodeplanMap('head_11416', 'CP HUKINT', 'cp_zhukint', 'CB_999'),
    CodeplanMap('head_37121', 'CP Winh DEVK', 'cp_f4_devk', 'CB_999'),
    CodeplanMap('head_46492', 'CP Adam Riese', 'cp_sfww2', 'CB_52'),
    CodeplanMap('head_31792', 'CP Winh Swiss Life', 'cp_f4_sl', 'CB_997'),
    CodeplanMap('head_46875', 'CP ZGENTESTB1', 'cp_zgentestb1', 'CB_99'),
    CodeplanMap('head_46876', 'CP ZGENTESTB2', 'cp_zgentestb2', 'CB_99'),
    CodeplanMap('head_46877', 'CP ZGENTESTE', 'cp_zgenteste', 'CB_99'),
    CodeplanMap('head_195', 'CP Winh Hannoversche Leben', 'cp_f4_hl', 'CB_99'),
    CodeplanMap('head_48183', 'CP ZCINT6a', 'cp_zcint6a', 'CB_999'),
    CodeplanMap('head_14437', 'CP ERGSO', 'cp_zergso', 'CB_997'),
	CodeplanMap('head_51522', 'CP ZHUKBER34', 'cp_ZHUKBER34', 'CB_999'),
	CodeplanMap('head_51578', 'CP ZSL_Marke', 'cp_ZSL_Marke', 'CB_999'),
    CodeplanMap(None, 'CP Winh Nürnberger Versicherung', 'cp_f4_nv', None),
    CodeplanMap(None, 'CP ZHUKUNTa', 'cp_zhuk24unta', None),
    CodeplanMap(None, 'CP ZHUKUNTb', 'cp_zhuk24untb', None),
    CodeplanMap(None, 'CP ZCINC2_6', 'cp_zcinc', None),
    CodeplanMap(None, 'CP ZHUKSK7', 'cp_zhuksk7', None),
    CodeplanMap(None, 'CP ZCDV4', 'cp_zcdv4', None),
    CodeplanMap(None, 'CP ZCDV1_2', 'cp_zcdv1_2', None),
    CodeplanMap(None, 'CP ZCREC4', 'cp_zcrec4', None),
    CodeplanMap(None, 'CP ZCREC2-3', 'cp_zcrec2_3', None),
    CodeplanMap(None, 'CP ZNUE_SPON1', 'cp_znue_spon', None),
    CodeplanMap(None, 'CP ZNUE_VEREIN2', 'cp_znue_verein', None),
    CodeplanMap(None, 'CP Cosmos VP Marken', 'cp_zcport1', None),
    CodeplanMap(None, 'CP Cosmos VP Nutzung', 'cp_zcport_5s_8s', None),
    CodeplanMap(None, 'CP Cosmos VP Begründung', 'cp_zcport_rest', None)
]
