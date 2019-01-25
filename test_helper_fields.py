import codeplans
from win32com import client
from dimensions_tools import remove_helper_fields, copy_mdd_ddf_data

input_path = 'test\\helper_fields\\KTV_Online_FINAL_2018-10_withOpens'
output_path = 'test\\helper_fields\\test'

copy_mdd_ddf_data(input_path, output_path, only_mdd=True)

remove_helper_fields(f'{output_path}.mdd')
