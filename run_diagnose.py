
from diagnose import create_excel_comparison

OLD_MDD = f'test\\diagnose\\KTVONLINE_1810.mdd'
NEW_MDD = f'test\\diagnose\\KTVONLINE_1811.mdd'
EXCEL_COMPARISON_FILE = f'test\\diagnose\\wave_comparison.xlsx'

create_excel_comparison(OLD_MDD, NEW_MDD, EXCEL_COMPARISON_FILE)