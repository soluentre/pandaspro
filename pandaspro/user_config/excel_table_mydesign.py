import pandas as pd
dpath_config = r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Databases\config'

_file = f'{dpath_config}/export_formatting.xlsm'
_sheet = 'config'
_df = pd.read_excel(_file, sheet_name=_sheet).drop('class', axis=1)
hrconfig = _df.set_index('column').T.to_dict(orient='dict')

excel_export_mydesign = {
    'wbblue': {
        'style': 'blue',
        'cd': 'ti; pr',
        'config': hrconfig
    },
    'wbblue_pivot': {
        'style': 'blue; pivot_gray_lastrow; borders',
        'cd': 'ti; pr',
        'config': hrconfig
    },
    'wbbluelist': {
        'style': 'bluelist',
        'cd': 'ti; pr',
        'config': hrconfig
    },
    'wbgreen': {
        'style': 'green',
        'cd': 'ti',
        'config': hrconfig
    },
    'wbblack': {
        'style': 'black',
        'cd': 'ti',
        'config': hrconfig
    },
    'wbblue_grade': {
        'style': 'blue',
        'cd': 'grade; ti',
        'config': hrconfig
    },
    'wbgreen_grade': {
        'style': 'green',
        'cd': 'grade; ti',
        'config': hrconfig
    },
    'wbblack_grade': {
        'style': 'black',
        'cd': 'grade; ti',
        'config': hrconfig
    },
    'wbblue_pg': {
        'style': 'blue',
        'cd': 'pg; ti',
        'config': hrconfig
    },
    'wbgreen_pg': {
        'style': 'green',
        'cd': 'pg; ti',
        'config': hrconfig
    },
    'wbblack_pg': {
        'style': 'black',
        'cd': 'pg; ti',
        'config': hrconfig
    },
    'wbblue_cmu_dept': {
        'style': 'blue',
        'cd': 'cmu_dept; ti',
        'config': hrconfig
    },
    'wbgreen_cmu_dept': {
        'style': 'green',
        'cd': 'cmu_dept; ti',
        'config': hrconfig
    },
    'wbblack_cmu_dept': {
        'style': 'black',
        'cd': 'cmu_dept; ti',
        'config': hrconfig
    },
}
