from pandaspro.io.excel._putexcel import PutxlSet
from pandaspro.sampledf.sampledf import wbuse_pivot


path = './sampledf.xlsx'

ps = PutxlSet(path)
ps.putxl(
    wbuse_pivot,
    sheet_name='newtab3',
    cell='B2',
    index=True,
    design='wbblue',
    tab_color='blue'
)