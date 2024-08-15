from pandaspro.io.excel.writer import FramexlWriter
from pandaspro.io.excel.range_operator import RangeOperator
import xlwings as xw


class CustomStyle:
    def __init__(
            self,
            ws,
            framewriter: FramexlWriter,
    ):
        self.ws = ws
        self.frame = framewriter

        self.allstyle = RangeOperator(self.ws.range(self.frame.range_all))
        self.indexstyle = RangeOperator(self.ws.range(self.frame.range_index))

    def style1(self):
        self.allstyle.format(border='all, 3')

    def style2(self):
        self.allstyle.format(border='inner, 1')
        self.allstyle.format(border='outer, 3')




if __name__ == '__main__':
    pass

    #
    # ws = xw.Book('sampledf.xlsx').sheets['FF']
    # ws.range('G1').value = df
    # a = FramexlWriter(df, 'G1', index=True)
    #
    # b = CustomStyle(ws, a)
    # print('end')44
    #
    # ## Create a CdFormat
    # myformat = CdFormat('grade', applyrange='self', rows={'GA':{'font_color':'#FFFF00'}})
    # myformat2 = CdFormat('salary', applrange='age', rows={range(0,100):'bold'})
    # a.putxl(df, 'sheet', 'G1', cdformat=[myformat, myformat2])
    #
    # cdFormat('grade', {}, ['age','name'])
    #
    # filter1  = df[df['grade']>df['hgrade2']]
    # myformat = (cdFormat(conplex)
    # myformat.add('#FFF000', index_mask=filter1, applyrange=['grade','grade2'])
    #
    # dfmap[filter1]


    # myformat.column = 'grade'
    # .applyrange
    #
    # applycells
    #
    # myformat.coredict = {'font_color':'#FFFF00'}
    # RangeOperator(ws.range(applycells)).format(**myformat.coredict))