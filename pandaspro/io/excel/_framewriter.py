from pandaspro.core.stringfunc import parsewild
from pandaspro.io.excel._utils import CellPro
import pandas as pd


class StringxlWriter:
    def __init__(
            self,
            content,
            cell: str,
    ) -> None:
        self.iotype = 'str'
        self.content = content
        self.cell = cell


class FramexlWriter:
    def __init__(
            self,
            content,
            cell: str,
            index: bool = False,
            header: bool = True,
            column_list: list = None,
            cols_index_merge: str | list = None,
            index_mask=None
    ) -> None:
        cellobj = CellPro(cell)
        header_row_count = len(content.columns.levels) if isinstance(content.columns, pd.MultiIndex) else 1
        index_column_count = len(content.index.levels) if isinstance(content.index, pd.MultiIndex) else 1

        dfmapstart = cellobj.offset(header_row_count, index_column_count)
        dfmap = content.copy()
        dfmap = dfmap.astype(str)

        # Create a cells Map
        i = 0
        for dfmap_index, row in dfmap.iterrows():
            j = 0
            for col in dfmap.columns:
                dfmap.loc[dfmap_index, col] = dfmapstart.offset(i, j).cell
                j += 1
            i += 1

        self.formatrange = "Please provide the column_list/index_mask to select a sub-range"

        if column_list:
            if isinstance(column_list, str):
                column_list = parsewild(column_list, dfmap.columns)
            self.formatrange = dfmap[column_list]
        if index_mask:
            self.formatrange = self.formatrange[index_mask]

        # Calculate the Ranges
        content = pd.DataFrame(content.to_dict())
        if header == True and index == True:
            self.export_type = 'htit'
            tr, tc = content.shape[0] + header_row_count, content.shape[1] + index_column_count
            export_data = content
            range_index = cellobj.offset(header_row_count, 0).resize(tr - header_row_count, index_column_count)
            range_indexnames = cellobj.resize(header_row_count, header_row_count)
            range_header = cellobj.offset(0, index_column_count).resize(header_row_count, tc - index_column_count)
        elif header == False and index == True:
            self.export_type = 'hfit'
            tr, tc = content.shape[0], content.shape[1] + index_column_count
            export_data = content.reset_index().to_numpy().tolist()
            range_index = cellobj.resize(tr, index_column_count)
            range_indexnames = 'N/A'
            range_header = 'N/A'
        elif header == False and index == False:
            self.export_type = 'hfif'
            tr, tc = content.shape[0], content.shape[1]
            export_data = content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = 'N/A'
        else:
            self.export_type = 'htif'
            tr, tc = content.shape[0] + header_row_count, content.shape[1]
            if isinstance(content.columns, pd.MultiIndex):
                column_export = [list(lst) for lst in list(zip(*content.columns.values))]
            else:
                column_export = [content.columns.to_list()]
            export_data = column_export + content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = cellobj.resize(header_row_count, tc)

        self.iotype = 'df'
        self.rawdata = content
        self.columns = self.rawdata.columns
        self.content = export_data
        self.cell = cell
        self.tr = tr
        self.tc = tc
        self.header_row_count = header_row_count
        self.index_column_count = index_column_count

        # data corners - cellpros
        self.start_cell = cellobj.offset(header_row_count, index_column_count)
        self.top_right_cell = cellobj.offset(0, self.tc - 1).cell
        self.bottom_left_cell = cellobj.offset(self.tr - 1, 0).cell
        self.end_cell = cellobj.offset(self.tr - 1, self.tc - 1).cell

        # ranges
        self.range_all = cell + ':' + self.end_cell
        self.range_data = self.start_cell.resize(tr - header_row_count, tc - index_column_count).cell
        self.range_index = range_index.cell if range_index != 'N/A' else 'N/A'
        self.range_index_outer = CellPro(self.cell).resize(self.tr, self.index_column_count).cell
        self.range_header = range_header.cell if range_header != 'N/A' else 'N/A'
        self.range_header_outer = CellPro(self.cell).resize(self.header_row_count, self.tc).cell
        self.range_indexnames = range_indexnames.cell if range_indexnames != 'N/A' else 'N/A'
        self.range_top_empty_checker = CellPro(self.cell).offset(-1, 0).resize(1, self.tc).cell if CellPro(self.cell).cell_index[0] != 1 else None
        self.cellmap = dfmap
        if cols_index_merge:
            self.cols_index_merge = cols_index_merge if isinstance(cols_index_merge, list) else parsewild(cols_index_merge, content)
        else:
            self.cols_index_merge = None

    def get_column_letter_by_name(self, colname):
        rowcount = list(self.columns).index(colname)
        col_cell = self.start_cell.offset(0, rowcount)
        if self.export_type in ['htif', 'hfif']:
            col_cell = col_cell.offset(0, -self.index_column_count)
        return col_cell

    def _index_break(self, level: str = None):
        temp = self.content.reset_index()

        def _count_consecutive_values(series):
            return series.groupby((series != series.shift()).cumsum()).size().tolist()

        return _count_consecutive_values(temp[level])

    def range_index_merge_inputs(self, level: str = None):
        result_dict = {}
        if self.cols_index_merge is None:
            raise ValueError('index_merge_inputs method requires cols_index_merge to be passed when constructing the FramexlWriter Object')
        else:
            for index, col in enumerate(self.cols_index_merge):
                merge_start_each = self.get_column_letter_by_name(col)
                for localid, rowspan in enumerate(self._index_break(level=level)):
                    result_dict[f'col{index}_{localid}_{rowspan}'] = merge_start_each.resize(rowspan, 1).cell
                    merge_start_each = merge_start_each.offset(rowspan, 0)
        return result_dict

    def range_index_horizontal_sections(self, level: str = None):
        if self.range_index is None:
            raise ValueError('index_sections method requires the input dataframe to have an index')
        else:
            result_dict = {}
            result_dict['headers'] = CellPro(self.cell).resize(self.header_row_count, self.tc).cell
            range_start_each = CellPro(self.cell).offset(self.header_row_count, 0)
            for localid, rowspan in enumerate(self._index_break(level=level)):
                result_dict[f'section_{localid}_{rowspan}'] = range_start_each.resize(rowspan, self.tc).cell
                range_start_each = range_start_each.offset(rowspan, 0)
        return result_dict

    @property
    def range_index_levels(self):
        if self.range_index is None or not isinstance(self.rawdata.index, pd.MultiIndex):
            raise ValueError('index_levels method requires the input dataframe to be multi-index frame')
        else:
            result_dict = {}
            range_start_each = CellPro(self.cell)
            for each_index in self.rawdata.index.names:
                result_dict[f'index_{each_index}'] = range_start_each.resize(self.tr, 1).cell
                range_start_each = range_start_each.offset(0, 1)
            return result_dict

    def range_columnspan(self, start_col, stop_col):
        col_index1 = self.get_column_letter_by_name(start_col).cell_index[1]
        col_index2 = self.get_column_letter_by_name(stop_col).cell_index[1]
        row_index = self.get_column_letter_by_name(start_col).cell_index[0]


        top_left = self.get_column_letter_by_name(start_col)
        top_right = self.get_column_letter_by_name(stop_col)
        start_range = CellPro(top_left + ':' + top_right)
        return start_range.resize_h(self.tc)


if __name__ == '__main__':

    # a = FramexlWriter(sysuse_auto, 'G1', column_list='Country', index=True)
    # print(a.formatrange)
    #
    # paintdict = {
    #     'all': {
    #         'logic': 'grade == 1 and grade <2',
    #         'format': {
    #             'fill': '#FFF000',
    #             'font': 'bold 12'
    #         }
    #     }
    # }


    import wbhrdata as wb
    import xlwings as xw
    from pandaspro import sysuse_auto
    ws = xw.Book('sampledf.xlsx').sheets['sob']
    ws.range('G1').value = pd.DataFrame(sysuse_auto)

    ws = xw.Book('sampledf.xlsx').sheets['sob']
    data = wb.sob(region='AFE').pivot_table(index=['cmu_dept_major', 'cmu_dept'], values=['upi','age'], aggfunc='sum', margins_name='Total', margins=True)

    # core
    io = FramexlWriter(sysuse_auto, 'G1', index=False, header=True, cols_index_merge='upi, age')
    ws.range('G1').value = io.content

    # a = io.range_index_horizontal_sections(level='cmu_dept_major')
    b = io.range_index_outer
    # c = io.range_index_levels
    # d = io.range_columnspan('upi', 'age')
    mpg_col = io.get_column_letter_by_name('mpg').cell

    # xw.apps.active.api.DisplayAlerts = False
    # ws.range('A4:A9').api.MergeCells = True