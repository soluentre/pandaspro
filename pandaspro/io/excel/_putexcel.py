from pathlib import Path
import os
import xlwings as xw
from pandaspro.core.stringfunc import parse_wild, parse_method
from pandaspro.io.excel._framewriter import FramexlWriter, StringxlWriter, cpdFramexl
from pandaspro.io.excel._xlwings import RangeOperator, parse_format_rule


def is_range_filled(ws, range_str: str = None):
    if range_str is None:
        return False
    else:
        rng = ws.range(range_str)
        for cell in rng:
            if cell.value is not None and str(cell.value).strip() != '':
                return True
        return False


def is_sheet_empty(sheet):
    used_range = sheet.used_range
    if used_range.shape == (1, 1) and not used_range.value:
        return True
    return False


class PutxlSet:
    def __init__(
            self,
            workbook: str,
            sheet_name: str = None,
            alwaysreplace: str = None,  # a global config that sets all the following actions to replace ...
            noisily: bool = None
    ):
        def _extract_filename_from_path(path):
            return Path(path).name

        def _get_open_workbook_by_name(name):
            # Return the open workbook by its name if exists, otherwise return None
            for curr_app in xw.apps:
                for curr_wb in curr_app.books:
                    if curr_wb.name == name:
                        return curr_wb, curr_app
            return None, None

        # App and Workbook declaration
        wb, app = _get_open_workbook_by_name(_extract_filename_from_path(workbook))  # Check if the file is already open
        if wb:
            if noisily:
                print(f"{workbook} is already open, closing ...")
            wb.save()
            wb.close()
            if not app.books:  # Check if the app has no more workbooks open; if true, then quit the app
                app.quit()
        elif noisily:
            print(f"Working on {workbook} now ...")

        if not os.path.exists(workbook):  # Check if the file already exists
            wb = xw.Book()  # If not, create a new Excel file
            wb.save(workbook)
        else:
            wb = xw.Book(workbook)

        # Worksheet declaration
        if sheet_name is None:
            sheet_name = wb.sheets[0].name

        current_sheets = [sheet.name for sheet in wb.sheets]
        if sheet_name in current_sheets:
            sheet = wb.sheets[sheet_name]
        else:
            sheet = wb.sheets.add(after=wb.sheets.count)
            sheet.name = sheet_name

        if 'Sheet1' in current_sheets and is_sheet_empty(wb.sheets['Sheet1']) and sheet_name != 'Sheet1':
            wb.sheets['Sheet1'].delete()

        self.app = app
        self.workbook = workbook
        self.worksheet = sheet_name
        self.wb = wb
        self.ws = sheet
        self.globalreplace = alwaysreplace
        self.io = None

    def putxl(
            self,
            content = None,
            sheet_name: str = None,
            cell: str = 'A1',
            index: bool = False,
            header: bool = True,
            replace: str = None,
            sheetreplace: bool = False,

            # Section. String Format
            font: str | tuple = None,
            font_name: str = None,
            font_size: int = None,
            font_color: str | tuple = None,
            italic: bool = False,
            bold: bool = False,
            underline: bool = False,
            strikeout: bool = False,
            align: str | list = None,
            merge: bool = None,
            border: str | list = None,
            fill: str | tuple | list = None,
            fill_pattern: str = None,
            fill_fg: str | tuple = None,
            fill_bg: str | tuple = None,
            appendix: bool = False,

            # Section. df format
            index_merge: dict = None,
            header_wrap: bool = None,
            adjust_height: dict = None,
            df_format: dict = None,

            debug: bool = False,
    ) -> None:

        # Pre-Cleaning: (1) transfer FramePro to dataframe; (2) change tuple cells to str
        ################################
        if hasattr(content, 'df'):
            content = content.df

        if not isinstance(content, str):
            for col in content.columns:
                content[col] = content[col].apply(lambda x: str(x) if isinstance(x, tuple) else x)

        # Sheetreplace? If a sheet_name is specified, then override the current sheet
        ################################
        replace_type = self.globalreplace if self.globalreplace else replace

        if sheet_name and sheet_name != self.worksheet:
            if sheet_name in [sheet.name for sheet in self.wb.sheets]:
                ws = self.wb.sheets[sheet_name]
            else:
                ws = self.wb.sheets.add(after=self.wb.sheets.count)
                ws.name = sheet_name
        else:
            ws = self.ws

        # If sheetreplace or replace is specified, then delete the old sheet and create a new one
        ################################
        if sheetreplace or replace_type == 'sheet':
            _sheetmap = {sheet.index: sheet.name for sheet in self.wb.sheets}
            original_index = ws.index
            original_name = ws.name
            total_count = self.wb.sheets.count
            if debug:
                print(f">>> Row 121: original index is {original_index}")
            if original_index == total_count:
                new_sheet = self.wb.sheets.add(after=self.wb.sheets[_sheetmap[original_index]])
                if debug:
                    print(f">>> Row 128: New sheet added after the sheet !'{_sheetmap[original_index]}'")
            else:
                new_sheet = self.wb.sheets.add(before=self.wb.sheets[_sheetmap[original_index + 1]])
                if debug:
                    print(f">>> Row 132: New sheet added before the sheet !'{_sheetmap[original_index + 1]}'")
            ws.delete()
            new_sheet.name = original_name
            ws = new_sheet
            self.ws = ws
            not_replace_warning = False
        else:
            if not isinstance(content, str):
                not_replace_warning = True
            else:
                not_replace_warning = False

        # Declare IO Object
        ################################
        if isinstance(content, str):
            io = StringxlWriter(content=content, cell=cell)
            RangeOperator(ws.range(io.cell)).format(
                font=font,
                font_name=font_name,
                font_size=font_size,
                font_color=font_color,
                italic=italic,
                bold=bold,
                underline=underline,
                strikeout=strikeout,
                align=align,
                merge=merge,
                border=border,
                fill=fill,
                fill_pattern=fill_pattern,
                fill_fg=fill_fg,
                fill_bg=fill_bg,
                appendix=appendix
            )
            self.io = io
            ws.range(io.cell).value = io.content

        else:
            io = FramexlWriter(content=content, cell=cell, index=index, header=header)
            ws.range(io.cell).value = io.content
            self.io = io

        # Format the sheet (Shelley, Li)
        ################################
        '''
         Extra Format (not in the group of format parameters): highlight area in existing-content excel
         This is embedded and will be triggered automatically if not replacing sheet 
         '''
        if not_replace_warning:
            matchdict = {
                'top': self.io.range_top_empty_checker,
                'bottom': self.io.range_bottom_empty_checker,
                'left': self.io.range_left_empty_checker,
                'right': self.io.range_right_empty_checker
            }
            for direction in list(matchdict.keys()):
                if is_range_filled(self.ws, matchdict[direction]):
                    RangeOperator(ws.range(self.io.range_all)).format(border=[direction, 'thicker', '#FF0000'])

        '''
        For index_merge para, the accepted dict only accepts two keys:
        1. level: for which level of the index to be set as merge benchmark
        2. columns: for which columns should apply the merge according to the benchmark index
        
        columns can either be a list or a str, and power-wildcard is embedded when using str:
        >>> ['grade', 'staff_id', 'age']
        >>> '* Total' 
        # this will match all columns in the dataframe ends with Total
        '''
        if index_merge:
            for key, local_range in io.range_index_merge_inputs(**index_merge).items():
                RangeOperator(self.ws.range(local_range)).format(merge=True)

        if header_wrap:
            RangeOperator(self.ws.range(io.range_header)).format(wrap=True)

        '''
        For adjust_height para, the accepted dict must use column/index name as keys
        The direct value follow each column/index name must be a dictionary, 
        and there must be a key of "width" in it
        
        For example:
        >>> {
        >>>     'staff id': {'width': 24, 'color': '#00FFFF'},
        >>>     'age': {'width': 15}
        >>>     'salary': {'width': 30, 'haligh': 'left'}
        >>> }
        '''
        if adjust_height:
            for name, setting in adjust_height.items():
                if name in io.columns:
                    RangeOperator(self.ws.range(io.get_column_letter_by_name(name).cell)).format(width=setting['width'])
                if name in io.rawdata.index.names:
                    RangeOperator(self.ws.range(io.get_column_letter_by_indexname(name).cell)).format(width=setting['width'])

        '''
        df_format: the main function to add format to ranges
        This parameter will take a dictionary which uses:
        (1) format prompt key words as the keys
        (2) a list of range key words, which may be just a str term (attribute) ... 
            or a cpdFramexl object 
            
        >>> ... df_format('msblue80': 'header')
        >>> ... df_format('msblue80': cpdFramexl(name='index_merge_inputs', level='cmu_dept_major', columns=['age', 'salary']))
        '''
        if df_format:
            for rule, rangeinput in df_format.items():
                # Parse the format to a dictionary, passed to the .format for RangeOperator
                # parse_format_rule is taken from _xlwings module
                format_kwargs = parse_format_rule(rule)

                # Declare range as list/cpdFramexl Object
                def _declare_ranges(local_input):
                    if isinstance(local_input, str):
                        parsedlist = [item.strip() for item in local_input.split(',')]
                        cpdframexl_dict = None
                    elif isinstance(local_input, list):
                        parsedlist = local_input
                        cpdframexl_dict = None
                    elif isinstance(local_input, cpdFramexl):
                        parsedlist = None
                        cpdframexl_dict = getattr(io, 'range' + local_input.name)(**local_input.paras)
                    else:
                        raise ValueError('Unsupported type in df_format dictionary values')
                    return parsedlist, cpdframexl_dict

                ioranges, dict_from_cpdframexl = _declare_ranges(rangeinput)

                if ioranges:
                    for each_range in ioranges:
                        # Parse the input string as method name + kwargs
                        print(parse_method(each_range)[1])
                        range_affix, method_kwargs = parse_method(each_range)[0], parse_method(each_range)[1]
                        attr_method = getattr(io, 'range_' + range_affix)
                        if callable(attr_method):
                            range_cells = attr_method(**method_kwargs)
                        else:
                            range_cells = attr_method

                        if isinstance(range_cells, dict):
                            for range_key, range_content in range_cells.items():
                                RangeOperator(self.ws.range(range_content)).format(**format_kwargs)
                        elif isinstance(range_cells, str):
                            RangeOperator(self.ws.range(range_cells)).format(**format_kwargs)

                if dict_from_cpdframexl:
                    for range_key, range_content in dict_from_cpdframexl.items():
                        RangeOperator(self.ws.range(range_content)).format(**format_kwargs)

        # Remove Sheet1 if blank and exists (the Default tab) ...
        ################################
        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if 'Sheet1' in current_sheets and is_sheet_empty(self.wb.sheets['Sheet1']):
            self.wb.sheets['Sheet1'].delete()

        self.wb.save()

        if debug:
            print(f"\n>>> Cell Range Analysis")
            print(f" ----------------------")
            print(f">>> Total row: {io.tr}, Total column: {io.tc}")
            print(
                f">>> Range index: {io.range_index}, Range header: {io.range_header}, Range index names: {io.range_indexnames}\n")

    def switchtab(self, sheet_name: str) -> None:
        """
        Switches to a specified sheet in the workbook.
        If the sheet does not exist, it creates a new one with the given name.

        Parameters
        ----------
        sheet_name : str
            The name of the sheet to switch to or create.
        """
        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if sheet_name in current_sheets:
            sheet = self.wb.sheets[sheet_name]
        else:
            sheet = self.wb.sheets.add(after=self.wb.sheets.count)
            sheet.name = sheet_name
        self.ws = sheet
        return


if __name__ == '__main__':

    from wbhrdata import wbuse_pivot, hrconfig
    from pandaspro import sysuse_auto, sysuse_countries

    df1 = sysuse_auto
    df2 = sysuse_countries

    # ps = PutxlSet('sampledf.xlsx', 'Sheet3', noisily=True)
    # ps.putxl(df, 'TT', 'A1', index=True, header=True, sheetreplace=True, debug=True)
    # ps.putxl(df1, 'TF', 'A1', index=True, header=False, sheetreplace=True, debug=True)
    # ps.putxl(df1, 'FT', 'A1', index=False, header=True, sheetreplace=True, debug=True)
    # ps.putxl(df1, 'FF', 'A1', index=False, header=False, sheetreplace=True, debug=True)
    # from _xlwings import cpdStyle
    e = PutxlSet('sampledf.xlsx', sheet_name='region')
    e.putxl(
        wbuse_pivot,
        cell='B2',
        index=True,
        index_merge={'level': 'cmu_dept_major', 'columns': '* Total'},
        adjust_height=hrconfig,
        header_wrap=True,
        df_format={
            'msblue80, align=center, border=outer_thick': ['index_hsections(level=cmu_dept_major)', 'columnspan(start_col=GC Total, stop_col=Ratio Total, header=True)'],
            'msgreen80, align="center"': 'header_outer',
        }
    )
