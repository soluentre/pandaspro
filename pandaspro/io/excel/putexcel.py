import re
from pathlib import Path
import os

import pandas
import pandas as pd
import xlwings as xw
from pandaspro.core.stringfunc import parse_method, str2list
from pandaspro.io.excel.writer import FramexlWriter, StringxlWriter, cpdFramexl, CellxlWriter
from pandaspro.io.cellpro.cellpro import CellPro, cell_combine_by_column, is_cellpro_valid
from pandaspro.io.excel.range_operator import RangeOperator, parse_format_rule, color_to_int, _cpdpuxl_color_map
from pandaspro.utils.cpd_logger import cpdLogger


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


def parse_header_rule(header_str: str) -> dict:
    """
    Parses the header string for additional header control keywords.

    The function checks if the input string contains any of the allowed control keywords.
    Allowed keywords are: "merge_up", "merge_add_top", "merge_top", and "merge_add_row".

    Rules:
    - If more than one control keyword is present, a ValueError is raised.
    - If no control keyword is found, the function returns a dictionary with:
         'rule_extracted': the original header string,
         'additional_header': None.
    - If a control keyword is found, the function extracts the part of the string before the keyword,
      removing any trailing spaces and an optional semicolon. The extracted part is stored in
      'rule_extracted', and 'additional_header' is set to the detected keyword.
    - Additionally, if the detected keyword is either "merge_top" or "merge_add_row", then:
         * The function checks if the extracted header already contains the substring "merge".
         * If not, it appends "merge" to the extracted header. If the extracted header does not end
           with a semicolon before appending, a semicolon is added to separate the appended text.
    - Improvement for "merge_add_top":
         * If "merge_add_top" is detected, the function looks for the format "merge_add_top(XXX)".
         * It extracts the content inside the parentheses and adds a new key "merge_add_top" in the
           returned dictionary with the extracted content. If the parentheses are missing or empty,
           the value will be None.

    :param header_str: The header string to parse.
    :return: A dictionary with keys 'rule_extracted', 'additional_header', and optionally 'merge_add_top'.
    :raises ValueError: If more than one control keyword is found in the string.
    """
    allowed_keywords = ["merge_up", "merge_add_top"]

    # Identify which of the allowed keywords are present in the input string.
    found = [kw for kw in allowed_keywords if kw in header_str]

    # If more than one keyword is found, raise an error.
    if len(found) > 1:
        raise ValueError(
            "Only one additional header control is allowed: either merge with the above row or merge the above row separately."
        )

    # If no keyword is found, return the original string with additional_header set to None.
    if not found:
        return {"rule_extracted": header_str, "additional_header": None, "merge_add_top_title": None}

    keyword = found[0]

    # If merge_add_top is detected, check if it follows the format "merge_add_top(XXX)"
    # and extract the content inside the parentheses.
    merge_add_top_value = None
    if keyword == "merge_add_top":
        m = re.search(r'merge_add_top\s*\(\s*(.*?)\s*\)', header_str)
        if m:
            merge_add_top_value = m.group(1) if m.group(1) else None
        else:
            merge_add_top_value = None

    # Use a regex to capture the part of the string before the control keyword.
    # For merge_add_top, allow an optional parentheses part after the keyword.
    keyword_pattern = re.escape(keyword)
    if keyword == "merge_add_top":
        keyword_pattern += r'(?:\s*\(.*?\))?'
    pattern = re.compile(r'^(.*?)(?:\s*;?\s*(%s))\s*$' % keyword_pattern)
    match = pattern.match(header_str)
    if match:
        extracted = match.group(1).rstrip()  # Remove trailing whitespace.
    else:
        # Fallback: if regex doesn't match, manually extract the substring before the keyword.
        index = header_str.find(keyword)
        extracted = header_str[:index].rstrip()
        if extracted.endswith(";"):
            extracted = extracted[:-1].rstrip()

    # If the detected keyword is merge_up or merge_add_top, check if 'merge' and 'wrap' are already present.
    if keyword in ["merge_up", "merge_add_top"]:
        if "merge" not in extracted:
            # If there is no semicolon at the end of extracted, add one before appending "merge".
            if not re.search(r';\s*$', extracted):
                extracted += ';'
            extracted += " merge"
        if "wrap" not in extracted:
            # If there is no semicolon at the end of extracted, add one before appending "wrap".
            if not re.search(r';\s*$', extracted):
                extracted += ';'
            extracted += " wrap"

    if keyword == "merge_add_top":
        merge_add_top_title = merge_add_top_value
    else:
        merge_add_top_title = None

    result = {"rule_extracted": extracted, "additional_header": keyword, "merge_add_top_title": merge_add_top_title}

    return result


@cpdLogger
class PutxlSet:
    def __init__(
            self,
            workbook: str,
            sheet_name: str = None,
            alwaysreplace: str = None,  # a global config that sets all the following actions to replace ...
            noisily: bool = None,
    ):
        # App and Workbook declaration
        open_wb, app = PutxlSet._get_open_workbook_by_name(
            PutxlSet._extract_filename_from_path(workbook))  # Check if the file is already open
        if open_wb:
            if noisily:
                print(f"{workbook} is already open, closing ...")
            open_wb.save()
            open_wb.close()
            if not app.books:  # Check if the app has no more workbooks open; if true, then quit the app
                app.quit()
        elif noisily:
            print(f"Working on {workbook} now ...")

        if not os.path.exists(workbook):  # Check if the file already exists
            open_wb = xw.Book()  # If not, create a new Excel file
            open_wb.save(workbook)
        else:
            open_wb = xw.Book(workbook)

        # Worksheet declaration
        if sheet_name is None:
            sheet_name = open_wb.sheets[0].name

        current_sheets = [sheet.name for sheet in open_wb.sheets]
        if sheet_name in current_sheets:
            sheet = open_wb.sheets[sheet_name]
        else:
            sheet = open_wb.sheets.add(after=open_wb.sheets.count)
            sheet.name = sheet_name

        if 'Sheet1' in current_sheets and is_sheet_empty(open_wb.sheets['Sheet1']) and sheet_name != 'Sheet1':
            open_wb.sheets['Sheet1'].delete()

        self.open_wb, self.app = PutxlSet._get_open_workbook_by_name(
            PutxlSet._extract_filename_from_path(workbook))  # Check if the file is already open
        self.workbook = workbook
        self.wb = open_wb
        self.ws = sheet
        self.alwaysreplace = alwaysreplace
        self.io = None
        self.next_cell_down = None
        self.next_cell_right = None

    @property
    def colormap(self):
        return _cpdpuxl_color_map

    @staticmethod
    def _extract_filename_from_path(path):
        return Path(path).name

    @staticmethod
    def _get_open_workbook_by_name(name):
        # Return the open workbook by its name if exists, otherwise return None
        for curr_app in xw.apps:
            for curr_wb in curr_app.books:
                if curr_wb.name == name:
                    return curr_wb, curr_app
        return None, None

    # noinspection PyMethodMayBeStatic
    def helpfile(self, para='all'):
        cd_file = """
        cd_format: the main function to add format to core export data ranges (exc. headers and indices)
        This parameter will take a dict which allows only three keys (and applyto maybe omitted)
        (refer to the module _cdformat on the class design: _cdformat >> _framewriter.range_cdformat >> _putexcel.PutxlSet.putxl)

        key 1: column = indicating the conditional formatting columns
        key 2: rules = a dict with formatting rules (only based on the column above, like inlist, value equals to, etc.)
        key 3: applyto = where to apply, whether column itself or the whole dataframe, or, several selected columns     
        (default = self)

        >>> ... cd_format={'column': 'age', 'rules': {...}}
        >>> ... cd_format={'column': 'grade', 'rules': {'GA':'#FF0000'}, 'applyto': 'self'}
        
        For rules, it should be a dict:
        (1) the key can be values in the selected, then the value should be parable format
        (2) the key can also be a rule token, then the value should be a dict with r and f
        
        >>> ... {'GA': '#FF0000'}
        >>> ... {   
                    'rule1':  
                        {
                            'r': ['GA', 'GB', 'GC'],
                            'f': 'blue'
                        }，
                    'rule1':
                        {
                            'r': mask,
                            'f': 'green'
                        }
                }
        """
        if para == 'cd_format':
            print(cd_file)

    def putxl(
            self,
            content,
            sheet_name: str = None,
            cell: str = 'A1',
            index: bool = True,
            header: bool = True,
            replace: str = None,
            sheetreplace: bool = None,
            replace_warning: bool = False,
            tab_color: str | tuple = None,

            # Section. String Format
            width=None,
            height=None,
            font: str | tuple = None,
            font_name: str = None,
            font_size: int = None,
            font_color: str | tuple = None,
            italic: bool = None,
            bold: bool = None,
            underline: bool = None,
            strikeout: bool = None,
            number_format: str = None,
            align: str | list = None,
            merge: bool = None,
            wrap: bool = None,
            border: str | list = None,
            fill: str | tuple | list = None,
            fill_pattern: str = None,
            fill_fg: str | tuple = None,
            fill_bg: str | tuple = None,
            color_scale: str = None,
            gridlines: bool = None,
            group: bool = None,
            ungroup: bool = None,
            characters_range: tuple = None,
            characters_split: str = None,
            split_picks: str | list = None,
            characters_format: str | dict = None,
            appendix: bool = False,

            # Section. special/personalize format
            index_merge: dict = None,
            header_wrap: bool = None,
            design: str = None,
            style: str | list = None,
            df_format: dict = None,
            cd_style: str | list = None,
            cd_format: list | dict = None,
            config: dict = None,

            # Section. img
            img_left: float = None,
            img_top: float = None,
            img_width: float = None,
            img_height: float = None,

            # Section. hyperlink
            goto: str = None,

            mode: str = None,
            debug: str | bool = None,
            debug_file: str | bool = None,
    ) -> None:
        if debug or debug_file:
            self.reconfigure_logger(debug=debug, debug_file=debug_file)

        self.logger.info("")
        self.logger.info(">" * 30)
        self.logger.info(">>>>>>> LOG FOR PUTXL  <<<<<<<")
        self.logger.info(">" * 30)
        self.logger.info(
            f"> CONTENT: {content if isinstance(content, str) or isinstance(content, list) else 'DataFrame with Size of: ' + str(content.shape)}")
        self.logger.info(f"> SHEET_NAME: {sheet_name}")
        self.logger.info(f"> CELL: {cell}")
        self.logger.info("> LOG ACTIVATED - INFO LEVEL")
        self.logger.debug("> LOG ACTIVATED - DEBUG LEVEL")

        # For Framepro objects
        if hasattr(content, 'df'):
            content = content.df

        # For FramexlWriter Objects
        if isinstance(content, FramexlWriter):
            cell = content.start_cell
            index = content.index_bool
            header = content.header_bool
            content = content.content

        # If content's columns is reachable
        if hasattr(content, 'columns'):
            for col in content.columns:
                content.loc[:, col] = content[col].apply(lambda x: str(x) if isinstance(x, tuple) else x)

        # Sheetreplace? If a sheet_name is specified, then override the current sheet
        ################################
        replace_type = self.alwaysreplace if self.alwaysreplace else replace

        if sheet_name and sheet_name != self.ws.name:
            if sheet_name in [sheet.name for sheet in self.wb.sheets]:
                self.ws = self.wb.sheets[sheet_name]
            else:
                self.ws = self.wb.sheets.add(after=self.wb.sheets.count)
                self.ws.name = sheet_name

        # If sheetreplace or replace is specified, then delete the old sheet and create a new one
        ################################
        if sheetreplace or replace_type == 'sheet':
            _sheetmap = {sheet.index: sheet.name for sheet in self.wb.sheets}
            original_index = self.ws.index
            original_name = self.ws.name
            total_count = self.wb.sheets.count
            self.info_section_lv1("SECTION: sheetreplace or replace_type")
            self.logger.info(
                f"Replacing sheet **!'{self.ws.name}'**: [sheetreplace] is declared as **True**, [alwaysreplace] for PutxlSet is declared as **{self.alwaysreplace}**")
            self.logger.info(
                f"In the workbook, total sheets number is **{total_count}**, while original index is **{original_index}**")

            if original_index == total_count:
                new_sheet = self.wb.sheets.add(after=self.wb.sheets[_sheetmap[original_index]])
                self.logger.info(
                    f"Sheet <is> the last sheet, new sheet added after the sheet **!'{_sheetmap[original_index]}'**")
            else:
                new_sheet = self.wb.sheets.add(before=self.wb.sheets[_sheetmap[original_index + 1]])
                self.logger.info(
                    f"Sheet <is not> the last sheet, new sheet added before the sheet **!'{_sheetmap[original_index + 1]}'**")

            self.ws.delete()
            new_sheet.name = original_name
            self.ws = new_sheet

        # Pre-Cleaning and content type parse: (1) transfer FramePro to dataframe; (2) change tuple cells to str
        ################################
        # For img objects
        if mode == 'img':
            if not isinstance(content, str):
                raise ValueError('Please use a file_path for an image when declaring mode <img>')
            self.ws.pictures.add(
                content,
                left=self.ws.range(cell).left + img_left,
                top=self.ws.range(cell).top + img_top,
                width=img_width,
                height=img_height
            )
            export_notice_name = self.wb.name
            export_notice_name = export_notice_name.replace('.xlsx', '')[0:35] + ' (...) .xlsx' if len(
                export_notice_name) > 36 else export_notice_name
            print(
                f"Image <<{content}>> successfully exported to <<{export_notice_name}>>, worksheet <<{self.ws.name}>> at cell {cell}")
            return

        # left, top, width, height
        # ps.ws.pictures.add(
        #     file_path,
        #     left=ps.ws.range(image_cell).left + 2,
        #     top=ps.ws.range(image_cell).top + 2,
        #     width=80,
        #     height=80
        # )
        # print(f'{upi} image successfully loaded')

        # Declare IO Object
        ################################
        self.info_section_lv1("SECTION: content (i.e. IO object) declaration")
        string_format_tag = False
        if isinstance(content, str):
            self.logger.info(f"Validation 1: [content] **{content}** is passed as a valid str type object")
            self.logger.info(
                f"Validation 2: [content] **{content}** value will lead to [CellPro(content)] taking the value of **{is_cellpro_valid(content)}**")

            if is_cellpro_valid(content) and mode != 'text':
                io = CellxlWriter(cell=content)
                self.logger.info(f"Passed <Cell>: updating sheet <{self.ws.name}> [content] **{content}** format")
                self.next_cell_down = CellPro(CellPro(io.range_cell).cell_stop).offset(1, 0)
                self.next_cell_right = CellPro(CellPro(io.range_cell).cell_stop).offset(0, 1)

            else:
                if mode == 'link':   # For hyperlink
                    if goto is not None and goto not in [sheet.name for sheet in self.wb.sheets]:
                        raise ValueError('Go-to sheet does not exist. Please create first.')
                    else:
                        content = f"=HYPERLINK(\"#'{goto}'!A1\", \"{content}\")"

                io = StringxlWriter(text=content, cell=cell)
                # Note: start_cell is named intentional to be consistent with DF mode and may refer to a cell range
                self.logger.info(
                    f"Passed <Text>: filling in sheet <{self.ws.name}> [content] **{io.content}** into **{io.range_cell}** plus any other format settings ... ")
                self.io = io
                self.ws.range(io.range_cell).value = io.content
                self.next_cell_down = CellPro(CellPro(io.range_cell).cell_stop).offset(1, 0)
                self.next_cell_right = CellPro(CellPro(io.range_cell).cell_stop).offset(0, 1)

            string_format_tag = True

        elif isinstance(content, pandas.DataFrame):
            self.logger.info(f"Validation: [content] type of **{type(content)}** object is passed")
            io = FramexlWriter(frame=content, cell=cell, index=index, header=header, debug=self.debug, debug_file=self.debug_file)
            self.logger.info(
                f"Passed <Frame>: exporting to sheet <{self.ws.name}> [content] frame with size of **{str(content.shape)}** into **{io.start_cell}** plus any other format settings ... ")
            self.ws.range(io.start_cell).value = io.content
            self.io = io
            self.next_cell_down = CellPro(io.bottom_left_cell).offset(1, 0)
            self.next_cell_right = CellPro(io.top_right_cell).offset(0, 1)

        else:
            raise ValueError(f'Invalid type for parameter [content] as {type(content)} is passed, only takes either str (for cell/text to fill in) or dataframe-like objects.')

        # Format the sheet (Shelley, Li)
        ################################
        '''
        Extra Format (not in the group of format parameters): highlight area in existing-content excel
        This is embedded and will be triggered automatically if not replacing sheet 
        '''
        if replace_warning:
            match_dict = {
                'top': self.io.range_top_empty_checker,
                'bottom': self.io.range_bottom_empty_checker,
                'left': self.io.range_left_empty_checker,
                'right': self.io.range_right_empty_checker
            }
            for direction in list(match_dict.keys()):
                if is_range_filled(self.ws, match_dict[direction]):
                    RangeOperator(self.ws.range(self.io.range_all)).format(border=[direction, 'thicker', '#FF0000'], debug=debug)

        if tab_color:
            self.info_section_lv1("SECTION: tab_color")
            paint_tab = color_to_int(tab_color)
            self.logger.info(
                f"Setting sheet <{self.ws.name}> tab color to **{tab_color}**, the value was transformed into int **{paint_tab}**")
            self.ws.api.Tab.Color = paint_tab

        if design:
            self.info_section_lv1("SECTION: design")
            message_init_design = "The design argument passed with look up values from the dict in excel_table_mydesign.py file in the pandaspro package. Both pre-defined style and cd rules can be passed through 1 design"
            self.logger.info(message_init_design)
            self.logger.info("A str is expected to be used as the lookup key")

            '''
            SPECIAL DESIGN: _index as suffix for design argument:
            -----------------------------------------------------------
            For index_merge, add the _index to the selected design like: wbblue_index(indexname, columnnames)
            This will add index_merge(level=..., columns=...) to the style keys

            For example
            >>> wbblue_index(PGs) === index_merge(level=PGs)
            '''
            from pandaspro.user_config.excel_table_mydesign import excel_export_mydesign as local_design
            if re.fullmatch(r'(.*)_index\(([^,]+),?\s*(.*)\)', design):
                match = re.fullmatch(r'(.*)_index\(([^,]+),?\s*(.*)\)', design)
                design = match.group(1)
                index_key = match.group(2)
                index_columns = match.group(3)
                design_style = local_design[design]['style'] + f"; index_merge({index_key},{index_columns})"
                self.info_section_lv2("Sub-section: _index as suffix for design argument")
                self.logger.info(
                    f"Recognized [design] of **{design}**, with extra style of **{local_design[design]['style']}** and added **index_merge({index_key}, {index_columns})** ")
            else:
                design_style = local_design[design]['style']
                self.logger.info(f"Recognized [design] of **{design}**, with extra style of **{design_style}**")

            design_config = local_design[design]['config']
            design_config_shorten_version = {key: design_config[key] for key in list(design_config.keys())[:3]}
            self.logger.info(
                f"Recognized [design] of **{design}**, with extra config of (shortened, use debug level to view all) **{design_config_shorten_version}**")
            self.logger.debug(f"Full-length design_config is **{design_config}**")

            design_cd = local_design[design]['cd']
            self.logger.info(f"Recognized [design] of **{design}**, with extra style of **{design_cd}**")

            message_warning_design = "Note that the design will not override, but instead added to the style, cd_style and config arguments you passed. And it will take effect before style, cd_style, ... which further means it could be overwritten by customized claimed arguments"
            self.logger.info(message_warning_design)
            if style:
                style = ";".join([design_style, style])
            else:
                style = design_style

            if config:
                config = config.update(design_config)
            else:
                config = design_config

            if cd_style:
                cd_style = ";".join([design_cd, cd_style])
            else:
                cd_style = design_cd

        '''
        For index_merge para, the accepted dict only accepts two keys:
        1. level: for which level of the index to be set as merge benchmark
        2. columns: for which columns should apply the merge according to the benchmark index
        
        columns can either be a list or a str, and power-wildcard is embedded when using str:
        >>> ['grade', 'staff_id', 'age']
        >>> '* Total' 
        # this will match all columns in the dataframe ends with Total
        
        Example: {'level': 'cmu_dept', 'columns': '*Total'}
        '''
        if index_merge:
            self.info_section_lv1("SECTION: index_merge")
            self.logger.info(f"[index_merge] is taking the value of **{index_merge}**")
            self.logger.info(f"Parsing into ...")
            for key, local_range in io.range_index_merge_inputs(**index_merge).items():
                self.logger.info(f"key: {key}, local_range: {local_range}")
                RangeOperator(self.ws.range(local_range)).format(merge=True, wrap=True, debug=debug)

        if header_wrap:
            RangeOperator(self.ws.range(io.range_header)).format(wrap=True, debug=debug)

        '''
        apply_df_format: the main function to add format to ranges
        This parameter will take a dict which uses:
        (1) format prompt key words as the keys
        (2) a list of range key words, which may be just a str term (attribute) ... 
            or a cpdFramexl object 

        >>> ... df_format={'msblue80': 'header'}
        >>> ... df_format={'msblue80': cpdFramexl(name='index_merge_inputs', level='cmu_dept_major', columns=['age', 'salary']}
        >>> ... df_format={'blued25; font_color=white': 'columns(['a','b'], header=only)'}

        NOTE! You must specify the kwargs' paras when declaring, like name=, c=, level=, otherwise will be error
        '''
        # Format with defined rules using a Dict
        def apply_df_format(localinput_format, style_name=None):
            i = 0
            for rule, rangeinput in localinput_format.items():
                # Parse the format to a dict, passed to the .format for RangeOperator
                # parse_format_rule is taken from _xlwings module
                self.logger.info("")
                self.logger.info(f"# df_format sub: {style_name} - Number {i + 1} df_format")
                self.logger.info(f"#" * 1 + ' ' + '-' * 45)
                self.logger.info(f"Viewing: key [rule] = **{rule}**, value [rangeinput] = **{rangeinput}**")
                self.logger.info(f"(1) Parsing the key [rule]")
                self.logger.debug(f"Method parse_format_rule is called ...")
                print(parse_header_rule(rule), rule)
                rule_extracted = parse_header_rule(rule)['rule_extracted']
                additional_header_rule = parse_header_rule(rule)['additional_header']
                merge_add_top_title = parse_header_rule(rule)['merge_add_top_title']
                format_kwargs = parse_format_rule(rule_extracted)
                self.logger.info(f"Parsed result: [format_kwargs] = **{format_kwargs}**")

                # Declare range as list/cpdFramexl Object
                def _declare_ranges(local_input):
                    self.logger.debug("_declare_ranges can detail with: str/list/cpdFramexl objects")
                    if isinstance(local_input, str):
                        parsedlist = [local_input]
                        cpdframexl_dict = None
                        self.logger.debug(f"<str> local_input **{local_input}** detected")

                    elif isinstance(local_input, list):
                        parsedlist = local_input
                        cpdframexl_dict = None
                        self.logger.debug(f"<list> local_input **{local_input}** detected")

                    elif isinstance(local_input, cpdFramexl):
                        parsedlist = None
                        cpdframexl_dict = getattr(io, 'range' + local_input.name)(**local_input.paras)
                        self.logger.debug(f"<cpdFramexl> local_input **{local_input}** detected")
                        self.logger.debug(f"[local_input] as cpdFramexl: local_input.name = **{local_input.name}**")
                        self.logger.debug(f"[local_input] as cpdFramexl: local_input.paras = **{local_input.paras}**")

                    else:
                        raise ValueError('Unsupported type in df_format dict values')

                    return parsedlist, cpdframexl_dict

                self.logger.info(f"(2) Parsing the value [rangeinput]")
                self.logger.debug(f"Method _declare_ranges is called ...")
                ioranges, dict_from_cpdframexl = _declare_ranges(rangeinput)
                self.logger.info(f"Parsed 1st result: [ioranges] = **{ioranges}**")
                self.logger.info(f"Parsed 2nd result: [dict_from_cpdframexl] = **{dict_from_cpdframexl}**")

                if ioranges:
                    self.logger.info("")
                    self.logger.info(f"\t[ioranges] - In total there are **{len(ioranges)}** ranges to be parsed")
                    j = 0
                    for each_range in ioranges:
                        self.logger.info(f"\t\t{j + 1}. [each_range] = **{each_range}**")
                        # Parse the input string as method name + kwargs
                        self.logger.debug(f"\t\tMethod parse_method is called ...")
                        range_affix, method_kwargs = parse_method(each_range)[0], parse_method(each_range)[1]
                        self.logger.info(f"\t\tParsed 1st result: [range_affix] = **{range_affix}**")
                        self.logger.info(f"\t\tParsed 2nd result: [method_kwargs] = **{method_kwargs}**")

                        attr_method = getattr(io, 'range_' + range_affix)
                        if callable(attr_method):
                            range_cells = attr_method(**method_kwargs)
                        else:
                            range_cells = attr_method
                        self.logger.info(f"\t\tParsed [range_cells] from the two results above: [range_cells] = **{range_cells}**")

                        if isinstance(range_cells, dict):
                            self.logger.info(f"\t\t[range_cells] is dict type, looping through items to apply [format_kwargs] **{format_kwargs}**")
                            for range_key, range_content in range_cells.items():
                                RangeOperator(self.ws.range(range_content)).format(**format_kwargs, debug=debug)
                        elif isinstance(range_cells, str) and range_cells != '' and range_cells != 'N/A':
                            self.logger.info(f"\t\t[range_cells] is str type, apply [format_kwargs] **{format_kwargs}**")
                            if additional_header_rule is not None:
                                if additional_header_rule == 'merge_up':
                                    updated_range_cells = CellPro(range_cells).offset(-1, 0).resize_h(2).cell
                                    self.logger.info(f"\t\t[merge_up] is detected, this is for header style, the updated range is **{updated_range_cells}**")
                                    RangeOperator(self.ws.range(updated_range_cells)).format(**format_kwargs, debug=debug)
                                    RangeOperator(self.ws.range(updated_range_cells)).format(**format_kwargs)
                                elif additional_header_rule == 'merge_add_top':
                                    updated_range_cells = CellPro(range_cells).offset(-1, 0).cell
                                    self.logger.info(f"\t\t[merge_add_top] is detected, this is for header style, the updated range is **{updated_range_cells}**")
                                    RangeOperator(merge_add_top_title, cell=self.ws.range(updated_range_cells)).format(**format_kwargs, debug=debug)
                                    RangeOperator(self.ws.range(updated_range_cells)).format(**format_kwargs)
                            else:
                                RangeOperator(self.ws.range(range_cells)).format(**format_kwargs, debug=debug)
                        elif range_cells == '' or range_cells == 'N/A':
                            self.logger.info(f"\t\t[range_cells] is empty('' or 'N/A'), no actions")
                        else:
                            raise ValueError(
                                'Invalid Parsed Range Cells from [range_affix] and [method_kwargs]: check <attr_method>')
                        j += 1
                    self.logger.info(f"\t[ioranges] - END")

                if dict_from_cpdframexl:
                    self.logger.info("")
                    self.logger.info(
                        f"[dict_from_cpdframexl] - With length of **{len(dict_from_cpdframexl)}**, formatting [range_content] in a loop")
                    k = 0
                    for range_key, range_content in dict_from_cpdframexl.items():
                        self.logger.info(f"\t{k + 1}. [range_content] = **{range_content}**")
                        RangeOperator(self.ws.range(range_content)).format(**format_kwargs, debug=debug)
                        k += 1
                    self.logger.info(f"\t[dict_from_cpdframexl] - END")

                i += 1

        '''
        style: the main parameter to add pre-defined format to core export data ranges (exc. headers and indices)
        use style_sheets command to view pre-defined formats
        '''
        if style:
            self.info_section_lv1("SECTION: style")
            from pandaspro.user_config.style_sheets import style_sheets

            # First parse string to lists
            if isinstance(style, str):
                loop_list = str2list(style)
            elif isinstance(style, list):
                loop_list = style
            else:
                raise ValueError('Invalid object for style parameter, only str or list accepted')
            self.logger.info(f"[style] argument is automatically parsed into a loop_list: **{loop_list}**")

            # Reorder the items in loop: first do the others, then come to index_merge (merge after fill color, etc.)
            # Allow to pass str like style = "index_merge(..., ...)"
            checked_dict = {}
            for element in loop_list:
                if element in style_sheets:
                    checked_dict[element] = element
                elif re.match(r'index_merge\(([^,]+),?\s*(.*)\)', element):
                    checked_dict['index_merge'] = element
                else:
                    raise ValueError(f'Specified style {element} not in style sheets')

            checked_list = []
            for key in style_sheets.keys():
                if key in checked_dict.keys():
                    checked_list.append(checked_dict[key])

            # Loop and apply style by checking the style py module
            for each_style in checked_list:
                self.info_section_lv2(f"Sub-section: {each_style}")
                self.logger.debug(f"Validation of index_merge: **{each_style}** vs. index_merge\(([^,]+),?\s*(.*)\)")
                self.logger.debug(f"The [apply_style] var will be checking **{each_style}** from <style_sheets>, check style_sheets.py under user_config directory")
                match = re.match(r'index_merge\(([^,]+),?\s*(.*)\)', each_style)
                self.logger.debug(f"Validation result: **{match}**")
                if match:
                    index_name = match.group(1)
                    columns = match.group(2) if match.group(2) != '' else 'None'
                    content_border = style_sheets['index_merge']['border=outer_thick']
                    content_border[1] = content_border[1].replace('__index__', index_name)
                    style_sheets['index_merge']['merge'] = style_sheets['index_merge']['merge'].replace(
                        '__index__', index_name).replace('__columns__', columns)
                    apply_style = style_sheets['index_merge']
                    self.logger.info("[apply_style] is the var passed to apply_df_format method")
                    self.logger.info(f"As index_merge <is> detected, [apply_style] is taking value **{apply_style}**")
                else:
                    apply_style = style_sheets[each_style]
                    self.logger.info(f"As index_merge <is not> detected, [apply_style] is taking value **{apply_style}**")

                apply_df_format(apply_style, each_style)

        '''
        For config para, the accepted dict must use column/index name as keys
        The direct value follow each column/index name must be a dict, 
        and there must be readable keys in it.

        Currently support: 
        1. width
        2. number_format

        For example:
        >>> {
        >>>     'staff id': {'width': 24, 'color': '#00FFFF'},
        >>>     'age': {'width': 15}
        >>>     'salary': {'width': 30, 'haligh': 'left'}
        >>> }
        '''
        if config:
            self.info_section_lv1("SECTION: config")
            self.logger.info(
                f"[config] is taking the value of a dict with length of **{len(config)}**, view details in debug level")
            self.logger.debug(f"Passed [config] argument value: **{config}**")
            for name, setting in config.items():
                if name in io.columns_with_indexnames:
                    self.debug_section_lv2(f"{name}")
                    format_update = {k: v for k, v in setting.items() if not pd.isna(v)}
                    self.logger.debug(
                        f"Adjusting [{name}]: 01 - from config file read format setting: **{format_update}**")
                    self.logger.debug(
                        f"Adjust [{name}]: 02 - range is analyzed as: **{self.ws.range(io.range_columns(name, header=True))}**")
                    RangeOperator(self.ws.range(io.range_columns(name))).format(**format_update, debug=debug)

        if df_format:
            self.info_section_lv1(f"df_format")
            self.logger.info(f"A length **{len(df_format)}** dict is passed to [df_format]")
            apply_df_format(df_format)

        '''
        cd_format: the main function to add format to core export data ranges (exc. headers and indices)
        This parameter will take a dict which allows only three keys (and applyto maybe omitted)
        (refer to the module _cdformat on the class design: _cdformat >> _framewriter.range_cdformat >> _putexcel.PutxlSet.putxl)

        key 1: column = indicating the conditional formatting columns
        key 2: rules = a dict with formatting rules (only based on the column above, like inlist, value equals to, etc.)
        key 3: applyto = where to apply, whether column itself or the whole dataframe, or, several selected columns     
        (default = self)

        >>> ... cd_format={'column': 'age', 'rules': {...}}
        >>> ... cd_format={'column': 'grade', 'rules': {'GA':'#FF0000'}, 'applyto': 'self'}
        >>> ... cd_format={'column': 'grade', 'rules': {'rule1':{'r':...(pd.Series), 'f':...}}, 'applyto': 'self'}
        '''
        # Conditional Format (1 column based)
        # This function will always check the type of the argument that is passed to this parameter
        # If a list type is detected, then use loop to loop through the list and call the cd_paint function many times
        # So whether dict or a list of dictionaries, the format has to comply with standard cpd cd dict format
        # .. which you may refer to the comments before "if cd" line
        def apply_cd_format(input_cd, cd_name=None):
            def cd_paint(input_cd_instance):
                self.logger.info("Parsing the dict [input_cd] with <io> and <range_cdformat> instance method")
                for key, value in input_cd_instance.items():
                    self.logger.info(f"Dict [**{key}**]: **{value}**")
                cleaned_rules = io.range_cdformat(**input_cd_instance)
                self.logger.info(f"This will result in a **cleaned dict with multi sub-dicts: [cleaned_rules] with {len(cleaned_rules)}**")

                # Work with the cleaned_rules to adjust the cell formats in Excel with RangeOperator
                for rulename, lc_content in cleaned_rules.items():
                    self.logger.info(f"")
                    self.logger.info(f"\t[rulename] = **{rulename}**")
                    cellrange = lc_content['cellrange']
                    cd_format_rule = lc_content['format']
                    self.logger.info(f"\t[cellrange] = **{cellrange}**")
                    self.logger.info(f"\t[cd_format_rule] = **{cd_format_rule}**")

                    if cellrange == 'no cells':
                        self.logger.info(f"\t.. because [cellrange] is taking value <no cells>, no actions needed")
                    else:
                        # Parse the cd_format_rule to a dict, as **kwargs to be passed to the .format for RangeOperator
                        # parse_format_rule is taken from _xlwings module
                        self.logger.info(f"\tParsing the [cd_format_rule] with <parse_format_rule> method from range_operator.py under io.excel directory")
                        cd_format_kwargs = parse_format_rule(cd_format_rule)
                        self.logger.info(f"\tResult: [cd_format_kwargs] = **{cd_format_kwargs}**")

                        if len(cellrange) <= 45:
                            self.logger.info(f"\tDirectly apply - length of [cellrange] is **{len(cellrange)}**, no larger than 45")
                            self.logger.info(f"\t--> Applying to range: **{cellrange}**")
                            RangeOperator(self.ws.range(cellrange)).format(debug=debug, **cd_format_kwargs)
                        else:
                            self.logger.info(f"\tCombine cells first - length of [cellrange] is **{len(cellrange)}**, larger than 45")
                            # Here is the combine function
                            '''
                            cell_range_combine method from _utils
                            takes a list and returns a dict (from 1 dimension to 2 dimensions)
                            
                            Previously like:
                            'B2,C2,D2,E2,F2,G2,H2,I2,J2,K2,L2,M2,O2,B3' 
                            
                            After combine will be:
                            {2: ['B2:M2', 'O2:O2'], 3: ['B3:B3']}
                            '''
                            cellrange_dict = cell_combine_by_column(cellrange.split(','))
                            self.logger.info(f"\tCombined into 1 dict [cell_range_combine] with length of **{len(cellrange_dict)}**")
                            self.logger.info(f"\t--> Applying to range:")

                            for key, range_list in cellrange_dict.items():
                                self.logger.info(f"\t\tRange ID: [column key] = **{key}**")
                                for combined_range in range_list:
                                    self.logger.info(f"\t\tRange Content: [combined_range] = **{combined_range}**")
                                    RangeOperator(self.ws.range(combined_range)).format(debug=debug, **cd_format_kwargs)

            # Decide if cd_format is a dict or not
            if isinstance(input_cd, dict):
                self.logger.info("")
                self.logger.info(f"# Only 1 [rule] as dict-type is passed to <apply_cd_format>")
                self.logger.info(f"#" * 1 + ' ' + '-' * 45)
                self.logger.info(f"This cd_format dict is built from keys: **{input_cd.keys()}**")

                cd_paint(input_cd)

            if isinstance(input_cd, list):
                l = 0
                for rule in input_cd:
                    self.logger.info("")
                    self.logger.info(f"# cd_format sub: {cd_name} - Number {l + 1}")
                    self.logger.info(f"#" * 1 + ' ' + '-' * 45)
                    self.logger.info(f"This cd_format dict is built from keys: **{rule.keys()}**")
                    cd_paint(rule)
                    l += 1

        '''
        cd: the main parameter to add pre-defined conditional formatting to core export data ranges (exc. headers and indices)
        use cd_sheets command to view pre-defined formats
        '''
        if cd_style:
            self.info_section_lv1("SECTION: cd_style")
            from pandaspro.user_config.cd_sheets import cd_sheets

            # First parse string to lists
            if isinstance(cd_style, str):
                loop_list = str2list(cd_style)
            elif isinstance(cd_style, list):
                loop_list = cd_style
            else:
                raise ValueError('Invalid object for cd parameter, only str or list accepted')
            self.logger.info(f"[cd_style] argument is automatically parsed into a loop_list: **{loop_list}**")
            self.logger.info(f"Note that some of these styles having multiple dictionaries as target reference, while others only refers to one")

            # Loop and apply cd by checking the cd py module
            for each_cd in loop_list:
                apply_cd = cd_sheets[each_cd]
                self.info_section_lv2(f"Sub-section: {each_cd}")
                self.logger.debug(f"The [apply_cd] var will be checking **{each_cd}** from <cd_sheets>, check cd_sheets.py under user_config directory")
                self.logger.debug(f"Checked [apply_cd]: **{apply_cd}** is a **{type(apply_cd)}**")
                apply_cd_format(apply_cd, each_cd)

                # To Delete, apply_cd_format can take dictionaries/lists
                # ------------------------------------------
                # if isinstance(apply_cd, list):
                #     for each_cd_sub in apply_cd:
                #         apply_cd_format(each_cd_sub)
                # elif isinstance(apply_cd, dict):
                #     apply_cd_format(apply_cd)
                # else:
                #     raise ValueError('Invalid type for [apply_cd]')

        if cd_format:
            self.info_section_lv1(f"cd_format")
            self.logger.info(f"A length **{len(cd_format)}** with type of **{type(cd_format)}** is passed to [df_format]")
            apply_cd_format(cd_format)

        if string_format_tag:
            RangeOperator(self.ws.range(io.range_cell)).format(
                width=width,
                height=height,
                font=font,
                font_name=font_name,
                font_size=font_size,
                font_color=font_color,
                italic=italic,
                bold=bold,
                underline=underline,
                strikeout=strikeout,
                number_format=number_format,
                align=align,
                merge=merge,
                wrap=wrap,
                border=border,
                fill=fill,
                fill_pattern=fill_pattern,
                fill_fg=fill_fg,
                fill_bg=fill_bg,
                color_scale=color_scale,
                gridlines=gridlines,
                group=group,
                ungroup=ungroup,
                appendix=appendix,
                debug=debug
            )

            if characters_range and io.iotype == 'cell':
                if not isinstance(characters_range, list) or not len(characters_range) == 2 or characters_format is None:
                    raise ValueError('font_characters_range argument must have the three keys below: start, end, and format')
                if isinstance(characters_format, str):
                    characters_format = parse_format_rule(characters_format)
                for cell in self.ws.range(io.range_cell):
                    if self.ws.range(cell).value is None:
                        continue
                    else:
                        RangeOperator(cell, get_characters=True, get_characters_type='range', start=characters_range[0], end=characters_range[1]).format(**characters_format)

            if characters_split and io.iotype == 'cell':
                if split_picks is None or characters_format is None:
                    raise ValueError('font_characters_range argument must have the three keys below: split, split_picks, and format')
                if isinstance(characters_format, str):
                    characters_format = parse_format_rule(characters_format)

                for cell in self.ws.range(io.range_cell):
                    if self.ws.range(cell).value is None:
                        continue
                    else:
                        RangeOperator(cell, get_characters=True, get_characters_type='split', split=characters_split, split_picks=split_picks).format(**characters_format)

        # Remove Sheet1 if blank and exists (the Default tab) ...
        ################################
        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if 'Sheet1' in current_sheets and is_sheet_empty(self.wb.sheets['Sheet1']):
            self.wb.sheets['Sheet1'].delete()

        self.wb.save()

        # Print Export Success Message to Console ...
        ################################
        export_notice_name = self.wb.name
        export_notice_name = export_notice_name.replace('.xlsx', '')[0:35] + ' (...) .xlsx' if len(
            export_notice_name) > 36 else export_notice_name

        if isinstance(content, str):
            if is_cellpro_valid(content) and mode != 'text':
                print(f"Cell range <<{content}>> successfully updated in <<{export_notice_name}>>, worksheet <<{self.ws.name}>> with declared format")
            else:
                print(f"Text <<{content}>> successfully filled in <<{export_notice_name}>>, worksheet <<{self.ws.name}>> in cell {cell}")

        elif isinstance(content, pandas.DataFrame):
            print(f"Frame with size <<{content.shape}>> successfully exported to <<{export_notice_name}>>, worksheet <<{self.ws.name}>> at cell {cell}")
        # for else, an error should already been thrown in the previous content/io declaration stage

    def tab(self, sheet_name: str, sheetreplace: bool = False, tab_color: str = None) -> None:
        """
        Switches to a specified sheet in the workbook.
        If the sheet does not exist, it creates a new one with the given name.

        Parameters
        ----------
        sheet_name : str
            The name of the sheet to switch to or create.
        sheetreplace: bool
            If true, replace the content in the sheet
        tab_color: str
            Control the tab color
        """
        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if sheet_name in current_sheets:
            sheet = self.wb.sheets[sheet_name]
        else:
            sheet = self.wb.sheets.add(after=self.wb.sheets.count)
            sheet.name = sheet_name
        self.ws = sheet

        # If sheetreplace is specified, then delete the old sheet and create a new one
        ################################
        if sheetreplace:
            _sheetmap = {sheet.index: sheet.name for sheet in self.wb.sheets}
            original_index = self.ws.index
            original_name = self.ws.name
            total_count = self.wb.sheets.count

            if original_index == total_count:
                new_sheet = self.wb.sheets.add(after=self.wb.sheets[_sheetmap[original_index]])
            else:
                new_sheet = self.wb.sheets.add(before=self.wb.sheets[_sheetmap[original_index + 1]])

            self.ws.delete()
            new_sheet.name = original_name
            self.ws = new_sheet

        if tab_color:
            paint_tab = color_to_int(tab_color)
            self.ws.api.Tab.Color = paint_tab

        return

    def close(self):
        self.open_wb.close()


if __name__ == '__main__':

    import pandaspro as cpd
    d = cpd.sysuse_auto
    debuglevel = 'info'
    ps = PutxlSet('temp.xlsx')
    ps.putxl('go back', mode='link', sheet_name='Sheet1', cell='A1', goto='Sheet3', font_color="red", font_size=24)
    # ps.putxl(d, cell='A4', cd_format={'column': 'rep78', 'rules': {1: 'red', 2: 'blue'}, 'applyto': 'self'})
    # ps.putxl('A1:A8', characters_split=" ", split_picks=[4], characters_format="font_color=red")
    # ps.putxl(
    #     r.table_region('AFW'),
    #     cell='A5', index=False,
    #     design='wbblue',
    #     df_format={
    #         'font_size=12': 'all',
    #         'number_format=0.0': 'cspan(s="Overall Rating Average", e="Results")'
    #     },
    #     debug=debuglevel
    # )
    # ps.close()
