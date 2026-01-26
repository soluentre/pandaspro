from pandaspro.core.stringfunc import parse_wild
from pandaspro.io.excel.cdformat import CdFormat
from pandaspro.core.tools.utils import df_with_index_for_mask
from pandaspro.io.cellpro.cellpro import CellPro, index_cell, is_cellpro_valid
import pandas as pd

from pandaspro.utils.cpd_logger import cpdLogger


class CellxlWriter:
    def __init__(
            self,
            cell: str = None,
    ) -> None:
        self.iotype = 'cell'
        self.range_cell = cell


class StringxlWriter:
    def __init__(
            self,
            text: str = None,
            cell: str = None,
    ) -> None:
        self.iotype = 'text'
        self.content = text
        self.range_cell = cell


class DictxlWriter:
    def __init__(
            self,
            mydict: dict = None,
    ) -> None:
        for key, value in mydict.items():
            if not is_cellpro_valid(key):
                raise ValueError(f'{key} is not a valid cellpro key in the keys of the dictionary passed')
            if not isinstance(value, str):
                raise ValueError(f'{value} must be a string in the values of the dictionary passed')
        self.iotype = 'dict'
        self.content = mydict
        self.keys = list(mydict.keys())
        self.values = list(mydict.values())
        self.last_cell = self.keys[-1]


@cpdLogger
class FramexlWriter:
    def __init__(
            self,
            frame,
            cell: str,
            index: bool = False,
            header: bool = True,
            debug: str = None,
            debug_file: str = None,
    ) -> None:
        cellobj = CellPro(cell)
        header_row_count = len(frame.columns.levels) if isinstance(frame.columns, pd.MultiIndex) else 1
        index_column_count = len(frame.index.levels) if isinstance(frame.index, pd.MultiIndex) else 1

        # Calculate the Ranges
        self.rawdata = frame
        frame = pd.DataFrame(frame)
        if header == True and index == True:
            self.export_type = 'htit'
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1] + index_column_count
            xl_header_count, xl_index_count = header_row_count, index_column_count
            export_data = frame
            range_index = cellobj.offset(header_row_count, 0).resize(tr - header_row_count, index_column_count)
            range_indexnames = cellobj.resize(header_row_count, header_row_count)
            range_header = cellobj.offset(0, index_column_count).resize(header_row_count, tc - index_column_count)
        elif header == False and index == True:
            self.export_type = 'hfit'
            tr, tc = frame.shape[0], frame.shape[1] + index_column_count
            xl_header_count, xl_index_count = 0, index_column_count
            export_data = frame.reset_index().to_numpy().tolist()
            range_index = cellobj.resize(tr, index_column_count)
            header_row_count = 0
            range_indexnames = 'N/A'
            range_header = 'N/A'
        elif header == False and index == False:
            self.export_type = 'hfif'
            tr, tc = frame.shape[0], frame.shape[1]
            xl_header_count, xl_index_count = 0, 0
            export_data = frame.to_numpy().tolist()
            header_row_count = 0
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = 'N/A'
        else:
            self.export_type = 'htif'
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1]
            xl_header_count, xl_index_count = header_row_count, 0
            if isinstance(frame.columns, pd.MultiIndex):
                column_export = [list(lst) for lst in list(zip(*frame.columns.values))]
            else:
                column_export = [frame.columns.to_list()]
            # noinspection PyTypeChecker
            export_data = column_export + frame.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = cellobj.resize(header_row_count, tc)

        # Calculate the Map
        dfmapstart = cellobj.offset(xl_header_count, 0)
        dfmap = df_with_index_for_mask(self.rawdata, force=index).copy()
        dfmap = dfmap.astype(str)

        for dfmap_index in range(len(dfmap)):
            for j, col in enumerate(dfmap.columns):
                dfmap.iloc[dfmap_index, j] = dfmapstart.offset(dfmap_index, j).cell

        self.iotype = 'df'
        self.columns_with_indexnames = self.rawdata.reset_index().columns
        self.columns = self.rawdata.columns
        self.content = export_data
        self.start_cell = cell
        self.index_bool = index
        self.header_bool = header
        self.tr = tr
        self.tc = tc
        self.header_row_count = header_row_count
        self.index_column_count = index_column_count
        self.data_height = self.tr - self.header_row_count
        self.data_width = self.tc - self.index_column_count

        # data corners - cellpros
        self.inner_start_cellobj = cellobj.offset(xl_header_count, xl_index_count)
        self.inner_start_cell = self.inner_start_cellobj.cell
        self.top_right_cell = cellobj.offset(0, self.tc - 1).cell
        self.bottom_left_cell = cellobj.offset(self.tr - 1, 0).cell
        self.end_cell = cellobj.offset(self.tr - 1, self.tc - 1).cell

        # ranges
        self.range_all = cell + ':' + self.end_cell
        self.range_data = self.inner_start_cell + ':' + self.end_cell
        self.range_index = range_index.cell if range_index != 'N/A' else 'N/A'
        self.range_index_outer = CellPro(self.start_cell).resize(self.tr, self.index_column_count).cell if range_index != 'N/A' else 'N/A'
        self.range_header = range_header.cell if range_header != 'N/A' else 'N/A'
        self.range_header_outer = CellPro(self.start_cell).resize(self.header_row_count, self.tc).cell if range_header != 'N/A' else 'N/A'
        self.range_indexnames = range_indexnames.cell if range_indexnames != 'N/A' else 'N/A'

        # format relevant
        self.dfmap = dfmap
        self.cols_index_merge = None

        # Conditional Formatting
        self.cd_dfmap_1col = None
        self.cd_cellrange_1col = None

        # Special - Checker for sheetreplace
        self.range_top_empty_checker = CellPro(self.start_cell).offset(-1, 0).resize(1, self.tc).cell if CellPro(self.start_cell).cell_index[0] != 1 else None
        self.range_bottom_empty_checker = CellPro(self.bottom_left_cell).offset(1, 0).resize(1, self.tc).cell if CellPro(self.bottom_left_cell).cell_index[0] != 1 else None
        self.range_left_empty_checker = CellPro(self.start_cell).offset(0, -1).resize(self.tr, 1).cell if CellPro(self.start_cell).cell_index[1] != 1 else None
        self.range_right_empty_checker = CellPro(self.top_right_cell).offset(0, 1).resize(self.tr, 1).cell if CellPro(self.top_right_cell).cell_index[0] != 1 else None

        # Special - first/second from bottom/right
        self.range_bottom1 = CellPro(self.bottom_left_cell).resize(1, self.tc).cell
        self.range_right1 = CellPro(self.top_right_cell).resize(self.tr, 1).cell

        # Debug
        self.debug = debug
        self.debug_file = debug_file
        self.logger = None
        self.debug_section_spec_start = None

    def range_multiindex_header_merge(self) -> dict:
        """
        Calculate merge ranges for MultiIndex columns header.
        Returns a dict with merge ranges for each level of the header.
        """
        if not isinstance(self.columns, pd.MultiIndex):
            return {}
        
        result_dict = {}
        num_levels = len(self.columns.levels)
        
        # For each level in the MultiIndex
        for level_idx in range(num_levels):
            # Get values at this level
            level_values = [col[level_idx] for col in self.columns]
            
            # Find consecutive same values
            merge_ranges = []
            start_idx = 0
            current_value = level_values[0]
            
            for i in range(1, len(level_values) + 1):
                # Check if we've reached the end or value changed
                if i == len(level_values) or level_values[i] != current_value:
                    # Only merge if span > 1
                    if i - start_idx > 1:
                        # Calculate the cell range
                        start_cell = CellPro(self.start_cell).offset(level_idx, self.index_column_count + start_idx)
                        merge_range = start_cell.resize(1, i - start_idx).cell
                        merge_ranges.append(merge_range)
                    
                    # Move to next group
                    if i < len(level_values):
                        start_idx = i
                        current_value = level_values[i]
            
            if merge_ranges:
                result_dict[f'level_{level_idx}'] = merge_ranges
        
        return result_dict

    def range_multiindex_columns_first_columns(self) -> list:
        """
        Get the first column of each top-level group in MultiIndex columns.
        Returns a list of cell ranges for the first column of each group.
        Useful for adding vertical borders to separate column groups.
        """
        if not isinstance(self.columns, pd.MultiIndex):
            return []
        
        result_ranges = []
        
        # Get values at the first level (top level)
        level_0_values = [col[0] for col in self.columns]
        
        # Find where values change (start of each new group)
        prev_value = None
        for i, value in enumerate(level_0_values):
            if value != prev_value:
                # This is the start of a new group
                # Calculate the cell range for this column (entire column including header and data)
                start_cell = CellPro(self.start_cell).offset(0, self.index_column_count + i)
                column_range = start_cell.resize(self.tr, 1).cell
                result_ranges.append(column_range)
                prev_value = value
        
        return result_ranges

    def range_index_sections_by_value(self, level: str, value: str) -> list:
        """
        Find all row ranges where a specific index level has a specific value.
        Useful for conditional formatting based on index values.
        
        :param level: Index level name
        :param value: Value to match (supports wildcards like '*Subtotal')
        :return: List of cell ranges
        """
        if level not in self.rawdata.index.names:
            raise ValueError(f"Level {level} not found in index names")
        
        temp = self.rawdata.reset_index()
        
        # Support wildcard matching
        if '*' in value:
            import re
            pattern = value.replace('*', '.*')
            mask = temp[level].astype(str).str.match(pattern)
        else:
            mask = temp[level] == value
        
        # Get row indices where condition is True
        matching_indices = temp[mask].index.tolist()
        
        # Convert to cell ranges
        result_ranges = []
        if matching_indices:
            # Group consecutive indices
            groups = []
            current_group = [matching_indices[0]]
            
            for idx in matching_indices[1:]:
                if idx == current_group[-1] + 1:
                    current_group.append(idx)
                else:
                    groups.append(current_group)
                    current_group = [idx]
            groups.append(current_group)
            
            # Convert each group to a cell range
            for group in groups:
                start_row = group[0] + self.header_row_count
                row_count = len(group)
                start_cell = CellPro(self.start_cell).offset(start_row, 0)
                cell_range = start_cell.resize(row_count, self.tc).cell
                result_ranges.append(cell_range)
        
        return result_ranges
    
    def range_subtotal_rows(self) -> list:
        """
        Find all Subtotal rows in the dataframe.
        Returns a list of cell ranges for each Subtotal row.
        """
        if not isinstance(self.rawdata.index, pd.MultiIndex):
            return []
        
        # Check all index levels for Subtotal
        import re
        result_ranges = []
        temp = self.rawdata.reset_index()
        
        # Find rows containing 'Subtotal' in any index level
        mask = pd.Series([False] * len(temp))
        for level in self.rawdata.index.names:
            if level is not None:
                level_mask = temp[level].astype(str).str.contains('Subtotal', na=False)
                mask = mask | level_mask
        
        matching_indices = temp[mask].index.tolist()
        
        if matching_indices:
            for idx in matching_indices:
                start_row = idx + self.header_row_count
                start_cell = CellPro(self.start_cell).offset(start_row, 0)
                cell_range = start_cell.resize(1, self.tc).cell
                result_ranges.append(cell_range)
        
        return result_ranges
    
    def range_subtotal_columns(self) -> list:
        """
        Find all Subtotal columns in the dataframe.
        Returns a list of cell ranges for each Subtotal column.
        Only returns data area, excluding headers.
        """
        if not isinstance(self.rawdata.columns, pd.MultiIndex):
            return []
        
        result_ranges = []
        
        # Find columns containing 'Subtotal' in any level
        subtotal_cols = []
        for col in self.rawdata.columns:
            if any('Subtotal' in str(level) for level in col):
                subtotal_cols.append(col)
        
        # Convert to cell ranges (only data area, excluding headers)
        for col in subtotal_cols:
            try:
                # get_column_letter_by_name returns position starting from data area (after headers)
                # So we don't need to offset again
                col_letter = self.get_column_letter_by_name(col)
                # col_letter is already at the first data row, just resize for all data rows
                data_row_count = self.rawdata.shape[0]
                cell_range = col_letter.resize(data_row_count, 1).cell
                result_ranges.append(cell_range)
            except:
                continue
        
        return result_ranges

    def get_column_letter_by_indexname(self, levelname):
        if not self.index_bool:
            raise ValueError(
                f'When searching column << {levelname} >>, the name appears in index columns. And system found an error with get_column_letter_by_indexname method because << index = False >> is specified')

        col_count = list(self.rawdata.index.names).index(levelname)
        col_cell = CellPro(self.start_cell).offset(self.header_row_count, col_count)
        return col_cell

    def get_column_letter_by_name(self, colname):
        # Support MultiIndex columns with __ separator
        if isinstance(self.columns, pd.MultiIndex):
            # If colname contains __, split it and convert to tuple
            if isinstance(colname, str) and '__' in colname:
                # Split by __ and convert to tuple
                colname_parts = colname.split('__')
                # Try to find matching column in MultiIndex
                for i, col in enumerate(self.columns):
                    if len(col) == len(colname_parts):
                        # Check if all parts match
                        if all(str(col[j]) == colname_parts[j] for j in range(len(colname_parts))):
                            col_count = i
                            col_cell = self.inner_start_cellobj.offset(0, col_count)
                            return col_cell
                # If no match found, raise error
                raise ValueError(f'Column {colname} not found in MultiIndex columns. Use __ to separate levels.')
            # If colname is already a tuple, use it directly
            elif isinstance(colname, tuple):
                col_count = list(self.columns).index(colname)
                col_cell = self.inner_start_cellobj.offset(0, col_count)
                return col_cell
            else:
                # Try to find as-is (for single-level matching)
                try:
                    col_count = list(self.columns).index(colname)
                    col_cell = self.inner_start_cellobj.offset(0, col_count)
                    return col_cell
                except ValueError:
                    raise ValueError(f'Column {colname} not found. For MultiIndex columns, use __ to separate levels (e.g., "Level1__Level2").')
        else:
            col_count = list(self.columns).index(colname)
            col_cell = self.inner_start_cellobj.offset(0, col_count)
            return col_cell

    def _index_break(self, level: str = None):
        temp = self.rawdata.reset_index()

        def _count_consecutive_values(series):
            return series.groupby((series != series.shift()).cumsum()).size().tolist()

        return _count_consecutive_values(temp[level])

    def range_index_merge_inputs(
            self,
            level: str = None,
            columns: str | list = None
    ) -> dict:
        result_dict = {}

        # Index Column
        merge_start_index = self.get_column_letter_by_indexname(level)
        for localid, rowspan in enumerate(self._index_break(level=level)):
            result_dict[f'indexlevel_{localid}_{rowspan}'] = merge_start_index.resize(rowspan, 1).cell
            merge_start_index = merge_start_index.offset(rowspan, 0)

        # Selected Columns
        if columns:
            # Handle MultiIndex columns
            if isinstance(self.columns, pd.MultiIndex):
                if isinstance(columns, list):
                    self.cols_index_merge = columns
                else:
                    # Create string representation for wildcard matching
                    columns_str_list = ['__'.join(str(x) for x in col) for col in self.columns]
                    matched = parse_wild(columns, columns_str_list)
                    self.cols_index_merge = matched
            else:
                self.cols_index_merge = columns if isinstance(columns, list) else parse_wild(columns, self.columns)
            # print("framewriter cols_index_merge:", self.cols_index_merge)
            # print("columns:", columns, self.columns)
            for index, col in enumerate(self.cols_index_merge):
                merge_start_each = self.get_column_letter_by_name(col)
                for localid, rowspan in enumerate(self._index_break(level=level)):
                    result_dict[f'col{index}_{localid}_{rowspan}'] = merge_start_each.resize(rowspan, 1).cell
                    merge_start_each = merge_start_each.offset(rowspan, 0)

        return result_dict

    def range_index_hsections(self, level: str = None) -> dict:
        if self.range_index is None:
            raise ValueError('index_sections method requires the input dataframe to have an index')
        else:
            result_dict = {'headers': CellPro(self.start_cell).resize(self.header_row_count, self.tc).cell}
            range_start_each = CellPro(self.start_cell).offset(self.header_row_count, 0)
            for localid, rowspan in enumerate(self._index_break(level=level)):
                result_dict[f'section_{localid}_{rowspan}'] = range_start_each.resize(rowspan, self.tc).cell
                range_start_each = range_start_each.offset(rowspan, 0)

        return result_dict

    def range_index_selected_hsection(self, level: str = None, token: str = 'Total') -> str:
        temp = self.rawdata.reset_index()

        def _find_occurrence_details(series, indexname):
            """
            This function finds the first occurrence of a specified token in a pandas Series,
            returns the index of its first appearance, and the count of its consecutive occurrences.
            """
            if indexname in series.values:
                first_occurrence_index = series[series == indexname].index[0]
                # Count the consecutive occurrences starting from the first occurrence index
                count = 1  # Start with 1 for the first occurrence
                for i in range(first_occurrence_index + 1, len(series)):
                    if series.iloc[i] == indexname:
                        count += 1
                    else:
                        break
                return first_occurrence_index, count
            else:
                return None, 0

        go_down_by, local_height = _find_occurrence_details(temp[level], token)
        result = self.get_column_letter_by_indexname(level).offset(go_down_by, 0).resize(local_height, self.tc).cell

        return result

    ''' this is returning the whole level by level ranges in selection '''

    @property
    def range_index_levels(self) -> dict:
        result_dict = {}
        range_start_each = CellPro(self.start_cell)
        for each_index in self.rawdata.index.names:
            result_dict[f'index_{each_index}'] = range_start_each.resize(self.tr, 1).cell
            range_start_each = range_start_each.offset(0, 1)
        return result_dict

    def range_columns(self, c, header=False):
        if isinstance(c, str):
            # For MultiIndex columns, don't use parse_wild on columns_with_indexnames directly
            # Instead, create a list of string representations for matching
            if isinstance(self.columns, pd.MultiIndex):
                # Create string representation of MultiIndex columns using __
                columns_str_list = ['__'.join(str(x) for x in col) for col in self.columns]
                # Also include original index names (filter out None)
                index_names = [name for name in self.rawdata.index.names if name is not None]
                all_searchable = index_names + columns_str_list
                # Try to match using wildcard
                matched = parse_wild(c, all_searchable)
                # Convert back matched string representations to actual column names
                clean_list = []
                for m in matched:
                    if m in self.rawdata.index.names:
                        clean_list.append(m)
                    else:
                        # Find the original tuple column
                        clean_list.append(m)  # Keep as string, will be processed later
            else:
                clean_list = parse_wild(c, self.columns_with_indexnames)
        elif isinstance(c, list):
            clean_list = c
        else:
            raise ValueError('range_columns only accept str/list as inputs')

        result_list = []
        for colname in clean_list:
            # Handle MultiIndex column names with __ separator
            if isinstance(self.columns, pd.MultiIndex) and isinstance(colname, str) and '__' in colname:
                start_range = self.get_column_letter_by_name(colname)
            elif colname in self.columns:
                start_range = self.get_column_letter_by_name(colname)
            elif colname in self.rawdata.index.names:
                start_range = self.get_column_letter_by_indexname(colname)
            else:
                raise ValueError(f'Searching name <<{colname}>> is not in column nor index.names. For MultiIndex columns, use __ to separate levels.')

            below_range = start_range.resize_h(self.tr - self.header_row_count).cell

            # noinspection PySimplifyBooleanCheck
            if header == True:
                below_range = CellPro(below_range).offset(-self.header_row_count, 0).resize_h(self.tr).cell
            if header == 'only':
                below_range = CellPro(below_range).offset(-self.header_row_count, 0).resize_h(
                    self.header_row_count).cell
            result_list.append(below_range)

        return ', '.join(result_list)

    def range_cspan(self, s=None, e=None, c=None, header=False):
        # Declaring starting and ending columns
        if s and e:
            col_index1 = self.get_column_letter_by_name(s).cell_index[1]
            col_index2 = self.get_column_letter_by_name(e).cell_index[1]
            row_index = self.get_column_letter_by_name(s).cell_index[0]

            # Decide the top row cells with min/max - allow invert orders
            top_left_index = min(col_index1, col_index2)
            top_right_index = max(col_index1, col_index2)
            top_left = index_cell(row_index, top_left_index)
            top_right = index_cell(row_index, top_right_index)
            start_range = CellPro(top_left + ':' + top_right)

        # Declaring only 1 column
        elif c:  # Para C: declare column only
            selected_column = self.get_column_letter_by_name(c)
            start_range = selected_column

        else:
            raise ValueError('At least 1 set of Paras: (1) s+e or (2) c must be declared ')

        final = start_range.resize_h(self.tr - self.header_row_count).cell
        # noinspection PySimplifyBooleanCheck
        if header == True:
            final = CellPro(final).offset(-self.header_row_count, 0).resize_h(self.tr).cell
        if header == 'only':
            final = CellPro(final).offset(-self.header_row_count, 0).resize_h(self.header_row_count).cell

        return final

    def range_cdformat(
            self,
            column,
            rules=None,
            applyto='self',
    ):
        mycd = CdFormat(
            df=self.rawdata,
            column=column,
            cd_rules=rules,
            applyto=applyto,
            debug=self.debug,
            debug_file=self.debug_file
        )
        # print(mycd.df.columns, mycd.df_with_index.columns, mycd.column)
        if mycd.col_not_exist:
            cd_cellrange_1col = {'void_rule': {'cellrange': 'no cells', 'format': ''}}
        else:
            apply_columns = mycd.apply
            this_rules_mask = mycd.get_rules_mask()

            # Deprecated?
            # -------------------------------------------
            # cd_dfmap_1col = {}
            # for key, mask_rule in this_rules_mask.items():
            #     cd_dfmap_1col[key] = {}
            #     cd_dfmap_1col[key]['dfmap'] = self.dfmap[mask_rule['mask']][apply_columns]
            #     cd_dfmap_1col[key]['format'] = mask_rule['format']
            # self.cd_dfmap_1col = cd_dfmap_1col

            def _df_to_mystring(df):
                lcarray = df.values.flatten()
                long_string = ','.join([str(value) for value in lcarray])

                self.logger.debug(f'++ \tUsing <_df_to_mystring() method>**')
                self.logger.debug(f'++ \t[flattened lcarray]: a numpy array **{lcarray}**')
                self.logger.debug(
                    f'++ \t[long string]: a string **<{long_string}>** with length **{len(long_string)}**')

                result_string = "no cells" if len(long_string) == 0 else long_string
                return result_string

            cd_cellrange_1col = {}

            self.debug_section_spec_start(
                'Parsing this_rules_mask which is the CdFormat class <get_rules_mask()> method')
            self.logger.debug(f'++ [this_rules_mask]: keys are **{this_rules_mask.keys()}**')

            for key, mask_rule in this_rules_mask.items():
                cd_cellrange_1col[key] = {}
                temp_dfmap = self.dfmap[mask_rule['mask']][apply_columns]

                self.logger.debug("")
                self.logger.debug("")
                self.logger.debug(f'++ \ttemp_dfmap = self.dfmap[mask_rule[mask]][apply_columns]')
                self.logger.debug(
                    f'++ \t[key]: **{key}**, [mask_rule]: a **{type(mask_rule)}** with [mask] and [format], [mask] being **{list(mask_rule["mask"])}**')
                self.logger.debug(f'++ \t[key]: **{key}**, [apply_columns]: **{apply_columns}**')
                self.logger.debug(f'++ \t[temp_dfmap]: a **{type(temp_dfmap)}** with size **{temp_dfmap.shape}**')

                cd_cellrange_1col[key]['cellrange'] = _df_to_mystring(temp_dfmap)
                cd_cellrange_1col[key]['format'] = mask_rule['format']

            self.cd_cellrange_1col = cd_cellrange_1col

        '''
        should be something like ...
        {
            "AFWDE": {
                "cellrange": "B2,C2,D2,E2,F2,G2,H2,I2,J2,K2,L2,M2", 
                "format": "blue"
            },
            "AFWVP": {
                "cellrange": "B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3", 
                "format": "orange"
            },
        }
        '''
        return cd_cellrange_1col


class cpdFramexl:
    def __init__(self, name, **kwargs):
        self.name = name
        self.paras = kwargs

# if __name__ == '__main__':
#     import wbhrdata as wb
#     import pandaspro as cpd
#     data = wb.sob().head(5).p.er
#
#     ps = cpd.PutxlSet('file.xlsx')
#     ps.putxl(data, cell='A1', design='wbblue')
