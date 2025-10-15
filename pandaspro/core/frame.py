import re

import numpy as np
import pandas as pd

from pandaspro.core.stringfunc import parse_wild
from pandaspro.core.tools.consecgrouper import ConsecGrouper
from pandaspro.core.tools.csort import csort
from pandaspro.core.tools.corder import corder
from pandaspro.core.tools.dfilter import dfilter
from pandaspro.core.tools.inrange import inrange
from pandaspro.core.tools.lowervarlist import lowervarlist
from pandaspro.core.tools.search2df import search2df
from pandaspro.core.tools.strpos import strpos
from pandaspro.core.tools.tab import tab
from pandaspro.core.tools.varnames import varnames
from pandaspro.core.tools.inlist import inlist
from pandaspro.core.tools.indate import indate
from pandaspro.io.excel.wbexportsimple import WorkbookExportSimplifier


class cpdBaseFrameMapper:
    def __init__(self, d):
        self.dict = d


class cpdBaseFrameList:
    def __init__(self, l):
        self.list = l


class FramePro(pd.DataFrame):
    def __init__(
            self,
            *args,
            uid: str = None,
            exr: str = None,
            rename_status: str = 'Process',
            **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.uid = uid
        self.export_mapper = cpdBaseFrameMapper(exr)
        self.rename_status = rename_status

    # noinspection PyFinal
    def __getattr__(self, item):
        def _parse_and_match(columns_list, attribute_name):
            """
            解析属性名并匹配列名
            对于 cpdtab2 系列，支持用 ___ 分隔 index 和 columns 字段
            例如：cpdtab2_Region__Product___Quarter__Category
            """
            # 对于包含 ___ 的 cpdtab2 系列，直接返回所有匹配的字段
            # 具体的分隔逻辑由调用者处理
            if attribute_name.startswith('cpdtab2') and '___' in attribute_name:
                # 提取所有字段名（去掉前缀）
                # 需要去掉 cpdtab2[s][aggfunc]_ 前缀
                temp = attribute_name
                if temp.startswith('cpdtab2s_'):
                    fields_part = temp[9:]
                elif temp.startswith('cpdtab2_'):
                    fields_part = temp[8:]
                else:
                    # 带聚合函数的版本，需要找到 _ 后的部分
                    # 例如：cpdtab2sum_... 或 cpdtab2ssum_...
                    for func_name in ['min', 'max', 'mean', 'median', 'sum', 'std', 'var', 'first', 'last']:
                        if temp.startswith('cpdtab2s' + func_name + '_'):
                            fields_part = temp[len('cpdtab2s' + func_name + '_'):]
                            break
                        elif temp.startswith('cpdtab2' + func_name + '_'):
                            fields_part = temp[len('cpdtab2' + func_name + '_'):]
                            break
                    else:
                        fields_part = None
                
                if fields_part:
                    # 提取所有字段名（用 __ 分隔，忽略 ___）
                    all_fields_str = fields_part.replace('___', '__')
                    all_fields = all_fields_str.split('__')
                    
                    # 匹配字段
                    matched_columns = [col for col in columns_list if col in all_fields]
                    
                    # 验证所有字段都能找到
                    if len(matched_columns) != len(all_fields):
                        missing = set(all_fields) - set(matched_columns)
                        raise ValueError(f"Some fields not found in dataframe columns. Missing: {missing}, Expected: {all_fields}, Found: {matched_columns}")
                    
                    # 按原始顺序排序
                    matched_columns.sort(key=lambda col: all_fields.index(col))
                    return matched_columns
            
            # 原有的解析逻辑（非 cpdtab2 系列或不含 ___）
            if attribute_name.startswith('cpdmap_'):
                key_part = attribute_name[7:].split('__')
            elif attribute_name.startswith('cpdlist_'):
                key_part = attribute_name[8:].split('__')
            elif attribute_name.startswith('cpdf_'):
                key_part = [attribute_name[5:].split('__')[0]]
            elif attribute_name.startswith('cpdfnot_'):
                key_part = [attribute_name[8:].split('__')[0]]
            elif attribute_name.startswith('cpdisna_'):
                key_part = attribute_name[8:].split('__')
            elif attribute_name.startswith('cpdnotna_'):
                key_part = attribute_name[9:].split('__')
            elif attribute_name.startswith('cpdtab_'):
                key_part = attribute_name[7:].split('__')
            elif attribute_name.startswith('cpdtabt_'):
                key_part = attribute_name[8:].split('__')
            elif attribute_name.startswith('cpdtabd_'):
                key_part = attribute_name[8:].split('__')
            elif attribute_name.startswith('cpdtab2s_'):
                key_part = attribute_name[9:].split('__')
            elif attribute_name.startswith('cpdtab2_'):
                key_part = attribute_name[8:].split('__')
            elif aggfunc and attribute_name.startswith('cpdtab2s' + aggfunc + '_'):
                prefix_length = len('cpdtab2s' + aggfunc + '_')
                key_part = attribute_name[prefix_length:].split('__')
            elif aggfunc and attribute_name.startswith('cpdtab2' + aggfunc + '_'):
                prefix_length = len('cpdtab2' + aggfunc + '_')
                key_part = attribute_name[prefix_length:].split('__')
            else:
                raise ValueError('prefix not added in [_parse_and_match] method')

            matched_columns = [col for col in columns_list if col in list(key_part)]

            if attribute_name.startswith('cpdmap_') and len(matched_columns) != 2:
                raise ValueError("Attribute var name parsing results does not match exactly two columns in the frame columns")
            if attribute_name.startswith('cpdlist_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            if attribute_name.startswith('cpdf_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            if attribute_name.startswith('cpdfnot_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            if attribute_name.startswith('cpdisna_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            if attribute_name.startswith('cpdnotna_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            if attribute_name.startswith('cpdtab_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            if attribute_name.startswith('cpdtabt_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            if attribute_name.startswith('cpdtabd_') and len(matched_columns) != 1:
                raise ValueError("Attribute var name parsing results does not match exactly 1 columns in the frame columns")
            # 检查 cpdtab2s_ (count with subtotals) - 不含 ___
            if attribute_name.startswith('cpdtab2s_') and '___' not in attribute_name and len(matched_columns) < 2:
                raise ValueError(f"Attribute var name parsing results needs at least 2 columns for pivot, matched columns are {matched_columns}")
            
            # 检查 cpdtab2s + aggfunc (e.g., cpdtab2ssum_) - 不含 ___
            if attribute_name.startswith('cpdtab2s') and '___' not in attribute_name and len(attribute_name) > 8 and attribute_name[8] != '_' and len(matched_columns) < 3:
                raise ValueError(f"Attribute var name parsing results needs at least 3 columns (index, columns, value), matched columns are {matched_columns}")
            
            # 检查 cpdtab2_ (count) - 不含 ___
            if attribute_name.startswith('cpdtab2_') and '___' not in attribute_name and not attribute_name.startswith('cpdtab2s') and len(matched_columns) < 2:
                raise ValueError(f"Attribute var name parsing results needs at least 2 columns for pivot, matched columns are {matched_columns}")
            
            # 检查 cpdtab2 + aggfunc (e.g., cpdtab2sum_) - 不含 ___
            if attribute_name.startswith('cpdtab2') and '___' not in attribute_name and not attribute_name.startswith('cpdtab2s') and len(attribute_name) > 7 and attribute_name[7] != '_' and len(matched_columns) < 3:
                raise ValueError(f"Attribute var name parsing results needs at least 3 columns (index, columns, value), matched columns are {matched_columns}")

            matched_columns.sort(key=lambda col: key_part.index(col))

            return matched_columns

        def _get_aggfunc(regex_item: str) -> str:
            pattern = r"^cpdtab2s?(min|max|mean|median|sum|std|var|first|last).*"
            match = re.search(pattern, regex_item)

            if match:
                return match.group(1)
            else:
                raise ValueError(f"Error: The input string '{regex_item}' is not in the correct format. "
                                 f"If you want to summarize by count, use only cpdtab2 followed by variable names. "
                                 f"If you want to use the aggregate shortcut of cpdtab2, "
                      f"it should start with 'cpdtab2' followed by a valid aggregation function (min, max, mean, median, sum, first, last, std, var).")

        def _add_subtotals(pivot_df):
            """
            为 pivot table 添加 subtotals
            - 对于 index 的第一级，为每个分组添加 subtotal 行（显示如 "华东 Subtotal"）
            - 对于 columns 的第一级（如果是 MultiIndex），为每个分组添加 subtotal 列（显示如 "Q1 Subtotal"）
            """
            result = pivot_df.copy()
            
            # 添加 index subtotals（如果 index 是 MultiIndex）
            if isinstance(result.index, pd.MultiIndex) and len(result.index.names) > 0:
                # 保存原始 index names
                index_names = result.index.names
                first_level_values = result.index.get_level_values(0).unique()
                
                subtotal_rows = []
                for value in first_level_values:
                    # 跳过 Total 行
                    if value in ['Total', 'All']:
                        continue
                        
                    # 选择该分组的所有行
                    mask = result.index.get_level_values(0) == value
                    group_data = result[mask]
                    
                    # 计算 subtotal
                    subtotal = group_data.sum(numeric_only=True)
                    
                    # 创建 subtotal 的 index - 带上分组名称
                    if len(index_names) == 2:
                        subtotal_index = (value, f'{value} Subtotal')
                    else:
                        # 对于更多层级，用 'Subtotal' 填充后续层级
                        subtotal_index = tuple([value] + [f'{value} Subtotal'] + [''] * (len(index_names) - 2))
                    
                    subtotal.name = subtotal_index
                    subtotal_rows.append(subtotal)
                
                # 将 subtotal 行添加到 DataFrame
                if subtotal_rows:
                    subtotal_df = pd.DataFrame(subtotal_rows)
                    subtotal_df.index.names = index_names  # 设置 index names
                    result = pd.concat([result, subtotal_df])
                    result = result.sort_index(level=0, sort_remaining=False)
                    result.index.names = index_names  # 确保 index names 保持不变
            
            # 添加 columns subtotals（如果 columns 是 MultiIndex）
            if isinstance(result.columns, pd.MultiIndex) and len(result.columns.levels) > 0:
                # 保存原始 column names
                column_names = result.columns.names
                first_level_values = [col[0] for col in result.columns]
                unique_first_levels = []
                seen = set()
                for val in first_level_values:
                    if val not in seen:
                        unique_first_levels.append(val)
                        seen.add(val)
                
                for value in unique_first_levels:
                    # 跳过 Total 列
                    if value in ['Total', 'All']:
                        continue
                        
                    # 选择该分组的所有列
                    cols_in_group = [col for col in result.columns if col[0] == value]
                    
                    # 计算 subtotal
                    subtotal_col = result[cols_in_group].sum(axis=1, numeric_only=True)
                    
                    # 创建 subtotal 的 column name - 带上分组名称
                    if len(result.columns.levels) == 2:
                        subtotal_col_name = (value, f'{value} Subtotal')
                    else:
                        # 对于更多层级，用 'Subtotal' 填充后续层级
                        subtotal_col_name = tuple([value] + [f'{value} Subtotal'] + [''] * (len(result.columns.levels) - 2))
                    
                    result[subtotal_col_name] = subtotal_col
                
                # 重新排序列，让 subtotal 列在每组的最后
                if isinstance(result.columns, pd.MultiIndex):
                    result = result.sort_index(axis=1, level=0, sort_remaining=False)
                    result.columns.names = column_names  # 确保 column names 保持不变
            
            return result

        if item in self.columns:
            return super().__getattr__(item)

        if item.startswith('cpdmap_'):
            dict_key_column, dict_value_column = _parse_and_match(self.columns, item)
            return self.set_index(dict_key_column)[dict_value_column].to_dict()

        elif item.startswith('cpdlist_'):
            list_column = _parse_and_match(self.columns, item)[0]
            return self[list_column].drop_duplicates().to_list()

        elif item.startswith('cpdf_'):
            list_column = _parse_and_match(self.columns, item)[0]
            value_filtered = item[10:].split('__')[1]
            return self.inlist(list_column, value_filtered)

        elif item.startswith('cpdfnot_'):
            list_column = _parse_and_match(self.columns, item)[0]
            value_filtered = item[10:].split('__')[1]
            return self.inlist(list_column, value_filtered, invert=True)

        elif item.startswith('cpdisna_'):
            notna_column = _parse_and_match(self.columns, item)[0]
            return self[self[notna_column].isna()]

        elif item.startswith('cpdnotna_'):
            notna_column = _parse_and_match(self.columns, item)[0]
            return self[self[notna_column].notna()]

        elif item.startswith('cpdtab_'):
            list_column = _parse_and_match(self.columns, item)[0]
            return self.tab(list_column)

        elif item.startswith('cpdtabt_'):
            list_column = _parse_and_match(self.columns, item)[0]
            return self.tab(list_column, 'detail')[[list_column, 'count']]

        elif item.startswith('cpdtabd_'):
            list_column = _parse_and_match(self.columns, item)[0]
            return self.tab(list_column, 'detail')

        elif item.startswith('cpdtab2s_'):
            # 检查是否包含 ___
            if '___' in item:
                # 使用 ___ 分隔 index 和 columns
                matched = _parse_and_match(self.columns, item)
                # 需要手动解析以确定分界
                fields_part = item[9:]
                parts = fields_part.split('___')
                index_fields = parts[0].split('__')
                columns_fields = parts[1].split('__')
                
                # 匹配字段到实际列名
                pivot_index = [col for col in self.columns if col in index_fields]
                pivot_columns = [col for col in self.columns if col in columns_fields]
                
                # 保持原始顺序
                pivot_index.sort(key=lambda x: index_fields.index(x))
                pivot_columns.sort(key=lambda x: columns_fields.index(x))
            else:
                # 原有逻辑：第一个是 index，第二个是 columns
                matched = _parse_and_match(self.columns, item)
                pivot_index = matched[0] if len(matched) > 0 else []
                pivot_columns = matched[1] if len(matched) > 1 else []

            if self.uid is None:
                idvar = self.columns[self.notnull().all()].tolist()[0]
            else:
                idvar = self.uid

            if self.export_mapper is not None and self.rename_status == 'Export':
                if isinstance(pivot_index, list):
                    pivot_index = [self.export_mapper.dict.get(x, x) for x in pivot_index]
                else:
                    pivot_index = self.export_mapper.dict.get(pivot_index, pivot_index)
                if isinstance(pivot_columns, list):
                    pivot_columns = [self.export_mapper.dict.get(x, x) for x in pivot_columns]
                else:
                    pivot_columns = self.export_mapper.dict.get(pivot_columns, pivot_columns)
                idvar = self.export_mapper.dict.get(idvar, idvar)

            pivot_result = self.pivot_table(
                index=pivot_index,
                columns=pivot_columns,
                values=idvar,
                aggfunc='count',
                margins=True,
                margins_name='Total'
            )
            
            return FramePro(_add_subtotals(pivot_result))

        elif item.startswith('cpdtab2_'):
            # 检查是否包含 ___
            if '___' in item:
                # 使用 ___ 分隔 index 和 columns
                matched = _parse_and_match(self.columns, item)
                # 需要手动解析以确定分界
                fields_part = item[8:]
                parts = fields_part.split('___')
                index_fields = parts[0].split('__')
                columns_fields = parts[1].split('__')
                
                # 匹配字段到实际列名
                pivot_index = [col for col in self.columns if col in index_fields]
                pivot_columns = [col for col in self.columns if col in columns_fields]
                
                # 保持原始顺序
                pivot_index.sort(key=lambda x: index_fields.index(x))
                pivot_columns.sort(key=lambda x: columns_fields.index(x))
            else:
                # 原有逻辑：第一个是 index，第二个是 columns
                matched = _parse_and_match(self.columns, item)
                pivot_index = matched[0] if len(matched) > 0 else []
                pivot_columns = matched[1] if len(matched) > 1 else []

            if self.uid is None:
                idvar = self.columns[self.notnull().all()].tolist()[0]
            else:
                idvar = self.uid

            if self.export_mapper is not None and self.rename_status == 'Export':
                if isinstance(pivot_index, list):
                    pivot_index = [self.export_mapper.dict.get(x, x) for x in pivot_index]
                else:
                    pivot_index = self.export_mapper.dict.get(pivot_index, pivot_index)
                if isinstance(pivot_columns, list):
                    pivot_columns = [self.export_mapper.dict.get(x, x) for x in pivot_columns]
                else:
                    pivot_columns = self.export_mapper.dict.get(pivot_columns, pivot_columns)
                idvar = self.export_mapper.dict.get(idvar, idvar)

            return FramePro(
                self.pivot_table(
                    index=pivot_index,
                    columns=pivot_columns,
                    values=idvar,
                    aggfunc='count',
                    margins=True,
                    margins_name='Total'
                )
            )

        elif item.startswith('cpdtab2s'):
            aggfunc = _get_aggfunc(item)
            
            # 检查是否包含 ___
            if '___' in item:
                # 使用 ___ 分隔 index 和 columns
                matched = _parse_and_match(self.columns, item)
                # 需要手动解析以确定分界
                prefix_length = len('cpdtab2s' + aggfunc + '_')
                fields_part = item[prefix_length:]
                parts = fields_part.split('___')
                index_fields = parts[0].split('__')
                columns_and_value = parts[1].split('__')
                
                # 最后一个是 value 字段
                columns_fields = columns_and_value[:-1]
                value_field = columns_and_value[-1]
                
                # 匹配字段到实际列名
                pivot_index = [col for col in self.columns if col in index_fields]
                pivot_columns = [col for col in self.columns if col in columns_fields]
                func_var = value_field if value_field in self.columns else None
                
                # 保持原始顺序
                pivot_index.sort(key=lambda x: index_fields.index(x))
                pivot_columns.sort(key=lambda x: columns_fields.index(x))
            else:
                # 原有逻辑：前面的是 index, columns, 最后一个是 value
                matched = _parse_and_match(self.columns, item)
                pivot_index = matched[0] if len(matched) > 0 else []
                pivot_columns = matched[1] if len(matched) > 1 else []
                func_var = matched[2] if len(matched) > 2 else None

            if self.export_mapper is not None and self.rename_status == 'Export':
                if isinstance(pivot_index, list):
                    pivot_index = [self.export_mapper.dict.get(x, x) for x in pivot_index]
                else:
                    pivot_index = self.export_mapper.dict.get(pivot_index, pivot_index)
                if isinstance(pivot_columns, list):
                    pivot_columns = [self.export_mapper.dict.get(x, x) for x in pivot_columns]
                else:
                    pivot_columns = self.export_mapper.dict.get(pivot_columns, pivot_columns)
                func_var = self.export_mapper.dict.get(func_var, func_var)

            pivot_result = self.pivot_table(
                index=pivot_index,
                columns=pivot_columns,
                values=func_var,
                aggfunc=aggfunc,
                margins=True,
                margins_name='Total'
            )
            
            return FramePro(_add_subtotals(pivot_result))

        elif item.startswith('cpdtab2'):
            aggfunc = _get_aggfunc(item)
            
            # 检查是否包含 ___
            if '___' in item:
                # 使用 ___ 分隔 index 和 columns
                matched = _parse_and_match(self.columns, item)
                # 需要手动解析以确定分界
                prefix_length = len('cpdtab2' + aggfunc + '_')
                fields_part = item[prefix_length:]
                parts = fields_part.split('___')
                index_fields = parts[0].split('__')
                columns_and_value = parts[1].split('__')
                
                # 最后一个是 value 字段
                columns_fields = columns_and_value[:-1]
                value_field = columns_and_value[-1]
                
                # 匹配字段到实际列名
                pivot_index = [col for col in self.columns if col in index_fields]
                pivot_columns = [col for col in self.columns if col in columns_fields]
                func_var = value_field if value_field in self.columns else None
                
                # 保持原始顺序
                pivot_index.sort(key=lambda x: index_fields.index(x))
                pivot_columns.sort(key=lambda x: columns_fields.index(x))
            else:
                # 原有逻辑：前面的是 index, columns, 最后一个是 value
                matched = _parse_and_match(self.columns, item)
                pivot_index = matched[0] if len(matched) > 0 else []
                pivot_columns = matched[1] if len(matched) > 1 else []
                func_var = matched[2] if len(matched) > 2 else None

            if self.export_mapper is not None and self.rename_status == 'Export':
                if isinstance(pivot_index, list):
                    pivot_index = [self.export_mapper.dict.get(x, x) for x in pivot_index]
                else:
                    pivot_index = self.export_mapper.dict.get(pivot_index, pivot_index)
                if isinstance(pivot_columns, list):
                    pivot_columns = [self.export_mapper.dict.get(x, x) for x in pivot_columns]
                else:
                    pivot_columns = self.export_mapper.dict.get(pivot_columns, pivot_columns)
                func_var = self.export_mapper.dict.get(func_var, func_var)

            return FramePro(
                self.pivot_table(
                    index=pivot_index,
                    columns=pivot_columns,
                    values=func_var,
                    aggfunc=aggfunc,
                    margins=True,
                    margins_name='All'
                )
            )

        else:
            return super().__getattr__(item)

    @property
    def _constructor(self):
        def _c(*args, **kwargs):
            return FramePro(*args, uid=self.uid, exr=self.export_mapper.dict, rename_status=self.rename_status, **kwargs)

        return _c

    @property
    def DF(self):
        return pd.DataFrame(self)

    @property
    def varnames(self):
        return varnames(self)

    def set_uid(self, varname):
        self.uid = varname
        return self._constructor()

    def set_exr(self, exr):
        self.export_mapper = cpdBaseFrameMapper(exr)
        return self._constructor()

    def set_rename_status(self, rename_status):
        self.rename_status = rename_status
        return self._constructor()

    def tab(self, name: str, d: str = 'brief', m: bool = False, sort: str = 'index', ascending: bool = True, label: str = None):
        return self._constructor(tab(self, name, d, m, sort, ascending, label))

    def dfilter(self, inputdict: dict = None, debug: bool = False):
        return self._constructor(dfilter(self, inputdict, debug))

    def csort(
            self,
            column,
            order=None,
            where=None,
            before=None,
            after=None,
            inplace=False
    ):
        return csort(
            self,
            column,
            order=order,
            where=where,
            before=before,
            after=after,
            inplace=inplace
        )

    def corder(
            self,
            column,
            before=None,
            after=None,
            pos='start'
    ):
        return corder(
            self,
            column,
            before=before,
            after=after,
            pos=pos
        )

    def inlist(
            self,
            colname: str,
            *args,
            engine: str = 'b',
            inplace: bool = False,
            invert: bool = False,
            rename: str = None,
            relabel_dict: dict = None,
            debug: bool = False
    ):
        result = inlist(
            self,
            colname,
            *args,
            engine=engine,
            inplace=inplace,
            invert=invert,
            rename=rename,
            relabel_dict=relabel_dict,
            debug=debug,
        )
        if debug:
            print("This is debugger for inlist method: ", result)
            print(type(result))
        if engine == 'm':
            return result
        else:
            return self._constructor(result)

    def inrange(
            self,
            colname: str,
            start,
            stop,
            inclusive: str = 'left',
            engine: str = 'b',
            inplace: bool = False,
            invert: bool = False,
            debug: bool = False
    ):
        result = inrange(
            self,
            colname,
            start,
            stop,
            inclusive=inclusive,
            engine=engine,
            inplace=inplace,
            invert=invert,
            debug=debug,
        )
        if debug:
            print(type(result))
        if engine == 'm':
            return result
        else:
            return self._constructor(result)

    def indate(
            self,
            colname,
            compare,
            date,
            end_date: str = None,
            inclusive: str = 'both',
            engine: str = 'b',
            inplace: bool = False,
            invert: bool = False,
    ):
        result = indate(
            self,
            colname,
            compare,
            date,
            end_date=end_date,
            inclusive=inclusive,
            engine=engine,
            inplace=inplace,
            invert=invert,
        )
        if engine == 'm':
            return result
        else:
            return self._constructor(result)

    def strpos(
            self,
            colname: str,
            *args,
            engine: str = 'b',
            inplace: bool = False,
            invert: bool = False,
            rename: str = None,
            debug: bool = False
    ):
        result = strpos(
            self,
            colname,
            *args,
            engine=engine,
            inplace=inplace,
            invert=invert,
            rename=rename,
            debug=debug,
        )
        if debug:
            print("This is debugger for strpos method: ", result)
            print(type(result))
        if engine == 'm':
            return result
        else:
            return self._constructor(result)

    def create_id(self):
        data = self.copy()
        if 'id' not in data.columns:
            data['id'] = range(1, len(data) + 1)
            data = data.corder('id')
            return data
        else:
            print('id column creation failure: already 1 column with the same name existed')
            return

    def create_ids(self):
        data = self.copy()
        if 'id' not in data.columns:
            data['id'] = range(1, len(data) + 1)
            data = data.corder('id')
            return data
        else:
            print('id column creation failure: already 1 column with the same name existed')
            return

    def lowervarlist(self, engine='columns', inplace=False):
        if engine == 'data':
            return self._constructor(lowervarlist(self, engine, inplace=inplace))
        return lowervarlist(self, engine, inplace=inplace)

    def excel_e(
            self,
            sheet_name: str = 'Sheet1',
            cell: str = 'A1',
            index: bool = False,
            header: bool = True,
            replace: str = None,
            sheetreplace: bool = False,
            design: str = None,
            style: str | list = None,
            cd: str | list = None,
            df_format: dict = None,
            cd_format: list | dict = None,
            config: dict = None,
            override: bool = None,
    ):
        declaredwb = WorkbookExportSimplifier.get_last_declared_workbook()
        if hasattr(self, 'df'):
            data = self.df
        else:
            data = self
        declaredwb.putxl(
            content=data,
            sheet_name=sheet_name,
            cell=cell,
            index=index,
            header=header,
            replace=replace,
            sheetreplace=sheetreplace,
            design=design,
            style=style,
            df_format=df_format,
            cd_format=cd_format,
            config=config,
            cd_style=cd
        )

        # ? Seems to return the declaredwb object to change
        if override:
            return declaredwb

    def expand_column(self, column_list):
        data = self.copy()
        data['expand_key'] = column_list[0]
        data['expand_value'] = data[column_list[0]]

        for i in range(1, len(column_list)):
            append = self.copy()
            append['expand_key'] = column_list[i]
            append['expand_value'] = append[column_list[i]]

            data = pd.concat([data, append], ignore_index=True)
        return data

    def cvar(self, promptstring):
        if self.empty:
            return []
        else:
            return parse_wild(promptstring, self.columns)

    def br(self, prompt):
        if isinstance(prompt, list):
            final_selection = []
            for item in prompt:
                if not self.cvar(item):
                    print('Nothing to check/browse in an empty dataframe')
                    return self
                else:
                    final_selection.extend(self.cvar(item))
            return self[final_selection]

        elif isinstance(prompt, str):
            if not self.cvar(prompt):
                print('Nothing to check/browse in an empty dataframe')
                return self
            else:
                return self[self.cvar(prompt)]
        else:
            raise TypeError('Invalid input type for prompt')

    def insert_blank(self, locator_dict: dict = None, how: str = 'after', nrows: int = 1):
        # Reset Index to Proceed
        org_cols = self.columns.to_list()
        new_cols = self.reset_index().columns.to_list()
        data_op = self.reset_index().copy()
        toResetIndex = [item for item in new_cols if item not in org_cols]
        if len(toResetIndex) != len(self.index.names):
            raise ValueError(
                "The insert_blank method only supports DataFrames where index labels and column names are unique and do not overlap.")

        # Location Dictionary Decipher into Slicing Points
        ##############################
        condition = pd.Series([True] * len(self), index=self.index)
        slice_indices = []

        if locator_dict is not None:
            for col, value in locator_dict.items():
                if not isinstance(value, list):
                    value = [value]
                else:
                    pass

                for v in value:
                    if col in self.columns:
                        locator = condition & (self[col] == v)
                        slice_indices.append(data_op.index[locator][0])
                    else:
                        print(f"Column '{col}' does not exist in the Frame.")
        #             return self
        else:
            pass

        # Define Cutting Machine
        ##############################
        def split_dataframe(df, indices, mode='before'):
            """
            Splits a DataFrame into segments based on a list of indices and a specified mode.

            Parameters:
            df (pd.DataFrame): The DataFrame to be split.
            indices (list): A list of indices where the splits should occur.
            mode (str): 'before' or 'after', indicating the split mode.

            Returns:
            list: A list of DataFrames resulting from the split.

            Example:
            --------
            Suppose you have a DataFrame `df`:

                A  B
            0   0 21
            1   1 22
            2   2 23
            3   3 24
            4   4 25
            5   5 26
            6   6 27
            7   7 28
            8   8 29
            9   9 30
            10 10 31

            And you want to split it using indices [2, 6] and mode 'before'.
            The function call would be: split_dataframe(df, [2, 6], 'before')

            This would produce three segments:
            Segment 1 (0 to 1):
                A  B
            0  0 21
            1  1 22

            Segment 2 (2 to 5):
                A  B
            2  2 23
            3  3 24
            4  4 25
            5  5 26

            Segment 3 (6 to end):
                A  B
            6  6 27
            7  7 28
            8  8 29
            9  9 30
            10 10 31
            """
            split_dfs = []

            if len(indices) == 0:
                split_dfs.append(df)

            else:
                indices = sorted(set(indices))
                if mode == 'before':
                    split_points = [0] + indices + [len(df)]
                elif mode == 'after':
                    split_points = [0] + [i + 1 for i in indices] + [len(df)]
                else:
                    raise ValueError("The mode parameter must be 'before' or 'after'")

                for i in range(len(split_points) - 1):
                    start, end = split_points[i], split_points[i + 1]
                    split_dfs.append(df.iloc[start:end])

            return split_dfs

        # Cut the DataFrames
        ##############################

        blank_fill = np.full((nrows, len(data_op.columns)), np.nan)
        blank_rows = pd.DataFrame(blank_fill, columns=data_op.columns)
        df_packages = split_dataframe(data_op, slice_indices, mode=how)

        output = pd.DataFrame()
        for index, dfl in enumerate(df_packages):
            output = pd.concat([output, dfl])
            if index + 1 != len(df_packages) or (len(df_packages) == 1 and how == 'after'):
                output = pd.concat([output, blank_rows])

        if len(df_packages) == 1 and how == 'before':
            output = pd.concat([blank_rows, output])

        output = self._constructor(output.set_index(toResetIndex))

        return output

    def search2df(
            self,
            data_large=None,
            dictionary=None,
            key=None,
            threshold=0.9,
            show=True,
            debug=False
    ):
        return search2df(
            data_small=self,
            data_large=data_large,
            dictionary=dictionary,
            key=key,
            threshold=threshold,
            show=show,
            debug=debug
        )

    @property
    def search2df_map(
            self,
    ):
        return search2df(
            data_small=self,
            mapsample=True
        )

    def consecgroup(self, groupby: str | list = None):
        return self._constructor(ConsecGrouper(self, groupby=groupby).group())

    def consecgroup_extract(
            self,
            groupby: str | list = None,
            value_at_top: str | list = None,
            value_at_bottom: str | list = None,
    ):
        return self._constructor(ConsecGrouper(self, groupby=groupby).extract(value_at_top, value_at_bottom))

    # __pandaspro_wangshiyao
    # add instruction and example of use
    def add_total(
            self,
            total_label_column,
            label: str = 'Total',
            sum_columns: str = '_all'
    ):
        total_row = {col: np.nan for col in self.columns}
        # noinspection PyTypeChecker
        total_row[total_label_column] = label

        if sum_columns == '_all':
            sum_columns = self.select_dtypes(include=[np.number]).columns.tolist()
        elif isinstance(sum_columns, (str, int)):
            sum_columns = [sum_columns]

        for col in sum_columns:
            if col in self.columns:
                total_sum = self[col].sum(min_count=1)  # 使用min_count=1确保全为np.nan时结果为0
                total_row[col] = total_sum if not pd.isna(total_sum) else 0

        total_df = pd.DataFrame([total_row], columns=self.columns)
        result = self._constructor(pd.concat([self, total_df], ignore_index=True))

        return result

    def show_duplicates(self, column_list):
        data = self.copy()
        result = self._constructor(data[data.duplicated(subset=column_list, keep='first')])
        return result

    # tab.__doc__ = pandaspro.core.tools.tab.tab.__doc__
    # dfilter.__doc__ = pandaspro.core.tools.dfilter.dfilter.__doc__
    # inlist.__doc__ = pandaspro.core.tools.inlist.__doc__
    # varnames.__doc__ = pandaspro.core.tools.varnames.varnames.__doc__
    # lowervarlist.__doc__ = lowervarlist.__doc__

    # Overwriting original methods
    def merge(self, *args, display=None, **kwargs):
        update = kwargs.pop('update', None)  # Extract the 'update' parameter and remove it from kwargs
        '''
        Think about updating this design in the future
        
        # Example usage
        left = CustomDataFrame({
            'key': ['K0', 'K1', 'K2', 'K3'],
            'A': ['A0', None, 'A2', 'A3'],
            'B': ['B0', 'B1', 'B2', None]
        })
        
        right = CustomDataFrame({
            'key': ['K0', 'K1', 'K2', 'K3'],
            'A': ['C0', 'C1', 'C2', 'C3'],
            'C': ['D0', 'D1', 'D2', 'D3']
        })
        
        # Use the new merge method with 'update' parameter
        result_missing = left.merge(right, on='key', update='missing')
        result_all = left.merge(right, on='key', update='all')
        
        print("Result with update='missing':\n", result_missing, "\n")
        print("Result with update='missing':\n", result_missing, "\n")
        print("Result with update='all':\n", result_all)
        '''

        result = super().merge(*args, **kwargs)

        if update == 'missing':
            for col in result.columns:
                if '_x' in col and col.replace('_x', '_y') in result.columns:
                    # Update only if the left column has missing values
                    result[col] = result[col].fillna(result[col.replace('_x', '_y')])
            # Drop the columns from the right DataFrame
            result = result.drop(columns=[col for col in result.columns if '_y' in col])

        elif update == 'all':
            for col in result.columns:
                if '_x' in col and col.replace('_x', '_y') in result.columns:
                    # Update the left column with values from the right column
                    result[col] = result[col.replace('_x', '_y')]
            # Drop the columns from the right DataFrame
            result = result.drop(columns=[col for col in result.columns if '_y' in col])

        result.columns = [col.replace('_x', '') for col in result.columns]
        if '_merge' in result.columns and display is not None:
            # noinspection PyTestUnpassedFixture
            print(result.tab('_merge'))
        return self._constructor(result)

    def rename(self, columns=None, *args, **kwargs):
        return self._constructor(super().rename(columns=columns, *args, **kwargs))


pd.DataFrame.excel_e = FramePro.excel_e
