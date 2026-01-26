import pandas as pd
from pandaspro.core.tools.utils import df_with_index_for_mask


def align_and_sort_by_order(
        data,
        input_col: str,
        order: list,
        output_col: str = None,
        inplace: bool = False,
        keep_structure: bool = True,
        fill_value=None
):
    """
    按照指定的 order 对某列进行排序，并可选地重命名该列。
    只保留 order 中存在的值，其他值会被过滤掉。
    
    :param data: 要排序的 DataFrame
    :param input_col: 要排序的列名
    :param order: 定义排序顺序的列表，只保留列表中的值
    :param output_col: 输出列名（可选），如果提供则重命名列，否则保持原列名
    :param inplace: 是否在原地修改 DataFrame（默认: False）
    :param keep_structure: 是否保持 order 的完整结构（默认: True）
                          如果为 True，order 中所有值都会在结果中，缺失的会填充空行
    :param fill_value: 当 keep_structure=True 且需要填充缺失行时，其他列的填充值（默认: None/NaN）
    
    :return: 排序后的 DataFrame，如果 inplace=True 则返回 None
    
    Example:
        >>> df = pd.DataFrame({'grade_fo': ['B', 'A', 'C', 'D'], 'value': [1, 2, 3, 4]})
        >>> order_staff = ['A', 'B', 'C']
        >>> result = align_and_sort_by_order(df, input_col='grade_fo', order=order_staff, output_col='grade')
        >>> # 结果只包含 A, B, C，D 被过滤掉
        
        >>> # 保持结构模式
        >>> df2 = pd.DataFrame({'grade': ['B', 'D'], 'value': [1, 2]})
        >>> order = ['A', 'B', 'C']
        >>> result = align_and_sort_by_order(df2, input_col='grade', order=order, keep_structure=True)
        >>> # 结果包含 A, B, C 三行，A 和 C 的 value 列为 NaN，D 被过滤掉
    """
    # 检查列是否存在
    if input_col in data.columns:
        pass
    elif input_col in data.index.names:
        data = df_with_index_for_mask(data)
    else:
        raise ValueError(f'列 {input_col} 既不在 DataFrame 中也不在索引中')
    
    # 如果没有指定 output_col，使用原列名
    if output_col is None:
        output_col = input_col
    
    # 根据 keep_structure 参数决定要使用的完整顺序列表
    if keep_structure:
        # 保持结构：使用完整的 order 列表
        full_orderlist = order
    else:
        # 只保留 order 中存在的值
        full_orderlist = [x for x in order if x in data[input_col].values]
    
    if inplace:
        # 先过滤只保留 order 中的值
        mask = data[input_col].isin(order)
        rows_to_drop = data.index[~mask]
        data.drop(rows_to_drop, inplace=True)
        
        # 如果 keep_structure=True，需要填充缺失的行
        if keep_structure:
            # 找出 order 中在 data 中不存在的值
            missing_values = [x for x in order if x not in data[input_col].values]
            
            if missing_values:
                # 创建缺失行的 DataFrame
                missing_rows = []
                for val in missing_values:
                    new_row = {col: fill_value for col in data.columns}
                    new_row[input_col] = val
                    missing_rows.append(new_row)
                
                missing_df = pd.DataFrame(missing_rows)
                
                # 合并到原 DataFrame
                # 注意：使用 pd.concat 而不是 append（已废弃）
                data_combined = pd.concat([data, missing_df], ignore_index=True)
                
                # 清空原 data 并用合并后的数据填充
                data.drop(data.index, inplace=True)
                for col in data_combined.columns:
                    data[col] = data_combined[col].values
        
        # 创建分类类型
        cat_type = pd.CategoricalDtype(categories=full_orderlist, ordered=True)
        data['__cpd_align_sort'] = data[input_col].astype(cat_type)
        
        # 原地排序
        data.sort_values(by='__cpd_align_sort', inplace=True, kind='mergesort')
        
        # 重置索引
        data.reset_index(drop=True, inplace=True)
        
        # 如果需要重命名列
        if output_col != input_col:
            data.rename(columns={input_col: output_col}, inplace=True)
        
        # 删除临时排序列
        data.drop('__cpd_align_sort', axis=1, inplace=True)
        
        # 如果索引名在列中，删除它们
        if set(data.index.names) <= set(data.columns):
            data.drop([name for name in data.index.names if name in data.columns], axis=1, inplace=True)
    else:
        # 先过滤只保留 order 中的值
        mask = data[input_col].isin(order)
        result = data[mask].copy()
        
        # 如果 keep_structure=True，需要填充缺失的行
        if keep_structure:
            # 找出 order 中在 result 中不存在的值
            missing_values = [x for x in order if x not in result[input_col].values]
            
            if missing_values:
                # 创建缺失行的 DataFrame
                missing_rows = []
                for val in missing_values:
                    new_row = {col: fill_value for col in result.columns}
                    new_row[input_col] = val
                    missing_rows.append(new_row)
                
                missing_df = pd.DataFrame(missing_rows)
                
                # 合并到结果 DataFrame
                result = pd.concat([result, missing_df], ignore_index=True)
        
        # 创建分类类型
        cat_type = pd.CategoricalDtype(categories=full_orderlist, ordered=True)
        result['__cpd_align_sort'] = result[input_col].astype(cat_type)
        
        # 排序
        result = result.sort_values(by='__cpd_align_sort', kind='mergesort')
        
        # 重置索引
        result.reset_index(drop=True, inplace=True)
        
        # 如果需要重命名列
        if output_col != input_col:
            result.rename(columns={input_col: output_col}, inplace=True)
        
        # 删除临时排序列
        result.drop('__cpd_align_sort', axis=1, inplace=True)
        
        # 如果索引名在列中，删除它们
        if result.index.names[0] is not None and set(result.index.names) <= set(result.columns):
            result.drop([name for name in result.index.names if name in result.columns], axis=1, inplace=True)
        
        return result
