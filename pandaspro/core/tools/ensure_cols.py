import pandas as pd
import numpy as np


def ensure_columns(
        data,
        columns: list,
        inplace: bool = False
):
    """
    确保 DataFrame 有且仅有指定列表中的所有列，并按照列表顺序排列。
    如果某列不存在则用 np.nan 填充创建，如果某列不在列表中则删除。
    
    :param data: 要处理的 DataFrame
    :param columns: 指定的列名列表
    :param inplace: 是否在原地修改 DataFrame（默认: False）
    
    :return: 处理后的 DataFrame，如果 inplace=True 则返回 None
    
    Example:
        >>> df = pd.DataFrame({'A': [1, 2], 'B': [3, 4], 'D': [5, 6]})
        >>> result = ensure_columns(df, columns=['A', 'B', 'C'])
        >>> # 结果: DataFrame 只有 A, B, C 三列，C 列用 NaN 填充，按 A, B, C 顺序排列
    """
    if not isinstance(columns, list):
        raise TypeError('columns 参数必须是列表类型')
    
    # 如果传入空列表，直接返回原样
    if len(columns) == 0:
        if inplace:
            return None
        else:
            return data.copy()
    
    if inplace:
        # 删除不在列表中的列
        cols_to_drop = [col for col in data.columns if col not in columns]
        if cols_to_drop:
            data.drop(columns=cols_to_drop, inplace=True)
        
        # 添加缺失的列（用 NaN 填充）
        for col in columns:
            if col not in data.columns:
                data[col] = np.nan
        
        # 按照指定顺序重排列
        data_reordered = data[columns]
        data.drop(columns=data.columns, inplace=True)
        for col in columns:
            data[col] = data_reordered[col]
    else:
        # 创建新的 DataFrame
        result = data.copy()
        
        # 删除不在列表中的列
        cols_to_drop = [col for col in result.columns if col not in columns]
        if cols_to_drop:
            result.drop(columns=cols_to_drop, inplace=True)
        
        # 添加缺失的列（用 NaN 填充）
        for col in columns:
            if col not in result.columns:
                result[col] = np.nan
        
        # 按照指定顺序重排列
        result = result[columns]
        
        return result
