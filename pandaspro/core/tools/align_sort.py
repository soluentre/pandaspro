import pandas as pd
from pandaspro.core.tools.utils import df_with_index_for_mask


def align_and_sort_by_order(
        data,
        input_col: str,
        order: list,
        output_col: str = None,
        inplace: bool = False
):
    """
    按照指定的 order 对某列进行排序，并可选地重命名该列。
    
    :param data: 要排序的 DataFrame
    :param input_col: 要排序的列名
    :param order: 定义排序顺序的列表
    :param output_col: 输出列名（可选），如果提供则重命名列，否则保持原列名
    :param inplace: 是否在原地修改 DataFrame（默认: False）
    
    :return: 排序后的 DataFrame，如果 inplace=True 则返回 None
    
    Example:
        >>> df = pd.DataFrame({'grade_fo': ['B', 'A', 'C'], 'value': [1, 2, 3]})
        >>> order_staff = ['A', 'B', 'C']
        >>> result = align_and_sort_by_order(df, input_col='grade_fo', order=order_staff, output_col='grade')
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
    
    # 获取所有唯一值
    unique_values = list(data[input_col].dropna().unique())
    
    # 补全 order 列表（将 order 中没有但数据中存在的值追加到末尾）
    provided_order = [x for x in order if x in unique_values]
    missing_values = [x for x in unique_values if x not in order]
    full_orderlist = provided_order + missing_values
    
    # 创建分类类型
    cat_type = pd.CategoricalDtype(categories=full_orderlist, ordered=True)
    
    # 创建临时排序列
    data['__cpd_align_sort'] = data[input_col].astype(cat_type)
    
    if inplace:
        # 原地排序
        data.sort_values(by='__cpd_align_sort', inplace=True, kind='mergesort')
        
        # 如果需要重命名列
        if output_col != input_col:
            data.rename(columns={input_col: output_col}, inplace=True)
        
        # 删除临时排序列
        data.drop('__cpd_align_sort', axis=1, inplace=True)
        
        # 如果索引名在列中，删除它们
        if set(data.index.names) <= set(data.columns):
            data.drop([name for name in data.index.names if name in data.columns], axis=1, inplace=True)
    else:
        # 返回新的 DataFrame
        result = data.sort_values(by='__cpd_align_sort', kind='mergesort').copy()
        
        # 如果需要重命名列
        if output_col != input_col:
            result.rename(columns={input_col: output_col}, inplace=True)
        
        # 删除临时排序列
        result.drop('__cpd_align_sort', axis=1, inplace=True)
        
        # 如果索引名在列中，删除它们
        if result.index.names[0] is not None and set(result.index.names) <= set(result.columns):
            result.drop([name for name in result.index.names if name in result.columns], axis=1, inplace=True)
        
        return result
