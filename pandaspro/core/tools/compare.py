def compare(df_output, df_used_for_check, index_col, keep_cols):
    # 1. 以 index_col 做索引对齐
    df1_idx = df_output.set_index(index_col)
    df2_idx = df_used_for_check.set_index(index_col)
    common_idx = df1_idx.index.intersection(df2_idx.index)
    df1c = df1_idx.loc[common_idx]
    df2c = df2_idx.loc[common_idx]

    # 2. 计算“至少一方非空”的掩码
    not_both_na = df1c.notna() | df2c.notna()

    # 3. 找到真正不同的单元格
    diff_mask = df1c.ne(df2c) & not_both_na

    # 4. 筛出有差异的行和列
    rows_diff = diff_mask.any(axis=1)
    diff_cols = diff_mask.any(axis=0)
    diff_col_list = diff_cols[diff_cols].index.tolist()

    # 5. 取出有差异的 df1 片段，并 reset_index 恢复 index_col 为列
    df1_diff = df1c.loc[rows_diff]
    df2_diff = df2c.loc[rows_diff]
    base = df1_diff.reset_index()

    # 6. 从 base 中选出 keep_cols（此时 index_col 已经是列了）
    result = base[keep_cols].copy()

    # 7. 把所有差异列在 df1/df2 中的值拼接进来
    for col in diff_col_list:
        # 这里直接按 index 顺序取值
        result[f'{col}_df1'] = df1_diff[col].values
        result[f'{col}_df2'] = df2_diff[col].values

    # 8. 计算 diff_fields 列，并映射到 result 中
    diff_fields = diff_mask.loc[rows_diff, diff_col_list] \
                         .apply(lambda row: row[row].index.tolist(), axis=1) \
                         .to_dict()
    # 将 diff_fields 插入到最前面
    result.insert(0, 'diff_fields', result[index_col].map(diff_fields))

    return result