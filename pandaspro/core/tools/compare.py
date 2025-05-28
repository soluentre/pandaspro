def compare(df_output, df_used_for_check, index_col, keep_cols):
    """
    Compare two DataFrames and return rows with differences.

    Args:
        df_output: First DataFrame to compare
        df_used_for_check: Second DataFrame to compare
        index_col: Column name to use as index for alignment
        keep_cols: List of columns to keep in the result

    Returns:
        DataFrame showing rows with differences between the two inputs
    """

    # 1. Validate input DataFrames
    # Check if shapes match
    if df_output.shape != df_used_for_check.shape:
        raise ValueError("DataFrames have different shapes. "
                         f"df_output: {df_output.shape}, df_used_for_check: {df_used_for_check.shape}")

    # Check if column names match exactly
    if not df_output.columns.equals(df_used_for_check.columns):
        raise ValueError("Column names do not match exactly between DataFrames.")

    # 2. Align DataFrames using index_col
    df1_idx = df_output.set_index(index_col)
    df2_idx = df_used_for_check.set_index(index_col)
    common_idx = df1_idx.index.intersection(df2_idx.index)
    df1c = df1_idx.loc[common_idx]
    df2c = df2_idx.loc[common_idx]

    # 3. Create mask for "at least one non-NA" cells
    not_both_na = df1c.notna() | df2c.notna()

    # 4. Find truly different cells
    diff_mask = df1c.ne(df2c) & not_both_na

    # 5. Filter rows and columns with differences
    rows_diff = diff_mask.any(axis=1)
    diff_cols = diff_mask.any(axis=0)
    diff_col_list = diff_cols[diff_cols].index.tolist()

    # 6. Extract differing portions and reset index
    df1_diff = df1c.loc[rows_diff]
    df2_diff = df2c.loc[rows_diff]
    base = df1_diff.reset_index()

    # 7. Select columns to keep from base (index_col is now a column)
    result = base[keep_cols].copy()

    # 8. Append all differing values from both DataFrames
    for col in diff_col_list:
        # Values are taken directly in index order
        result[f'{col}_df1'] = df1_diff[col].values
        result[f'{col}_df2'] = df2_diff[col].values

    # 9. Calculate diff_fields column and map to result
    diff_fields = diff_mask.loc[rows_diff, diff_col_list] \
        .apply(lambda row: row[row].index.tolist(), axis=1) \
        .to_dict()
    # Insert diff_fields at the beginning
    result.insert(0, 'diff_fields', result[index_col].map(diff_fields))

    return result
