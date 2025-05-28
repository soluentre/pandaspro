import pandas as pd
import numpy as np


def replace_left_with_right(key, left, right, override_column='override_status'):
    if not set(right.columns).issubset(left.columns):
        extra_cols = set(right.columns) - set(left.columns)
        raise ValueError(f"Right contains columns not present in left: {extra_cols}")

    merged = left.merge(right, on=key, how='left', suffixes=('_left', '_right'))

    right_suffix_cols = [f"{col}_right" for col in right.columns if col != key]
    if merged[right_suffix_cols].isna().all().all():
        print(
            f"Warning: No matching rows found between left and right on key '{key}'. Returning original left DataFrame.")
        left[override_column] = ''
        return left

    right_cols_to_process = [col for col in right.columns if col != key]
    override = pd.Series(False, index=merged.index)

    for col in right_cols_to_process:
        left_col = f"{col}_left"
        right_col = f"{col}_right"

        merged[col] = merged[left_col].mask(merged[right_col].notna(), merged[right_col])
        override |= merged[right_col].notna()

    merged[override_column] = np.where(override, 'Y', '')

    result_df = merged.copy()
    for col in right.columns.difference([key]):
        result_df[col] = merged[col]
    result_df = result_df[left.columns.tolist() + [override_column]]

    return result_df


def replace_left_with_target(
        old_df,
        new_df,
        changed_col: str,
        old_id_col: str,
        new_id_col: str,
        check_na: list = None
):
    """
    在 old_df 中根据 new_df 的 changed_col 列记录，对指定列进行定向更新。

    参数：
    - old_df: 原始 DataFrame，需要被更新的表
    - new_df: 更新用的 DataFrame，包含 changed_col 和 new_id_col
    - changed_col: new_df 中记录要更新列名的列名（逗号分隔）
    - old_id_col: old_df 中用于匹配新旧记录的 ID 列名
    - new_id_col: new_df 中用于匹配新旧记录的 ID 列名
    - check_na: （可选）要检查缺失值的列名列表；若不为 None，函数会打印出 old_df

    返回：
    - 更新后的 old_df（在原 DataFrame 基础上修改并返回同一个对象）
    """
    if check_na:
        narows = old_df[old_df[old_id_col].isna()][[old_id_col] + check_na]
        print(f"============ Below rows don't have {old_id_col} ============")
        print(narows)

    mask = new_df[changed_col].notna()
    for _, row in new_df[mask].iterrows():
        new_id = row[new_id_col]
        cols_to_update = [c.strip() for c in row[changed_col].split(',')]
        values = row[cols_to_update].values
        old_df.loc[old_df[old_id_col] == new_id, cols_to_update] = values

    return old_df
