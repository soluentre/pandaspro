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
        left[override_column] = 'N'
        return left

    right_cols_to_process = [col for col in right.columns if col != key]
    override = pd.Series(False, index=merged.index)

    for col in right_cols_to_process:
        left_col = f"{col}_left"
        right_col = f"{col}_right"

        merged[col] = merged[left_col].mask(merged[right_col].notna(), merged[right_col])
        override |= merged[right_col].notna()

    merged[override_column] = np.where(override, 'Y', 'N')

    result_df = merged.copy()
    for col in right.columns.difference([key]):
        result_df[col] = merged[col]
    result_df = result_df[left.columns.tolist() + [override_column]]

    return result_df
