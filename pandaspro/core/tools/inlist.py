import pandas as pd
from pandaspro.core.tools.utils import df_with_index_for_mask


def inlist(
    data,
    colname: str,
    *args,
    engine: str = 'b',
    inplace: bool = False,
    invert: bool = False,
    rename: str = None,
    relabel_dict: dict = None,
    debug: bool = False,
):
    """
    Filters a DataFrame based on whether values in a specified column are in a given list. Supports various
    operation types including filtering, masking, and creating a new indicator column.

    Parameters
    ----------
    data : DataFrame
        The DataFrame to operate on.
    colname : str
        The name of the column to check values against the list.
    *args : list or elements
        The list of values to check against or multiple arguments forming the list.
    engine : str, optional
        The operation type:
        'b' for boolean indexing (default)
        'r' for row filtering
        'm' for mask
        'c' for adding a new column.
    inplace : bool, optional
        If True (for 'r', 'b', 'c'), operates on a copy to avoid SettingWithCopyWarning.
    invert : bool, optional
        If True, inverts the condition. Defaults to False.
    rename : str, optional
        If provided, the new column created under 'c' will use this name.
    relabel_dict : dict, optional
        If provided, must be {1: ..., 0: ...}, to relabel the indicator column.
    debug : bool, optional
        If True, prints debug information. Defaults to False.

    Returns
    -------
    DataFrame or Series or None
        Depends on engine:
        - 'r', 'b', 'c': returns a new DataFrame (even if inplace=True).
        - 'm': returns a boolean Series.
    """

    if data.empty:
        raise ValueError('Cannot use inlist on an empty dataframe')

    # handle index-based mask if column is in index
    if colname in data.columns:
        working = data
    elif colname in data.index.names:
        working = df_with_index_for_mask(data)
    else:
        raise ValueError(f'Column {colname} not found in DataFrame or index')

    # if we will mutate, ensure we're on a copy to avoid SettingWithCopyWarning
    if inplace or engine == 'r':
        working = working.copy()

    # flatten arguments into a list of values
    bool_list = []
    for arg in args:
        if isinstance(arg, (list, tuple, set)):
            bool_list.extend(arg)
        else:
            bool_list.append(arg)
    if debug:
        print('Values to match:', bool_list)

    # filtering logic
    if engine == 'r':
        if debug:
            print("[r] filtering rows")
        mask = working[colname].isin(bool_list)
        if invert:
            mask = ~mask
        result = working[mask]
        # drop index columns if they were made into columns
        if set(result.index.names) <= set(result.columns):
            result = result.drop(list(result.index.names), axis=1)
        return result

    elif engine == 'b':
        if debug:
            print("[b] boolean indexing")
        mask = working[colname].isin(bool_list)
        if invert:
            mask = ~mask
        result = working[mask]
        if set(result.index.names) <= set(result.columns):
            result = result.drop(list(result.index.names), axis=1)
        return result

    elif engine == 'm':
        if debug:
            print("[m] creating mask")
        mask = working[colname].isin(bool_list)
        return ~mask if invert else mask

    elif engine == 'c':
        if debug:
            print("[c] creating indicator column")
        new_name = rename if rename else '_inlist'
        yes_label, no_label = (relabel_dict.get(1), relabel_dict.get(0)) if relabel_dict else (1, 0)
        cond = working[colname].isin(bool_list)
        if invert:
            cond = ~cond
        # build new DataFrame
        df = working.copy()
        df[new_name] = no_label
        df.loc[cond, new_name] = yes_label
        if set(df.index.names) <= set(df.columns):
            df = df.drop(list(df.index.names), axis=1)
        return df

    else:
        raise ValueError(f"Unsupported engine type '{engine}'")
