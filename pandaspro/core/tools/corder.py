import pandas as pd

def corder(
        data,
        column: str | list,
        before=None,
        after=None,
        pos: str = 'start',
):
    """
    Reorder the DataFrame columns by positioning specified columns before or after the corresponding columns.
    Supports MultiIndex columns using __ separator (e.g., "Level1__Level2").

    :param data: DataFrame to reorder
    :param column: Column name on which to change position, will be switched to the beginning of the DataFrame if param before and after are not specified
                   For MultiIndex columns, use __ to separate levels (e.g., "CMU__ACS_Staff")
    :param before: The column name before which the specified column should be placed (optional)
    :param after: The column name after which the specified column should be placed (optional)
    :param pos: Position 'start' or 'end' (default: 'start')

    :return: Reordered DataFrame or None if reordered in place
    """
    # Helper function to convert string to actual column (handles MultiIndex)
    def _resolve_column_name(col_str, df_columns):
        """Convert column string to actual column object (handles MultiIndex)"""
        if isinstance(df_columns, pd.MultiIndex):
            # Check if it's already a tuple (actual column)
            if isinstance(col_str, tuple):
                return col_str
            # Check if it contains __ separator
            if isinstance(col_str, str) and '__' in col_str:
                parts = col_str.split('__')
                # Find matching column in MultiIndex
                for col in df_columns:
                    if len(col) == len(parts) and all(str(col[i]) == parts[i] for i in range(len(parts))):
                        return col
                # If no match found, return as-is (will be caught in validation)
                return col_str
            # Try to find in columns as-is
            if col_str in df_columns:
                return col_str
            # Return as-is for error handling
            return col_str
        else:
            # For regular columns, return as-is
            return col_str
    
    # Parse column parameter
    if isinstance(column, str):
        cols_str = [i.strip() for i in column.split(';')]
    elif isinstance(column, list):
        cols_str = column
    else:
        raise TypeError("column parameter must be str or list")
    
    # Resolve column names (handle MultiIndex)
    cols = [_resolve_column_name(c, data.columns) for c in cols_str]
    
    # Also resolve before/after if provided
    if before:
        before = _resolve_column_name(before, data.columns)
    if after:
        after = _resolve_column_name(after, data.columns)

    # Validate columns and collect those not in dataframe
    remove_list = []
    for i, col in zip(cols_str, cols):
        if col in data.columns:
            pass
        else:
            remove_list.append(i)
    
    if len(remove_list) != 0:
        print(f'Columns {remove_list} not in the dataframe, removed and new lists updated')
    
    retain_list = [col for col in cols if col in data.columns]
    old_order = [i for i in list(data.columns) if i not in retain_list]

    # Build new column order
    if before:
        if before not in old_order:
            raise ValueError(f"Column {before} not found in dataframe (after excluding columns to move)")
        index = old_order.index(before)
        new_order = old_order[:index] + retain_list + old_order[index:]
    elif after:
        if after not in old_order:
            raise ValueError(f"Column {after} not found in dataframe (after excluding columns to move)")
        index = old_order.index(after)
        new_order = old_order[:index+1] + retain_list + old_order[index+1:]
    elif pos == 'end':
        new_order = old_order + retain_list
    else:
        new_order = retain_list + old_order

    # Reorder columns - works for both regular and MultiIndex columns
    return data[new_order]


if __name__ == '__main__':
    import sprnldata as spr
    a = spr.monades()
    b = a.corder('cancel; amount', after='country')