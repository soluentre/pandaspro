import re
from typing import Union

import numpy as np
import pandas as pd


def lowervarlist(
        data,
        engine: str = 'data',
        inplace: bool = False
) -> Union[pd.DataFrame, list, dict]:
    """
    This function renames the columns of a DataFrame by formatting the original column names
    according to a specified pattern, primarily to ensure the column names are
    valid identifiers suitable for use in queries or further processing.
    The function provides options for the type of output through the `engine` parameter
    and can perform the operation in-place if desired.

    Parameters
    ----------
    data : DataFrame
        The DataFrame whose columns are to be renamed.
    engine : str, optional
        Determines the type of output returned by the function. Options include:
        - 'data': Returns a new DataFrame with renamed columns (default).
        - 'column': Returns a list of new column names.
        - 'update_map': Returns a dictionary mapping original column names to new column names.
        - 'revert_map': Returns a dictionary mapping new column names back to original column names.
    inplace : bool, optional
        If True, the column renaming is applied in-place, and the function returns None. Defaults to False.

    Returns
    -------
    DataFrame, list, dict, or None
        The return type depends on the `engine` parameter:
        - If `engine` is 'data', returns a new DataFrame with renamed columns.
        - If `engine` is 'column', returns a list of the new column names.
        - If `engine` is 'update_map', returns a dictionary mapping from original to new column names.
        - If `engine` is 'revert_map', returns a dictionary mapping from new to original column names.
        - If `inplace` is True, the function modifies the input DataFrame in-place and returns None.

    Notes
    -----
    The function formats column names by replacing non-alphanumeric characters with underscores, converting to lowercase, and appending a suffix to duplicate names to ensure uniqueness.

    Examples
    --------
    >>> mydf = pd.DataFrame(np.random.rand(3, 3), columns=['Column 1', 'Column-2', 'Column 3'])
    >>> lowervarlist(mydf, 'update_map')
    This will return a dictionary mapping the original column names to their new, formatted names, such as {'Column 1': 'column_1', 'Column-2': 'column_2', 'Column 3': 'column_3'}.
    """

    _engines = {
        'data': 0,
        'columns': 1,
        'update': 2,
        'revert': 3
    }

    # Get the original list of column names
    oldname = data.columns.to_list()
    pattern = re.compile(r'\W+')

    # Dictionary to track the occurrence of each formatted column name
    name_count = {}
    newname = []
    for name in oldname:
        # Format the column name
        formatted_name = re.sub(pattern, '_', str(name)).lower().strip("_")

        # Increment count and modify name if it's a duplicate
        if formatted_name in name_count:
            name_count[formatted_name] += 1
            formatted_name += f"_{name_count[formatted_name]}"
        else:
            name_count[formatted_name] = 0

        newname.append(formatted_name)

    # Create a mapping of old column names to new column names
    mapping_update = {old: new for old, new in zip(oldname, newname)}
    mapping_revert = {new: old for new, old in zip(newname, oldname)}

    # Rename the columns in the DataFrame
    df = data.rename(columns=mapping_update)
    cols = df.columns.to_list()

    if inplace:
        data.rename(columns=mapping_update, inplace=True)
    else:
        return [df, cols, mapping_update, mapping_revert][_engines[engine]]
