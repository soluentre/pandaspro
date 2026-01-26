import pandas as pd


def varnames(self,
             rows: int = None,
             cols: int = None) -> pd.DataFrame:
    r"""
         _  _   _  ___ __   __ _____  ___  ___    ___   _  _  _  __   __
      _ | || | | || _ \\ \ / /|_   _|| __|| _ \  / _ \ | \| || | \ \ / /
     | || || |_| ||  _/ \ V /   | |  | _| |   / | (_) || .` || |__\ V /
      \__/  \___/ |_|    |_|    |_|  |___||_|_\  \___/ |_|\_||____||_|

    This function rearranges the column names of a DataFrame into a tabular format for
    easier visualization, with options to specify the number of rows or columns.

    The output is a styled DataFrame where each cell contains a column name.
    Users can customize the layout by specifying either the number of rows or columns,
    and the function will automatically adjust the other dimension based on the total number
    of columns in the DataFrame. Additionally, the function enhances readability by
    adding CSS styles to the output table, including cell padding, text alignment,
    font weight, background color, and border style.

    Parameters
    ----------
    self : DataFrame
        The DataFrame whose column names are to be rearranged and displayed.
    rows : int, optional
        The desired number of rows in the output table.
        If specified, the number of columns will be calculated accordingly.
        Defaults to None, in which case a default of 20 rows is used unless `cols` is specified.
    cols : int, optional
        The desired number of columns in the output table.
        If specified, the number of rows will be calculated accordingly.
        Defaults to None.

    Returns
    -------
    Styler
        A pandas.io.formats.style.Styler object representing the DataFrame of column names
        arranged in the specified tabular format and styled with CSS for improved readability.

    Notes
    -----
    - If neither `rows` nor `cols` is specified, the function defaults to 20 rows
      and calculates the necessary number of columns.
    - If both `rows` and `cols` are specified, the function prioritizes `rows`, ignoring `cols`.
    - The function adds an index number to the left of each row for easier reference.
    - The table's appearance is customized using CSS codes to enhance readability and aesthetics.

    Examples
    --------
    >>> df = pd.DataFrame(np.random.rand(4, 25), columns=[f'Var{i}' for i in range(1, 26)])
    >>> df.varnames(rows=5)
    This will display the column names of `df` arranged in a table with 5 rows,
    the number of columns being automatically calculated, and styled with the specified CSS properties.
    """

    if not isinstance(self, pd.DataFrame):
        print('Please declare one dataframe')
    else:
        names = self.columns.to_list()
        if (not rows and not cols) or (rows and cols):
            num_rows = 20
            num_cols = -(-len(names) // num_rows)
        if rows and not cols:
            num_rows = rows
            num_cols = -(-len(names) // num_rows)
        if not rows and cols:
            num_cols = cols
            num_rows = -(-len(names) // num_cols)

        self = [['' for i in range(num_cols)] for j in range(num_rows)]

        for k, name in enumerate(names):
            row = k % num_rows
            col = k // num_rows
            self[row][col] = name
        out = pd.DataFrame(self).style.hide(axis=0).hide(axis=1).set_table_styles([
            {'selector': 'td',
             'props': 'padding: 10px; text-align: center; font-weight: regular; background: lightgoldenrodyellow; border: 1px dotted black;'},
            {'selector': 'tr', 'props': 'width: 100% !important'},
            {'selector': 'table', 'props': 'width: 100% !important'}
        ], overwrite=True)
    return out
