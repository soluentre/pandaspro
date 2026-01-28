import pandas as pd


def tab(
    data,
    name: str | list,
    d: str = 'brief',
    m: bool = False,
    sort: str = 'index',
    ascending: bool | list = True,
    label: str = None,
):
    """
    The `tab` function provides various tabulations of a specified column in a DataFrame,
    with options for including missing values, and sorting the results by
    index, value count, or percentage.

    It supports three modes of output: 'brief', 'detail', and 'export', each offering a
    different level of information and formatting suited for quick review,
    detailed analysis, or preparation for export, respectively.

    Parameters
    ----------
    data : DataFrame
        The DataFrame containing the data to be tabulated.
    name : str
        The name of the column to tabulate.
    d : str, optional and default ='brief'
        The detail level of the output. Options are
            - 'brief' (default)
            - 'detail'
            - 'export'
    m : bool or 'missing', optional and default =False
        If True or 'missing', includes missing values in the tabulation. Defaults to False.
    sort : str, optional, default =index
        The criterion for sorting the results. Options are
            - 'index' (default)
            - 'value'
            - 'percent'.
    ascending : bool, optional and default =True
        Determines the sorting order. Defaults to True (ascending).

    Returns
    -------
    DataFrame
        A DataFrame containing the tabulated data. The structure of the DataFrame varies
        depending on the `d` parameter:
        - 'brief': Returns counts sorted according to the `sort` parameter.
        - 'detail': Returns counts along with their percentage of the total and
                    cumulative percentage, sorted as specified.
        - 'export': Similar to 'detail', but formatted for export, with a total row at the bottom.

    Examples
    --------
    >>> df = pd.DataFrame({'A': [1, 2, 2, np.nan]})
    >>> tab(df, 'A', d='detail', m=True, sort='percent', ascending=False)
    Returns a detailed tabulation of column 'A', including missing values, sorted by percentage in descending order.
    """

    if isinstance(name, list):
        ascending = [True for i in range(len(name))]

    sort_dict = {
        f'{name}': name,
        'index': name,
        'percent': 'Percent'
    }
    if m == 'missing' or m == True:
        df = data[name].value_counts(dropna=False).sort_index().to_frame()
    else:
        df = data[name].value_counts().sort_index().to_frame()

    if d == 'brief':
        # Sort
        if sort == 'index':
            df = df.sort_index(ascending=ascending)
        else:
            df = df.sort_values(sort_dict[sort], ascending=ascending)
        
        # Add Total row
        df = df.reset_index()
        if isinstance(name, list):
            total_row = pd.Series({**{col: 'Total' for col in name}, 'count': df['count'].sum()})
        else:
            total_row = pd.Series({
                name: 'Total',
                'count': df['count'].sum()
            })
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
        df = df.set_index(name)

    elif d == 'detail':
        # Calculate Percent and Cumulative Percent
        df = df.reset_index()
        df['Percent'] = (df['count'] / df['count'].sum() * 100).round(2)

        # Sort
        df = df.sort_values(sort_dict[sort], ascending=ascending)
        df['Cum.'] = df['Percent'].cumsum().round(2)

        # Create a Total row
        total_row = pd.Series({
            name: 'Total',
            'count': df['count'].sum(),
            'Percent': 100.00,
            'Cum.': ''
        })

        # Concatenate the Total row to the DataFrame
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    elif d == 'export':
        df = df.reset_index()
        df['Percent'] = (df['count'] / df['count'].sum()).round(3)
        total_row = pd.Series({
            name: 'Total',
            'count': df['count'].sum(),
            'Percent': 1
        })

        # Sort
        df = df.sort_values(sort_dict[sort], ascending=ascending)

        # Concatenate the Total row to the DataFrame
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
        df.columns = [name, 'Count', 'Percent']

    if label:
        df = df.rename(columns={'count': label})

    return df
