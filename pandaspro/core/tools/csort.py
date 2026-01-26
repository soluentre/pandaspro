import pandas as pd
from pandaspro.core.tools.utils import df_with_index_for_mask


def csort(
        data,
        column,
        order: str | list = None,
        where: str = 'first',
        before: str = None,
        after: str = None,
        inplace: bool = False
):
    """
    Sorts the DataFrame by the given column according to a custom or dynamically generated order.
    Automatically completes the orderlist with all unique values in the column if it's partially provided.
    Optionally, positions rows with a specified value before or after another value.

    :param data: DataFrame to sort
    :param column: Column name on which to sort
    :param order: List defining the custom order, dynamically completed if partially provided
    :param where: How the value will be put: either in the first or last place
    :param before: The value before which the specified value should be placed (optional)
    :param after: The value after which the specified value should be placed (optional)
    :param inplace: Whether to sort the DataFrame in place (default: False)

    :return: Sorted DataFrame or None if sorted in place

    """
    if column in data.columns:
        pass
    elif column in data.index.names:
        data = df_with_index_for_mask(data)
    else:
        raise ValueError(f'Column {column} not found in either the dataframe nor the index namelist')

    orderlist = list(data[column].dropna().unique())

    if isinstance(order, list):
        provided_reorder = [x for x in order]
        missing_reorder = [x for x in data[column].dropna().unique() if x not in order]
        full_orderlist = provided_reorder + missing_reorder
        orderlist = full_orderlist

    elif isinstance(order, str):
        value = order
        # Order for certain value
        if value and not before and not after:
            if value in orderlist:
                orderlist.remove(value)
            if where == 'first':
                orderlist.insert(0, value)
            elif where == 'last':
                orderlist.append(value)

        # Reorder the list if value and before/after are specified
        if value and (before or after):
            if before and (value in orderlist) and (before in orderlist):
                orderlist.remove(value)
                before_index = orderlist.index(before)
                orderlist.insert(before_index, value)
            elif after and (value in orderlist) and (after in orderlist):
                orderlist.remove(value)
                after_index = orderlist.index(after) + 1
                orderlist.insert(after_index, value)

    cat_type = pd.CategoricalDtype(categories=orderlist, ordered=True)
    data.loc[:, "__cpd_sort"] = data[column].astype(cat_type)

    if inplace:
        data.sort_values(by='__cpd_sort', inplace=True, kind='mergesort')
        if set(data.index.names) <= set(data.columns):
            data.drop(list(data.index.names) + ['__cpd_sort'], axis=1, inplace=True)
    else:
        result = data.sort_values(by='__cpd_sort', kind='mergesort')
        if result.index.names[0] is None:
            result.drop('__cpd_sort', axis=1, inplace=True)
        elif set(result.index.names) <= set(result.columns):
            result.drop(list(data.index.names) + ['__cpd_sort'], axis=1, inplace=True)

        return result
