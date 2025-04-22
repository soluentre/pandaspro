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

    :param data: DataFrame to reorder
    :param column: Column name on which to change position, will be switched to the beginning of the DataFrame if param before and after are not specified
    :param before: The column name before which the specified column should be placed (optional)
    :param after: The column name after which the specified column should be placed (optional)

    :return: Reordered DataFrame or None if reordered in place
    """
    if isinstance(column, str):
        cols = [i.strip() for i in column.split(';')]
    elif isinstance(column, list):
        cols = column

    remove_list = []
    for i in cols:
        if i in data.columns:
            pass
        else:
            remove_list.append(i)
    if len(remove_list) != 0:
        print(f'Columns {remove_list} not in the dataframe, removed and new lists updated')
    retain_list = [item for item in cols if item not in remove_list]
    old_order = [i for i in list(data.columns) if i not in retain_list]

    if before:
        index = old_order.index(before)
        new_order = old_order[:index] + retain_list + old_order[index:]
    elif after:
        index = old_order.index(after)
        new_order = old_order[:index+1] + retain_list + old_order[index+1:]
    elif pos == 'end':
        new_order = old_order + retain_list
    else:
        new_order = retain_list + old_order

    return data[new_order]


if __name__ == '__main__':
    import sprnldata as spr
    a = spr.monades()
    b = a.corder('cancel; amount', after='country')