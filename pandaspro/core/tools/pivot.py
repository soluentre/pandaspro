import pandas as pd
from pandaspro.core.tools.corder import corder


#########################################
##使用**kwargs的方式传递参数以及提取
# def fun(**kwargs):
#     if 'aggfunc' in kwargs.keys():
#     pivot_table(**kwargs)



def add_subtotals_above(
        df,
        index,
        columns,
        values,
        aggfunc,
        subtotal: dict
):
    if isinstance(index, str):
        index = [index]

    pivot_df = df.pivot_table(index=index,
                              columns=columns,
                              values=values,
                              aggfunc=aggfunc).reset_index()

    sub_level = list(subtotal.keys())[0]
    col_refill = [col for col in index if col != sub_level][0]
    subtotal_df = df.pivot_table(index=sub_level,
                                columns=columns,
                                values=values,
                                aggfunc=aggfunc).reset_index()
    subtotal_df[col_refill] = 'Subtotal'

    merged_list = []
    for key, group in pivot_df.groupby(sub_level):

        sub_row = subtotal_df[subtotal_df[sub_level] == key]
        print(sub_row, group)

        if list(subtotal.values())[0] == 'top':
            merged_list.append(sub_row)
            merged_list.append(group)
        elif list(subtotal.values())[0] == 'bottom':
            merged_list.append(group)
            merged_list.append(sub_row)

    result_df = pd.concat(merged_list, ignore_index=True)
    result_df = corder(result_df, sub_level)

    return result_df



if __name__ == '__main__':
    data = {
        'employee': ['Alice', 'Bob', 'Charlie', 'David', 'Eva', 'Frank', 'Grace', 'Hannah', 'Ian', 'Julia'],
        'level': ['Senior', 'Senior', 'Junior', 'Senior', 'Senior', 'Junior', 'Junior', 'Senior', 'Junior', 'Junior'],
        'department': ['Sales', 'Marketing', 'Sales', 'HR', 'Sales', 'Sales', 'HR', 'Marketing', 'Marketing', 'Sales'],
        'sales': [50000, 70000, 45000, 60000, 90000, 55000, 85000, 75000, 49000, 95000],
        'location': ['HQ', 'HQ', 'HQ', 'CO', 'HQ', 'CO', 'CO', 'CO', 'CO', 'HQ'],
        'gender': ['Male', 'Male', 'Male', 'Male', 'Female', 'Female', 'Female', 'Male', 'Female', 'Female']
    }

    df = pd.DataFrame(data)
    final_df = add_subtotals_above(df,
                               index=['department', 'level'],
                               columns='gender',
                               values='sales',
                               aggfunc='count',
                               subtotal={'level': 'bottom'})


