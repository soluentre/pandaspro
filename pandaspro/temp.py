import pandas as pd
import pandaspro as cpd

mona = cpd.pwread(r'C:\Users\xli7\OneDrive - International Monetary Fund (PRD)\Databases\MONA\Description\Description_20240328.xlsx')[0]

def test(
        df,
        sort_by: str | list,
        split_by: str,
        var_of_interest: str | list,
        top: bool = True,
        bottom: bool = False):

    df_sorted = df.sort_values(by=sort_by).reset_index(drop=True)
    df_sorted['group'] = (df_sorted[split_by] != df_sorted[split_by].shift()).cumsum()

    # sort_by
    if isinstance(sort_by, str):
        sort_by_key = sort_by
    elif isinstance(sort_by, list):
        sort_by_key = sort_by[0]
    else:
        raise ValueError('sort_by parameter only takes string or list.')

    # top/bottom parameter
    if top & bottom:
        top_bottom = ['first', 'last']
    elif top:
        top_bottom = 'first'
    elif bottom:
        top_bottom = 'last'
    else:
        raise ValueError('Please specify True value for at least one parameter between top and bottom.')

    if isinstance(var_of_interest, str):
        var_dict = {var_of_interest: top_bottom}
    elif isinstance(var_of_interest, list):
        var_dict = {
            var: top_bottom for var in var_of_interest
        }
    else:
        raise ValueError('var_of_interest parameter only takes string or list.')
    print(var_dict)

    agg_dict = {
        sort_by_key: 'first',
        split_by: 'first',
    }
    agg_dict = agg_dict | var_dict
    print(agg_dict)

    grouped_df = df_sorted.groupby('group').agg(agg_dict).reset_index(drop=True)

    ren_dict = {
        f'{roster}/first': roster for roster in [sort_by_key, split_by] if f'{roster}/first' in grouped_df.columns
    }
    var_ren_top = {
        f'{var}/first': f'{var}_top' for var in var_dict.keys() if f'{var}/first' in grouped_df.columns
    }
    var_ren_bottom = {
        f'{var}/last': f'{var}_bottom' for var in var_dict.keys() if f'{var}/last' in grouped_df.columns
    }
    ren_dict = ren_dict | var_ren_top | var_ren_bottom
    print(ren_dict)

    grouped_df = grouped_df.rename(columns=ren_dict)
    return grouped_df

b = test(mona, sort_by=['arrangement_number', 'board_action_date'], split_by='review_sequence', var_of_interest=['review_type', 'board_action_date'], bottom=True)

# import pandas as pd
#
# data = {
#     'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
#     'Age': [24, 27, 22, 32, 24],
#     'City': ['New York', 'Los Angeles', 'Chicago', 'Houston', 'Phoenix'],
#     'Salary': [70000, 80000, 65000, 90000, 75000]
# }
#
# sample_df = pd.DataFrame(data)
#
# sample_df['s2'] = sample_df['Salary'].shift(2)
# sample_df['s1'] = sample_df['Salary'].shift(1)
# sample_df['s0'] = sample_df['Salary'].shift(0)
# sample_df['s_1'] = sample_df['Salary'].shift(-1)