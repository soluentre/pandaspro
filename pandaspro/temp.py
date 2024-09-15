import pandas as pd
import pandaspro as cpd

def smart_group(
        df: cpd.FramePro = None,
        sort_by: str | list = None,
        split_by: str = None,
        var_of_interest: str | list = None,
        top: bool = True,
        bottom: bool = False,
):
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
    grouped_df.columns = [col[0] if col[0] in [sort_by_key, split_by]
                          else '_'.join(filter(None, col)).rstrip('_')
                          for col in grouped_df.columns]
    grouped_df.columns = [col.replace('first', 'top').replace('last', 'bottom') for col in grouped_df.columns]

    return grouped_df

class test(cpd.FramePro):
    def __init__(self,
                 df: cpd.FramePro = None,
                 sort_by: str | list = None,
                 split_by: str = None,
                 var_of_interest: str | list = None,
                 top: bool = True,
                 bottom: bool = False,
                 *args, **kwargs):

        if args or kwargs:
            super().__init__(*args, **kwargs)
        else:
            grouped_df = smart_group(df=df,
                                     sort_by=sort_by,
                                     split_by=split_by,
                                     var_of_interest=var_of_interest,
                                     top=top,
                                     bottom=bottom)
            super().__init__(grouped_df)

if __name__ == '__main__':

    # mona = cpd.pwread(r'C:\Users\xli7\OneDrive - International Monetary Fund (PRD)\Databases\MONA\Description\Description_20240328.xlsx')[0]
    # b = test(mona, sort_by=['arrangement_number', 'board_action_date'], split_by='review_sequence', var_of_interest=['review_type', 'board_action_date'], bottom=True)

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

    from wbhrdata import *
    d = psoftjob()
    df = d.upi83315

    this_class = test(
        df=df,
        sort_by='eff_dt',
        split_by='location',
        var_of_interest=['eff_dt', 'salary'],
        top=True,
        bottom=True,
    )
