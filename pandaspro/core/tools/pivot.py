import pandas as pd
from pandaspro.core.tools.corder import corder
from pandaspro.sample_df import df


#########################################
##使用**kwargs的方式传递参数以及提取
# def fun(**kwargs):
#     if 'aggfunc' in kwargs.keys():
#     pivot_table(**kwargs)

def gen_all_level_index(df, columns_list):

    groupid_full = df[columns_list].apply(lambda row: '_'.join(row.values.astype(str)), axis=1)
    unique_groupid_full = sorted(groupid_full.unique())
    current_list = unique_groupid_full

    if len(columns_list) >= 2:
        for i in range(len(columns_list) - 1, 0, -1):
            next_list = []
            unique_groups = df[columns_list[:i]].drop_duplicates()

            for _, group in unique_groups.iterrows():
                base_groupid = '_'.join(group.values.astype(str))

                matching_groupid = [gid for gid in current_list if gid.startswith(base_groupid)]
                next_list.append(f"{base_groupid}_top")
                next_list.extend(matching_groupid)
                next_list.append(f"{base_groupid}_bottom")

            current_list = next_list

    else:
        if isinstance(columns_list, list):
            columns_list = columns_list[0]
        elif isinstance(columns_list, str):
            columns_list = columns_list

        current_list = df[columns_list].drop_duplicates().to_list()

    final_groupids = ['Grand Total_top'] + current_list + ['Grand Total_bottom']

    return pd.DataFrame({"groupid": final_groupids})


def gen_multi_level_pivot(
        df,
        index,
        columns,
        values,
        aggfunc,
        levels_for_pivot: list
):
    if isinstance(index, str):
        index = [index]

    pivot_main = df.pivot_table(index=index,
                              columns=columns,
                              values=values,
                              aggfunc=aggfunc).reset_index()
    pivot_main['id'] = df[columns_list].apply(lambda row: '_'.join(row.values.astype(str)), axis=1)

    pivot_subtotal = pd.DataFrame()
    for level in levels_for_pivot:
        print(level)
        addtional_level_index = index[:index.index(level)+1]
        pivot_additonal_level = df.pivot_table(index=addtional_level_index,
                              columns=columns,
                              values=values,
                              aggfunc=aggfunc).reset_index()
        pivot_additonal_level['id'] = df[addtional_level_index].apply(lambda row: '_'.join(row.values.astype(str)), axis=1)
        pivot_subtotal = pd.concat([pivot_subtotal, pivot_additonal_level])

    result_df = pd.concat([pivot_main, pivot_subtotal])

    return result_df

def gen_pivot_pro(
        df,
        index,
        columns,
        values,
        aggfunc,
        subtotal_level: dict
):
    all_level_index


if __name__ == '__main__':
    final_df = gen_multi_level_pivot(df,
                                     index=['department', 'grade'],
                                     columns='gender',
                                     values='salary',
                                     aggfunc='count',
                                     subtotal={'department': 'bottom'})

    # 根据描述的逻辑生成所需的 groupid 列表

    # 先定义给定的列列表
    columns_list = ["department", 'unit']

    # 调用函数生成符合层级关系的groupid DataFrame
    all_level_index = gen_all_level_index(df, columns_list)


