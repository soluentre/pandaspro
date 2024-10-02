import pandas as pd
from pandaspro.core.tools.corder import corder
from pandaspro.sample_df import df


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
    final_df = add_subtotals_above(df,
                               index=['department', 'grade'],
                               columns='gender',
                               values='salary',
                               aggfunc='count',
                               subtotal={'grade': 'bottom'})

    # 根据描述的逻辑生成所需的 groupid 列表

    # 先定义给定的列列表
    columns_list = ["department"]


    # 调整代码以按照每一层级插入top和bottom
    def generate_groupids_correct_order(df, columns_list):

        # 初始完整层级
        groupid_full = df[columns_list].apply(lambda row: '_'.join(row.values.astype(str)), axis=1)
        unique_groupid_full = sorted(groupid_full.unique())

        # 从上往下逐层处理

        current_list = unique_groupid_full

        # 迭代每个层级，
        if len(columns_list) >= 2:
            for i in range(len(columns_list)-1, 0, -1):
                next_list = []
                unique_groups = df[columns_list[:i]].drop_duplicates()

                # 按每个分组生成top和bottom，并将其插入到对应位置
                for _, group in unique_groups.iterrows():
                    base_groupid = '_'.join(group.values.astype(str))

                    # 定位该分组在当前列表中的所有条目
                    matching_groupid = [gid for gid in current_list if gid.startswith(base_groupid)]

                    # 插入 top groupid
                    next_list.append(f"{base_groupid}_top")

                    # 将当前分组的所有groupid按顺序插入
                    next_list.extend(matching_groupid)

                    # 插入 bottom groupid
                    next_list.append(f"{base_groupid}_bottom")

                # 更新current_list
                current_list = next_list

        else:
            if isinstance(columns_list, list):
                columns_list = columns_list[0]
            elif isinstance(columns_list, str):
                columns_list = columns_list

            current_list = df[columns_list].drop_duplicates().to_list()

        final_groupids = ['Grand Total_top'] + current_list + ['Grand Total_bottom']

        return pd.DataFrame({"groupid": final_groupids})


    # 调用函数生成符合层级关系的groupid DataFrame
    groupid_correct_order_df = generate_groupids_correct_order(df, columns_list)


