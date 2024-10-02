'''
let's create a pivot pro:
pivot_table will return a pd.DataFrame object, let's first inherit the DataFrame, then make a child class:
    1. the child class will recognize the header, index, Row total and Column total
    2. we need to create subtotal rows and columns according to the header/index -> we are going to play with multi_index
    3. we also need to create total rows and columns quickly
    4. we need be able to claim a dataframe as this class easily: classname(df, '0-y-y-0)

'''
# mapper_nb = {..para..} + {..position..}
#
# i_give = {'top1': True}
# i_give_parse = {
#     'id1 - deptGA': {
#         'top1': True
#     },
#     'id2 - deptGB': {
#         'top1': True
#     }
#     ...
# }
# from_structure_parse = {
#     'id1 - deptGA': {
#         'top1': True,
#     },
#     'id2 - deptGB': {
#         'top1': False,
#     }
# }
#
# for id in grpn_1:  ## id == deptGA is a string not a numeric
#     df_basic = [df.inlist('grpn_1', id)]
#     df_tops = []
#     for i in range(n-1, 1):
#         top_row = df.pivot(...n-1...).inlist('grpn-1', id)
#         if mapper_nb[id]['top'][n-1]['Yes_or_No']:  ## this is mapper_nb
#             df_tops = df_tops.append(top_row)
