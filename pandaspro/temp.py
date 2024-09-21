from pandaspro.core.tools.corder import corder


def unify_to_list(para: int | str | list):
    if isinstance(para, int | str):
        para_list = [para]
    elif isinstance(para, list):
        para_list = para
    else:
        raise ValueError('Incorrect type of argument [para]: only supports int/str/list')
    return para_list


class cpdConsecGrouper:
    def __init__(
            self,
            df,
            group_by,
    ):
        self.df = df
        self.df_grouped = None

        # Split Section
        if isinstance(group_by, str):
            self.group_by = [group_by]
        elif isinstance(group_by, list):
            self.group_by = group_by
        else:
            raise ValueError('Incorrect type for group_by, can only be str/list')

    # noinspection PyShadowingNames
    def group(self):
        self.df_grouped = self.df.copy()
        self.df_grouped['group_num'] = (self.df_grouped[self.group_by] != self.df_grouped[self.group_by].shift()).any(axis=1).cumsum()
        self.df_grouped['group_id'] = self.df_grouped[['group_num'] + self.group_by].astype(str).agg('-'.join, axis=1)
        self.df_grouped = corder(self.df_grouped, ['group_num', 'group_id'])
        return self.df_grouped

    def extract(
            self,
            value_at_top: str | list = None,
            value_at_bottom: str | list = None,
    ):
        grouped_df = self.group()
        group_id_dict = {'group_num': 'first', 'group_id': 'first'}
        group_by_dict = {var: 'first' for var in unify_to_list(self.group_by)}
        top_dict = {var: 'first' for var in unify_to_list(value_at_top)}
        bottom_dict = {var: 'last' for var in unify_to_list(value_at_bottom)}

        agg_dict = {}
        for d in [group_id_dict, group_by_dict, top_dict, bottom_dict]:
            for key, value in d.items():
                if key not in agg_dict:
                    agg_dict[key] = []
                agg_dict[key].append(value)
        print(agg_dict)

        grouped_df = grouped_df.groupby('group_num').agg(agg_dict).reset_index(drop=True)
        grouped_df.columns = [col[0] if col[0] in unify_to_list(['group_num', 'group_id'] + self.group_by)
                              else '_'.join(filter(None, col)).rstrip('_')
                              for col in grouped_df.columns]
        grouped_df.columns = [col.replace('first', 'top').replace('last', 'bottom') for col in grouped_df.columns]

        return grouped_df


if __name__ == '__main__':
    # mona = cpd.pwread(
    #     r'C:\Users\xli7\OneDrive - International Monetary Fund (PRD)\Databases\MONA\Description\Description_20240328.xlsx')[
    #     0]
    #
    # a = Spliter(mona)
    # b = a.split(split_by=['arrangement_number', 'review_sequence'])
    # c = a.smart_group(sort_by=['arrangement_number', 'board_action_date'],
    #                   split_by=['arrangement_number', 'review_sequence'],
    #                   top=['review_type', 'board_action_date'],
    #                   bottom='review_type',
    #                   pick_group=[1, '2-501-BLANK'])
    from wbhrdata import *
    data = psoftjob()
    r = cpdConsecGrouper(data, ['upi', 'location']).extract(value_at_top=['eff_dt'], value_at_bottom='eff_dt')
    r = r.sort_values('group_num')
    # b = a.group()

