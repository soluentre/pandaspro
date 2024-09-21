import pandas as pd
import pandaspro as cpd


def unify_to_list(para):
    if isinstance(para, str | int):
        para_list = [para]
    if isinstance(para, list):
        para_list = para
    return para_list


class Spliter:
    def __init__(self, df: pd.DataFrame | cpd.FramePro):
        self.df = cpd.FramePro(df)
        self.split_result = None

    # noinspection PyShadowingNames
    def split(self,
              sort_by: str | list = None,
              split_by: str | list = None,
              ):

        if sort_by:
            df_sorted = self.df.sort_values(by=sort_by).reset_index(drop=True)
        else:
            df_sorted = self.df

        if split_by:
            if isinstance(split_by, str):
                split_by_list = [split_by]
            if isinstance(split_by, list):
                split_by_list = split_by

            df_sorted['group_num'] = (df_sorted[split_by_list] != df_sorted[split_by_list].shift()).any(axis=1).cumsum()
            df_sorted['group_id'] = df_sorted[['group_num'] + split_by_list].astype(str).agg('-'.join, axis=1)
            df_sorted = df_sorted.corder(['group_num', 'group_id'])

            self.split_result = df_sorted
            return df_sorted
        else:
            return df_sorted

    def smart_group(self,
                    sort_by: str | list = None,
                    split_by: str | list = None,
                    top: str | list = None,
                    bottom: str | list = None,
                    pick_group: int | str | list = None
                    ):
        grouped_df = self.split(sort_by=sort_by, split_by=split_by)

        group_id_dict = {'group_num': 'first', 'group_id': 'first'}
        split_by_dict = {var: 'first' for var in unify_to_list(split_by)}
        top_dict = {var: 'first' for var in unify_to_list(top)}
        bottom_dict = {var: 'last' for var in unify_to_list(bottom)}

        agg_dict = {}
        for d in [group_id_dict, split_by_dict, top_dict, bottom_dict]:
            for key, value in d.items():
                if key not in agg_dict:
                    agg_dict[key] = []
                agg_dict[key].append(value)
        print(agg_dict)

        grouped_df = grouped_df.groupby('group_id').agg(agg_dict).reset_index(drop=True)
        grouped_df.columns = [col[0] if col[0] in unify_to_list(['group_num', 'group_id'] + split_by)
                              else '_'.join(filter(None, col)).rstrip('_')
                              for col in grouped_df.columns]
        grouped_df.columns = [col.replace('first', 'top').replace('last', 'bottom') for col in grouped_df.columns]

        if pick_group:
            return grouped_df[grouped_df['group_num'].isin(unify_to_list(pick_group)) | grouped_df['group_id'].isin(unify_to_list(pick_group))]

        else:
            return grouped_df


if __name__ == '__main__':
    mona = cpd.pwread(
        r'C:\Users\xli7\OneDrive - International Monetary Fund (PRD)\Databases\MONA\Description\Description_20240328.xlsx')[
        0]

    a = Spliter(mona)
    b = a.split(split_by=['arrangement_number', 'review_sequence'])
    c = a.smart_group(sort_by=['arrangement_number', 'board_action_date'],
                      split_by=['arrangement_number', 'review_sequence'],
                      top=['review_type', 'board_action_date'],
                      bottom='review_type',
                      pick_group=[1, '2-501-BLANK'])
