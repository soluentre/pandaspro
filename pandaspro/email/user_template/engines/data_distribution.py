import pandas as pd
import pandaspro as cpd

country_info = cpd.pwread(r'pandaspro/email/user_template/email_list.xlsx', sheet_name='info')[0]
country_info['email_list'] = country_info.apply(lambda row: '; '.join(filter(pd.notna, [row['email1'], row['email2'], row['email3'], row['email4']])), axis=1)

# noinspection PyTypedDict
class data_distribution(cpd.emailfetcher):
    # noinspection PyAttributeOutsideInit
    def fetch_data(self, ifscode=None, year=2024):
        self.input = {}
        self.data_pick = {}
        for col in ['country', 'email_list', 'ngdpd', 'status']:
            self.data_pick[col] = country_info.set_index('ifscode').at[ifscode, col]

        self.input['__[COUNTRY]__'] = self.data_pick['country']
        self.input['__[NGDPD]__'] = round(self.data_pick['ngdpd'], 1)
        self.input['__[NOTE]__'] = f"In the case of {self.data_pick['country']}, the GDP for {year} is {round(self.data_pick['ngdpd'], 1)}."

        # Conditional content
        if self.data_pick['status'] == 'Yes':
            self.input['__[STATUS]__'] = f"Please confirm if the status is YES."
        else:
            self.input['__[STATUS]__'] = f"Please confirm if the status is NO."

        # Subject, To and CC
        self.input['__subject__'] = 'GDP data'
        self.input['__to__'] = self.data_pick['email_list']
        self.input['__cc__'] = "cc_email@cpd.com"

        return self.input

    # noinspection PyAttributeOutsideInit
    def fetch_showitems(self, **kwargs):
        self.render_dict = {}
        if self.data_pick['status'] == 'Yes':
            self.render_dict['show_1'] = True
            self.render_dict['show_2'] = False
        else:
            self.render_dict['show_1'] = False
            self.render_dict['show_2'] = True
        return self.render_dict
