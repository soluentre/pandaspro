import maya
import pandas as pd
from datetime import datetime


class DatePro:
    @staticmethod
    def get_strftime_format(attr_name):
        result = attr_name.replace('_c', ',').replace('_s', ' ').replace('_u', '_')
        result = ''.join(['%' + char if char.isalpha() else char for char in result])
        return result

    def __init__(self, date='today', format=None):
        self.original_date = date
        if format is not None:
            self.maya = maya.parse(datetime.strptime(date, format))
        else:
            if date == 'today':
                self.maya = maya.parse(datetime.now())
                self.date = 'today'
            elif isinstance(date, pd.Timestamp):
                self.maya = maya.parse(str(date))
                self.datetype = 'pd.Timestamp'
            elif isinstance(date, datetime):
                self.maya = maya.MayaDT.from_datetime(date)
                self.datetype = 'datetime'
            elif isinstance(date, str):
                self.maya = maya.parse(date)
                self.datetype = 'str'
            else:
                raise ValueError('Invalid type for date passed, only support [pd.Timestamp, str] objects for this version')

        self.dt = self.maya.datetime()

    def __getattr__(self, item):
        if item == 'readable':
            return self.dt.strftime(DatePro.get_strftime_format('b_sd_c_sY'))
        elif item == 'simple':
            return self.dt.strftime(DatePro.get_strftime_format('b_sd_sY'))
        elif item == 'detail':
            return self.dt.strftime(DatePro.get_strftime_format('b_sd_c_sY_sH_sM'))
        elif item == 'iso':
            return self.dt.strftime(DatePro.get_strftime_format('Ymd_uH_uM_uS'))
        elif item == 'weekday':
            return self.dt.weekday()
        elif item == 'isoweekday':
            return self.dt.isoweekday()
        elif hasattr(self.dt, item):
            return getattr(self.dt, item)
        elif item in ['monthB', 'monthb', 'dayA', 'daya']:
            parse_format = item[-1]
            return self.dt.strftime('%' + parse_format)
        elif not item.startswith('_'):
            return self.dt.strftime(DatePro.get_strftime_format(item))

    def __repr__(self):
        return f"DatePro(date={self.original_date}, datetype={self.datetype}, dt={self.dt})"

    def __str__(self):
        return f"DatePro: {self.dt.strftime('%Y-%m-%d %H:%M:%S')} (original: {self.original_date})"

    @staticmethod
    def help():
        print('DatePro object supports ... ')
        print('.original_date: to get the input object')
        print('.datetype: to get the input format type')
        print('.maya: to get the mayaDT object for a date')
        print('.dt: to get the parsed datetime object for a date')
        print('-------------------')
        print('Almost all traditional attributes like year, month, day, weekday are available, too')
        print('Plus monthB, monthb, dayA and daya for humanized strings')
        print('And the following map applies ...')
        print('')
        print('>>>')
        for key in DatePro.map.keys():
            print(f'{key} = using {DatePro.map[key]}, like << {getattr(DatePro("2020-1-1"), key)} >>')

    @property
    def wbfy(self, short: bool = True) -> str:
        fy_year = self.dt.year + 1 if self.dt.month >= 7 else self.dt.year
        return f"FY{fy_year % 100:02d}" if short else f"FY{fy_year}"


if __name__ == '__main__':
    # print(DatePro('2024-1-1').BdY1)
    print(DatePro('202401', format='%Y%m').readable)

    # def get_strftime_format(attr_name):
    #     result = attr_name.replace('_c', ',').replace('_s', ' ')
    #     result = ''.join(['%' + char if char.isalpha() else char for char in result])
    #     return result
    #
    #
    # print(get_strftime_format('bd_c_sy'))