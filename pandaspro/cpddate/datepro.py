import maya
import pandas as pd
from datetime import datetime


class DatePro:
    map = {
        'BdY': '%B %d, %Y',
        'bdY': '%b %d, %Y',
        'BY': '%B, %Y',
        'bY': '%b, %Y',
        'BdYdash': '%B-%d-%Y',
        'bdYdash': '%b-%d-%Y',
    }

    def __init__(self, date):
        self.original_date = date
        if isinstance(date, pd.Timestamp):
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
        if item == 'weekday':
            return self.dt.weekday()
        elif item == 'isoweekday':
            return self.dt.isoweekday()
        elif item in DatePro.map.keys():
            return self.dt.strftime(DatePro.map[item])
        elif hasattr(self.dt, item):
            return getattr(self.dt, item)
        elif item in ['monthB', 'monthb', 'dayA', 'daya']:
            parse_format = item[-1]
            return self.dt.strftime('%' + parse_format)

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


if __name__ == '__main__':
    # print(DatePro('2024-1-1').BdY1)
    d = DatePro('2024-1-7')
    print(d.BdY)
