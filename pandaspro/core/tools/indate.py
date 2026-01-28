import re
import pandas as pd
from pandas.tseries.offsets import MonthEnd, YearEnd


def indate(df, colname, compare, date, end_date=None, inclusive='both', engine='b', inplace=False, invert=False,
           debug=False, dayfirst=False, format=None):
    mapper = {
        'lt': '<',
        'gt': '>',
        'le': '<=',
        'ge': '>=',
        'eq': '='
    }
    if compare in mapper.keys():
        compare = mapper[compare]

    if compare not in ['<', '>', '<=', '>=', '=', 'between']:
        raise ValueError("Invalid comparison operator")

    if compare == 'between' and inclusive not in ["both", "neither", "left", "right"]:
        raise ValueError("Invalid compare_rule for 'between'")

    if compare != 'between' and end_date is not None:
        raise ValueError("'end_date' should only be used with 'between'")

    # Function to handle fiscal year
    #####################
    def handle_fiscal_year(fy_string):
        match = re.match(r'FY(\d{2})', fy_string)
        if match:
            year = int(match.group(1))
            start_year = 1900 + year - 1 if year > 50 else 2000 + year - 1
            fy_start_date = pd.to_datetime(f'{start_year}-07-01', format='%Y-%m-%d', errors='coerce')
            fy_end_date = pd.to_datetime(f'{start_year + 1}-06-30', format='%Y-%m-%d', errors='coerce')
            return fy_start_date, fy_end_date
        else:
            raise ValueError(f"Invalid fiscal year format: {fy_string}")

    # Check for fiscal year format in date and end_date, and get the right starting point
    #####################
    fytag, fyendtag, fystart, fyend, temp_end, temp_start = 0, 0, 0, 0, 0, 0
    if isinstance(date, str) and date.startswith('FY'):
        fystart, temp_end = handle_fiscal_year(date)
        fytag = 1
    if isinstance(end_date, str) and end_date.startswith('FY'):
        temp_start, fyend = handle_fiscal_year(end_date)
        fyendtag = 1

    # Separate format determination for date and end_date
    #####################
    def determine_format(dt):
        if 'FY' in dt:
            return False, False
        # Example: 2024
        elif '-' not in dt:
            return True, False
        # Example: 2024-06
        elif dt.count('-') == 1:
            return False, True
        else:
            return False, False

    # Date: the main input
    #####################
    is_year, is_month = determine_format(date)
    df[colname] = pd.to_datetime(df[colname], errors='coerce', format=format, dayfirst=dayfirst)
    if fystart == 0:
        date = pd.to_datetime(date, errors='coerce', format=format, dayfirst=dayfirst)
        if is_month and compare in ['<=', '>']:
            date += MonthEnd(1)
        if is_year and compare in ['<=', '>']:
            date += YearEnd(1)
        if compare == 'between' and is_month and inclusive in ['neither', 'right']:
            date += MonthEnd(1)
        if compare == 'between' and is_year and inclusive in ['neither', 'right']:
            date += YearEnd(0)
    else:
        date = fystart

    # End_Date: the second input when between is arguments
    #####################
    # end_date_format, is_end_year, is_end_month = (None, False, False)
    if end_date:
        is_end_year, is_end_month = determine_format(end_date)
        if fyendtag == 0:
            end_date = pd.to_datetime(end_date, errors='coerce', format=format, dayfirst=dayfirst)
            if is_end_month and compare == 'between' and inclusive in ['both', 'right']:
                end_date += MonthEnd(1)
            if is_end_year and compare == 'between' and inclusive in ['both', 'right']:
                end_date += YearEnd(0)
        else:
            if compare == 'between' and inclusive in ['both', 'right']:
                end_date = fyend
            if compare == 'between' and inclusive in ['neither', 'left']:
                end_date = temp_start - pd.Timedelta(days=1)

    if debug:
        print(f"Comparing using {compare} with date {date} and end_date {end_date} with rule of {inclusive}")

    mask = None
    if compare == 'between':
        mask = df[colname].between(date, end_date, inclusive=inclusive)
    else:
        if compare == '<':
            mask = df[colname] < date
        elif compare == '>':
            mask = df[colname] > date if fytag == 0 else df[colname] > temp_end
        elif compare == '<=':
            mask = df[colname] <= date if fytag == 0 else df[colname] <= temp_end
        elif compare == '>=':
            mask = df[colname] >= date
        elif compare == '=':
            if fytag == 1:
                mask = df[colname].between(fystart, temp_end, inclusive='both')
            elif is_month:
                mask = df[colname].between(date, date + MonthEnd(1), inclusive='both')
            elif is_year:
                mask = df[colname].between(date, date + YearEnd(1), inclusive='both')
            elif is_month == False and is_year == False:
                mask = df[colname] == date
        else:
            raise ValueError('Invalid compare argument passed')

    # Mask is successfully generated, next is to create the df
    ##########################################################
    if engine == 'r' or inplace == True:
        if not invert:
            df.drop(df[~mask].index, inplace=True)
        else:
            df.drop(df[mask].index, inplace=True)
    elif engine == 'b':
        return df[mask] if invert == False else df[~mask]
    elif engine == 'm':
        return mask if invert == False else ~mask
    elif engine == 'c':
        if not invert:
            df.loc[mask, '_indate'] = 1
        else:
            df.loc[~mask, '_indate'] = 0
    else:
        print('Unsupported type')
