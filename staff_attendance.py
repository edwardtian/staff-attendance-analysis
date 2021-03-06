#!/usr/bin/env python3
import sys
import re
import datetime as dt
import pandas as pd

def handle_datetime_column(col):
    result = None
    if type(col) is dt.datetime:
        result = col
    else:
        try:
            result = dt.datetime.fromisoformat(col)
        except ValueError:
            pass
    return result

def handle_date_column(col):
    if type(col) is dt.datetime:
        return col.strftime('%Y-%m-%d')
    else:
        return col

def handle_time_column(col):
    result = None
    if type(col) is dt.time:
        result = col
    else:
        try:
            result = dt.datetime.strptime(col, '%H:%M:%S').time()
        except ValueError:
            pass
    return result

def time_diff_by_minute(time1, time2):
    if type(time1) is dt.time and type(time2) is dt.time:
        return (time1.hour - time2.hour) * 60 + (time1.minute - time2.minute)
    else:
        return None

def merge_possible_overlayed_abrs(absence_records):
    new_abrs = []
    for index, abr in absence_records.iterrows():
        if len(new_abrs) > 0:
            if abr['dt_from'].timestamp() - new_abrs[len(new_abrs) - 1]['dt_to'].timestamp() < 0.1:
                new_abrs[len(new_abrs) - 1]['dt_to'] = abr['dt_to']
            else:
                new_abrs.append({'dt_from': abr['dt_from'], 'dt_to': abr['dt_to']})
        else:
            new_abrs.append({'dt_from': abr['dt_from'], 'dt_to': abr['dt_to']})
    return new_abrs

def calc_actual_duty_time(on_duty_time, off_duty_time, absence_records, is_verbose=False):
    last_abr_dt_from = None
    last_abr_dt_to = None
    
    for abr in merge_possible_overlayed_abrs(absence_records):
        if is_verbose:
            print(f'on_duty_time={on_duty_time}, off_duty_time={off_duty_time}')
            print(f'on_duty_time.timestamp()={on_duty_time.timestamp()}, off_duty_time.timestamp()={off_duty_time.timestamp()}')
        abr_dt_from = abr['dt_from'].to_pydatetime()
        abr_dt_to = abr['dt_to'].to_pydatetime()
        if is_verbose:
            print(f'abr_dt_from={abr_dt_from}, abr_dt_to={abr_dt_to}')
            print(f'abr_dt_from.timestamp()={abr_dt_from.timestamp()}, abr_dt_to.timestamp()={abr_dt_to.timestamp()}')
        if on_duty_time.timestamp() < abr_dt_from.timestamp():
            if (off_duty_time.timestamp() > abr_dt_from.timestamp()) and (off_duty_time.timestamp() <= abr_dt_to.timestamp()): 
                if is_verbose:
                    print('????????????????????????')
                off_duty_time = abr_dt_from # ????????????????????????
        elif on_duty_time.timestamp() < abr_dt_to.timestamp():
            if off_duty_time.timestamp() <= abr_dt_to.timestamp(): # ????????????????????????
                if is_verbose:
                    print('????????????????????????')
                on_duty_time = None #dt.datetime.fromisoformat('2099-12-31T23:59:00')
                off_duty_time = None #dt.datetime.fromisoformat('2000-01-01T00:00:00')
                break
            else:
                if is_verbose:
                    print('????????????????????????')
                on_duty_time = abr_dt_to # ????????????????????????

    if is_verbose:
        print(f'actual_on_duty_time={on_duty_time}, actual_off_duty_time={off_duty_time}')
        print('-'*80)
    return \
        on_duty_time.time() if type(on_duty_time) is dt.datetime else on_duty_time, \
        off_duty_time.time() if type(off_duty_time) is dt.datetime else off_duty_time

def mapping_color(cell_value):
    bg_color = ''
    default = ''
    new_style = ''

    if type(cell_value) is str:
        if   '??????' in cell_value:
            bg_color = 'violet'
        elif '??????' in cell_value:
            bg_color = 'green'
        elif '??????' in cell_value:
            bg_color = 'yellow'
        elif '????????????' in cell_value:
            bg_color = 'red'

        if not bg_color == '':
            new_style = f'background-color: {bg_color};'

    return new_style

if len(sys.argv) != 3:
    print('Please provide excel file to be processed.')
    print('Usage: ')
    print(f'    python3 {sys.argv[0]} <OA Excel File> <HR Excel File>')
    exit()

print(f'Handling {sys.argv[1]}(as OA data) and {sys.argv[2]}(as HR data) ...')

in_oa_df = pd.read_excel(sys.argv[1], skiprows=5, usecols='A:T', 
                         converters={'??????': handle_date_column}
                        )
dates = in_oa_df['??????'].unique().tolist()
#in_oa_df = in_oa_df.sort_values(by='??????', ascending=True)

in_hr_df = pd.read_excel(sys.argv[2], skiprows=1, header=None, usecols=[3,4,5], 
                         names=['name', 'dt_from', 'dt_to'], 
                         converters={'name': lambda n : re.search(r'^([^A-Za-z ]+)', n).group(1) if re.search(r'^([^A-Za-z ]+)', n) else n}
                        ).sort_values(by='dt_from', ascending=True)

out_df = pd.DataFrame(index=[], columns=(['??????', '??????'] + dates))

i = 0
cur_user = None

while(i < len(in_oa_df)):
    row = in_oa_df.loc[i, :]
    if row['?????????'] == '(-)':
        i += 1
        continue

    cur_user = re.sub(r'[??? ]+', '', row['??????']) if re.search(r'[^A-Za-z ]', row['??????']) else row['??????']
    if not (cur_user in out_df.index):
        print(f'Processing {cur_user}')
        out_df.loc[cur_user] = [None] * len(out_df.columns)
        out_df.loc[cur_user, '??????'] = row['??????']
        out_df.loc[cur_user, '??????'] = row['??????']

    result = 1
    result_desc = ''

    absence_records = in_hr_df.loc[lambda df: df['name'] == cur_user]

    on_duty_time = handle_datetime_column(row['??????'] + ' ' + re.search(r'([0-9:]+)-', row['?????????']).group(1))
    off_duty_time = handle_datetime_column(row['??????'] + ' ' + re.search(r'-([0-9:]+)', row['?????????']).group(1))

    on_duty_time, off_duty_time = calc_actual_duty_time(on_duty_time, off_duty_time, absence_records) # DEBUG_PARAMS: , cur_user == '??????'

    checkin_time = handle_time_column(row['????????????'])
    checkout_time = handle_time_column(row['????????????'])

    if(on_duty_time == None and off_duty_time == None):
        result_desc = '??????'
    else:
        if type(checkin_time) is dt.time and type(on_duty_time) is dt.time:
            if(checkin_time > on_duty_time):
                result *= 2
                result_desc += f'????????????{time_diff_by_minute(checkin_time, on_duty_time)}??????,'
            else:
                result_desc += '????????????,'
        else:
            result *= 3
            result_desc += '????????????,'
        if type(checkout_time) is dt.time and type(off_duty_time) is dt.time:
            if checkout_time < off_duty_time:
                result *= 5
                result_desc += f'????????????{time_diff_by_minute(off_duty_time, checkout_time)}??????,'
            else:
                result_desc += '????????????'
        else:
            result *= 7
            result_desc += '????????????'
        if result == 1:
            result_desc = '??????'
        if result % (3 * 7) == 0:
            result_desc = '??????'
    
    out_df.loc[cur_user, row['??????']] = result_desc
    i += 1

out_df.style.applymap(mapping_color).to_excel('output.xlsx', sheet_name='????????????', startrow=0, startcol=0, index=True)