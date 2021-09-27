#!/usr/bin/env python3
import sys
import re
import datetime as dt
import pandas as pd

def time_diff_by_minute(time1, time2):
    if type(time1) is dt.time and type(time2) is dt.time:
        return (time1.hour - time2.hour) * 60 + (time1.minute - time2.minute)
    else:
        return None

def calc_actual_duty_time(on_duty_time, off_duty_time, absence_records, is_verbose=False):
    actual_duty_time = { 'on':  on_duty_time, 'off': off_duty_time }
    if is_verbose:
        print(f'on_duty_time={on_duty_time}, off_duty_time={off_duty_time}')
        print(f'on_duty_time.timestamp()={on_duty_time.timestamp()}, off_duty_time.timestamp()={off_duty_time.timestamp()}')
    for index, abr in absence_records.iterrows():
        abr_dt_from = abr['dt_from'].to_pydatetime()
        abr_dt_to = abr['dt_to'].to_pydatetime()
        if is_verbose:
            print(f'abr_dt_from={abr_dt_from}, abr_dt_to={abr_dt_to}')
            print(f'abr_dt_from.timestamp()={abr_dt_from.timestamp()}, abr_dt_to.timestamp()={abr_dt_to.timestamp()}')
        if on_duty_time.timestamp() < abr_dt_from.timestamp():
            if (off_duty_time.timestamp() > abr_dt_from.timestamp()) and (off_duty_time.timestamp() < abr_dt_to.timestamp()): # 正常上班，早下班
                if is_verbose:
                    print('正常上班，早下班')
                actual_duty_time['off'] = abr_dt_from
        elif on_duty_time.timestamp() < abr_dt_to.timestamp():
            if off_duty_time.timestamp() <= abr_dt_to.timestamp(): # 休假，不用上下班
                if is_verbose:
                    print('休假，不用上下班')
                actual_duty_time['on'] = None #dt.datetime.fromisoformat('2099-12-31T23:59:00')
                actual_duty_time['off'] = None #dt.datetime.fromisoformat('2000-01-01T00:00:00')
            else:
                if is_verbose:
                    print('晚上班，正常下班')
                actual_duty_time['on'] = abr_dt_to # 晚上班，正常下班

    if is_verbose:
        print(f'actual_duty_time[on]={actual_duty_time["on"]}, actual_duty_time[off]={actual_duty_time["off"]}')
        print('-'*80)
    return actual_duty_time['on'], actual_duty_time['off']

def mapping_color(cell_value):
    bg_color = ''
    default = ''
    new_style = ''

    if type(cell_value) is str:
        if   '旷工' in cell_value:
            bg_color = 'violet'
        elif '迟到' in cell_value:
            bg_color = 'green'
        elif '早退' in cell_value:
            bg_color = 'yellow'
        elif '下班缺卡' in cell_value:
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

in_oa_df = pd.read_excel(sys.argv[1], skiprows=4, usecols='A:T', 
                         converters={'日期': lambda d : d.strftime('%Y-%m-%d')}
                        )
dates = in_oa_df['日期'].unique().tolist()
#in_oa_df = in_oa_df.sort_values(by='姓名', ascending=True)

in_hr_df = pd.read_excel(sys.argv[2], skiprows=1, header=None, usecols=[3,4,5], 
                         names=['name', 'dt_from', 'dt_to'], 
                         converters={'name': lambda n : re.search(r'^([^A-Za-z ]+)', n).group(1) if re.search(r'^([^A-Za-z ]+)', n) else n}
                        )

out_df = pd.DataFrame(index=[], columns=(['部门', '性别'] + dates))

i = 0
cur_user = None

while(i < len(in_oa_df)):
    row = in_oa_df.loc[i, :]
    if row['时间段'] == '(-)':
        i += 1
        continue

    cur_user = re.sub(r'[　 ]+', '', row['姓名']) if re.search(r'[^A-Za-z ]', row['姓名']) else row['姓名']
    if not (cur_user in out_df.index):
        print(f'Processing {cur_user}')
        out_df.loc[cur_user] = [None] * len(out_df.columns)
        out_df.loc[cur_user, '部门'] = row['组织']
        out_df.loc[cur_user, '性别'] = row['性别']

    result = 1
    result_desc = ''

    absence_records = in_hr_df.loc[lambda df: df['name'] == cur_user]
    on_duty_time = dt.datetime.fromisoformat(row['日期'] + ' ' + re.search(r'([0-9:]+)-', row['时间段']).group(1))
    off_duty_time = dt.datetime.fromisoformat(row['日期'] + ' ' + re.search(r'-([0-9:]+)', row['时间段']).group(1))

    on_duty_time, off_duty_time = calc_actual_duty_time(on_duty_time, off_duty_time, absence_records)

    if(on_duty_time == None and off_duty_time == None):
        result_desc = '休假'
    else:
        if type(row['签到时间']) is dt.time and type(on_duty_time) is dt.datetime:
            if(row['签到时间'] > on_duty_time.time()):
                result *= 2
                result_desc += f'上班迟到{time_diff_by_minute(row["签到时间"], on_duty_time.time())}分钟,'
            else:
                result_desc += '上班正常,'
        else:
            result *= 3
            result_desc += '上班缺卡,'
        if type(row['签退时间']) is dt.time and type(off_duty_time) is dt.datetime:
            if row['签退时间'] < off_duty_time.time():
                result *= 5
                result_desc += f'下班早退{time_diff_by_minute(off_duty_time.time(), row["签退时间"])}分钟,'
            else:
                result_desc += '下班正常'
        else:
            result *= 7
            result_desc += '下班缺卡'
        if result == 1:
            result_desc = '正常'
        if result % (3 * 7) == 0:
            result_desc = '旷工'
    
    out_df.loc[cur_user, row['日期']] = result_desc
    i += 1

out_df.style.applymap(mapping_color).to_excel('output.xlsx', sheet_name='考勤统计', startrow=0, startcol=0, index=True)