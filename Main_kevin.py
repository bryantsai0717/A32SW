//
//                       _oo0oo_
//                      o8888888o
//                      88" . "88
//                      (| -_- |)
//                      0\  =  /0
//                    ___/`---'\___
//                  .' \\|     |// '.
//                 / \\|||  :  |||// \
//                / _||||| -:- |||||- \
//               |   | \\\  -  /// |   |
//               | \_|  ''\---/''  |_/ |
//               \  .-\__  '-'  ___/-. /
//             ___'. .'  /--.--\  `. .'___
//          ."" '<  `.___\_<|>_/___.' >' "".
//         | | :  `- \`.;`\ _ /`;.`/ - ` : | |
//         \  \ `_.   \_ __\ /__ _/   .-` /  /
//     =====`-.____`.___ \_____/___.-`___.-'=====
//                       `=---='
//
//
//     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//
//               佛祖保佑         永無bug
//
//***************************************************
import logging
import os
import pandas as pd
from datetime import datetime, timedelta
import time
import sys
import math
import numpy as np

def use_logging(func):
    def wraper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
        except Exception as e:
            logger.exception(e)
            result = False
        return result
    return wraper


@use_logging
def main():
    dt_today = datetime(2018, 10, 15)
    dt_nextday = dt_today + timedelta(days=1)
    
    #pd.set_option('display.max_rows', None)
    #pd.set_option('display.max_columns', None)

    a = pd.read_excel('I:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_20181015_N2.xls', '工作表1')
    b = pd.read_excel('I:\ARM-HR-Dept\HR-Data\P2\OT\\20181015PRE_OVERTIME.xls', '工作表1')
    c = pd.read_excel('I:\A32_Operation\FATP_CPB\CPB\PECC\班別.xlsx', 'Shift')

    b.rename(columns={'工號': '員工工號'}, inplace=True)
    
    res = pd.merge(a, b, on='員工工號')

    res.rename(columns={'部門代號_x':'部門代號', '部門名稱_x':'部門名稱', '年資日期_x':'年資日期', '班別_x':'班別'}, inplace=True)
    
    res = pd.merge(res, c, on='班別')
    
    res = res.dropna(subset=['上班時間', '下班時間'])
    
    res['上班時間'] = pd.to_datetime(res['上班時間'], format='%Y/%m/%d %H:%M:%S')
    res['下班時間'] = pd.to_datetime(res['下班時間'], format='%Y/%m/%d %H:%M:%S')

    res['上班時間_1'] = res['上班時間'].apply(floor_time_to_30m)
    res['下班時間_1'] = res['下班時間'].apply(ceil_time_to_30m)

    templist = ['出勤結束時間', '休息開始時間', '休息結束時間', '加班開始時間', '出勤開始時間']

    for temp in templist:
        res[temp] = res.apply(add_date_tag, axis=1, args=(temp,))

    res['計薪開始時間'] = res.apply(mony_start_time, axis=1)
    res['計薪結束時間'] = res.apply(mony_end_time, axis=1)
    res['計薪時數'] = res.apply(total_hours, axis=1)

    for i,j,k,l,m in zip(res['上班時間'], res['計薪開始時間'], res['下班時間'], res['計薪結束時間'], res['計薪時數']):print(i,'|',j,'|',k,'|',l,'|',m)

    return True

def total_hours(row):
    if row['計薪開始時間'] > row['休息結束時間']:
        return row['計薪結束時間'] - row['計薪開始時間']
    
    elif row['計薪結束時間'] < row['休息開始時間']:
        return row['計薪結束時間'] - row['計薪開始時間']

    else:
        if row['休息開始時間'] > row['計薪開始時間']:
            tat1 = row['休息開始時間'] - row['計薪開始時間']
        else:
            tat1 = timedelta(0)

        if row['計薪結束時間'] > row['休息結束時間']:
            tat2 = row['計薪結束時間'] - row['休息結束時間']
        else:
            tat2 = timedelta(0)
            
        return tat1 + tat2
    
def mony_start_time(row):
    if row['上班時間_1'] > row['出勤開始時間']:
        return row['上班時間_1']
    else:
        return row['出勤開始時間']

def mony_end_time(row):
    if row['下班時間_1'] > row['出勤結束時間']:
        return row['出勤結束時間']
    else:
        return row['下班時間_1']
    
def add_date_tag(row, temp):
    if row[temp] < row['出勤開始時間']:
        return datetime(dt_nextday.year, dt_nextday.month, dt_nextday.day, row[temp].hour, row[temp].minute, row[temp].second)
    else:
        return datetime(dt_today.year, dt_today.month, dt_today.day, row[temp].hour, row[temp].minute, row[temp].second)
        
def floor_time_to_30m(dt):
    if dt.minute > 30:
        return datetime(dt.year, dt.month, dt.day, dt.hour) + timedelta(hours=1)
    else:
        return datetime(dt.year, dt.month, dt.day, dt.hour) + timedelta(minutes=30)

def ceil_time_to_30m(dt):

    if dt.minute > 30:
        return datetime(dt.year, dt.month, dt.day, dt.hour, 30)
    else:
        return datetime(dt.year, dt.month, dt.day, dt.hour)


if __name__ == '__main__':
    logger = logging.getLogger()
    formatter = logging.Formatter('%(asctime)s %(levelname)-8s: %(message)s')

    #file_handler = logging.FileHandler('log1.log')
    #file_handler.setFormatter(formatter)
    
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    #logger.addHandler(file_handler)

    logger.setLevel(logging.INFO)

    logger.info('system start')

    tStart = datetime.now()  #計時開始

    logger.info('start time:%s' % tStart)

    if main():
        logger.info('main ok')
    else:
        logger.info('main fail')

    tEnd = datetime.now()  #計時結束

    logger.info('end time:%s' % tEnd)

    logger.info('system exit cost:%s' % (tEnd - tStart))
    logger.info('-' * 50)

    console_handler.close()
    logger.removeHandler(console_handler)

    #file_handler.close()
    #logger.removeHandler(file_handler)

    sys.exit()
