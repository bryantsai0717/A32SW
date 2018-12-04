import logging
import os
import pandas as pd
#from datetime import datetime,timedelta
import datetime
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

def Main():
    #在shell介面中顯示所有column
    pd.set_option('display.width',None)

    a = pd.DataFrame(pd.read_excel('I:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_20181015_N2.xls','工作表1'))
    b = pd.DataFrame(pd.read_excel('I:\ARM-HR-Dept\HR-Data\P2\OT\\20181015PRE_OVERTIME.xls','工作表1'))
    a = a.rename(columns={'員工工號':'工號'})
    b = b[['工號','上班時間','下班時間','加班時數估算']]
    main_a = pd.merge(a,b,on='工號',how='inner',indicator=True)

    work_on = pd.to_datetime(main_a['上班時間'].dropna(), format='%Y/%m/%d %H:%M')
    work_off = pd.to_datetime(main_a['下班時間'].dropna(), format='%Y/%m/%d %H:%M')
    print(work_on)
    print(work_off)
    work_on_time = work_on.apply(lambda x: round_minutes(x, 'up', 30))
    work_off_time = work_off.apply(lambda x: round_minutes(x, 'down', 30))
    print(work_on_time)
    print(work_off_time)
    
##    看單個cell值    
##    work_on = pd.to_datetime(main_a.loc[1]['上班時間'], format='%Y/%m/%d %H:%M')
##    work_off = pd.to_datetime(main_a.loc[1]['下班時間'], format='%Y/%m/%d %H:%M')
##    print(work_on)
##    print(round_minutes(work_on, 'up', 30))
##    print(work_off)
##    print(round_minutes(work_off, 'down', 30))
    
    return True

def round_minutes(dt, direction, resolution):
    return dt + datetime.timedelta(minutes=((dt.minute // resolution + (1 if direction == 'up' else 0)) * resolution) - dt.minute) - datetime.timedelta(seconds=dt.second)
    

if __name__ == '__main__':
    logger = logging.getLogger()
    formatter = logging.Formatter('%(asctime)s %(levelname)-8s: %(message)s')
    
    file_handler = logging.FileHandler('log1.log')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    logger.setLevel(logging.INFO)
        
    logger.info('system start')
    
    tStart = datetime.datetime.now()#計時開始

    logger.info('start time:%s', tStart)

    if Main():
        logger.info('main ok')
    else:
        logger.info('main fail')

    tEnd  = datetime.datetime.now()#計時結束

    logger.info('end time:%s', tEnd)
          
    logger.info('system exit cost:%s' % (tEnd - tStart))
    logger.info('-' * 50)
        
    console_handler.close()
    logger.removeHandler(console_handler)
        
    file_handler.close()
    logger.removeHandler(file_handler)    
    
    sys.exit()
