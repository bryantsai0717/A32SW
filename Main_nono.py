import logging
import os
import pandas as pd
from datetime import datetime,timedelta
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
    
    pd.set_option('display.max_rows',None)
    pd.set_option('display.max_columns',None)
    '''
    Day = input('輸入日期(格式yyyymmdd) :')
    order = ['Y:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_%s_D2.xls' % Day,
             'Y:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_%s_N2.xls' % Day,
             'Y:\ARM-HR-Dept\HR-Data\P2\OT\%sPRE_OVERTIME.xls' % Day,
             'Y:\ARM-HR-Dept\HR-Data\P2\HC\employee_%s.xls' % Day,
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\班別.xlsx',
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\money.xlsx',
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\排程.xlsx',
             'Y:\A32_Operation\HCP_data\INPUT\match_table.xlsx',
             'Y:\A32_Operation\FATP_CPB\ASSY\%s_ASSY.csv' % Day,
             'Y:\\A32_Operation\\FATP_CPB\\PACK\\%s_PACK.csv' % Day,
             'Y:\A32_Operation\FATP_CPB\CPB\Input\%s產量目標.xlsx' % Day,
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\MES_Transform.xlsx',
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\CPB目標.xlsx',
             'Y:\A32_Operation\HCP_data\PECC OT\FATP\%s.xlsx' % Day,
             'Y:\A32_Operation\HCP_data\在職人力\在職人力_%s.xls' % Day]

    writer = pd.ExcelWriter('C:\\Users\\Nono_Wang\\Desktop\\Python\\CPB_PY.xlsx')
    '''
    #read_excel
    '''
    File_1 = pd.read_excel(order[0],'工作表1')
    File_2 = pd.read_excel(order[1],'工作表1')
    File_3 = pd.read_excel(order[2],'工作表1')
    File_4 = pd.read_excel(order[3],'工作表1')
    File_5 = pd.read_excel(order[4],'Shift')
    File_6 = pd.read_excel(order[5],'money')
    File_7 = pd.read_excel(order[6],'工作表1')
    File_8 = pd.read_excel(order[7],'SMT',header=2)
    File_9 = pd.read_csv(order[8])
    File_10 = pd.read_csv(order[9],encoding='big5')
    File_11 = pd.read_excel(order[10],'工作表1',header=1)
    File_12 = pd.read_excel(order[11],'Transform')
    File_13 = pd.read_excel(order[12],'Goal',header=1)
    File_14 = pd.read_excel(order[13],'Sheet1')
    File_15 = pd.read_excel(order[14],'工作表1')
    '''

    #set_index
    '''
    File_1_match = pd.DataFrame(File_1.set_index('員工工號').to_dict())
    File_2_match = pd.DataFrame(File_2.set_index('員工工號').to_dict())
    File_3_match = pd.DataFrame(File_3.set_index('工號').to_dict())
    File_4_match = pd.DataFrame(File_4.set_index('工號').to_dict())
    File_5_match = pd.DataFrame(File_5.set_index('班别').to_dict())
    File_6_match = pd.DataFrame(File_6.set_index('年資(月)').to_dict())
    #File_7_match = pd.DataFrame(File_7.set_index('').to_dict())
    File_8_match = pd.DataFrame(File_8.set_index('Unnamed: 3').to_dict())
    File_9_match = pd.DataFrame(File_9.set_index('線別').to_dict())
    File_10_match = pd.DataFrame(File_10.set_index('線別').to_dict())
    File_11_match = pd.DataFrame(File_11.set_index('线别&班别').to_dict())
    File_12_match = pd.DataFrame(File_12.set_index('MES代碼').to_dict())
    File_13_match = pd.DataFrame(File_13.set_index('線別&班別').to_dict())
    File_14_match = pd.DataFrame(File_14.set_index('工號(必填)').to_dict())
    File_15_match = pd.DataFrame(File_15.set_index('工號').to_dict())
    '''
    #Merge
    '''
    #File_5_match = pd.DataFrame(File_5.set_index('班别').to_dict())
    File_1 = pd.read_excel(order[0],'工作表1')
    File_2 = pd.read_excel(order[1],'工作表1')
    File_5 = pd.read_excel(order[4],'Shift')

    File_1.to_excel(writer,'DayShift',index=0,columns=['員工工號','部門代號','年資日期','線別卡 實際刷卡線別','線別卡 系統維護線別','班別','考勤卡上班時間'])
    File_2.to_excel(writer,'NightShift',index=0,columns=['員工工號','部門代號','年資日期','線別卡 實際刷卡線別','線別卡 系統維護線別','班別','考勤卡上班時間'])

    RD_Writer = pd.read_excel(writer,sheet_name='DayShift')
    RD_Writer.rename(columns={RD_Writer.columns[5]:'班别'},inplace=True)

    right = pd.DataFrame(File_5,columns=['班别','出勤開始時間'])
    left = pd.DataFrame(RD_Writer)

    left['班别'] = left['班别'].map(lambda x: str(x)[0:3])

    print(left['班别'])

    classes = pd.merge(left,right,on='班别',how='inner',indicator=True)
    classes.to_excel('Merge_FT.xlsx',index=0)
    
    mix = pd.merge(pd.DataFrame(RD_Writer),pd.DataFrame(File_3,columns=['工號','姓名']),on='工號')
    mix.to_excel('Merge_FT.xlsx',index=0)
    
    writer.save()

    '''
    #time,datetime

    '''
    now = int(time.time())
    print(now)

    timeArray = time.localtime(now)
    otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)
    print(otherStyleTime)

    nowdate = datetime.datetime.now()
    print(nowdate.strftime("%Y-%m-%d %H:%M:%S"))

    threeDayAgo = (datetime.datetime.now() - datetime.timedelta(days = 2,hours = 3))
    print (threeDayAgo)
    #timeStamp = int(time.mktime(threeDayAgo.timetuple()))
    #otherStyleTime = threeDayAgo.strftime("%Y-%m-%d %H:%M:%S")
    print (threeDayAgo.strftime("%Y-%m-%d %H:%M:%S"))
    
    file_1 = pd.read_excel('123444.xls')
    a = pd.DataFrame(file_1,columns=['線別刷卡','考勤卡上班時間','考勤卡下班時間'])
    print (a.head(5))


    file_1['考勤卡下班時間'] = list(map(lambda x: x.days, pd.to_datetime('today') - pd.to_datetime(file_1['考勤卡上班時間'],format='%Y%m%d %H:%M:%S')))
    print(file_1['考勤卡下班時間'])
    '''

    #Test

    '''
    Text_File = pd.DataFrame(pd.read_excel(pd.ExcelWriter('Merge_FT.xlsx'),'Sheet1'))
    Text_File['_merge'] = list(map(lambda x: x.seconds, pd.to_datetime('today') - pd.to_datetime(Text_File['考勤卡上班時間'],format='%Y%m%d %H:%M:%S')))
    Text_File['_merge'] = Text_File['_merge'] / 60
    print(Text_File.columns)
    print(Text_File['_merge'])
    '''
    
    #Main
    #writer = pd.ExcelWriter('C:\\Users\\Nono_Wang\\Desktop\\Python\\CPB_PY.xlsx')
    
    '''
    Day = input('輸入日期(格式yyyymmdd) :')
    order = ['Y:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_%s_D2.xls' % Day,
             'Y:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_%s_N2.xls' % Day,
             'Y:\ARM-HR-Dept\HR-Data\P2\OT\%sPRE_OVERTIME.xls' % Day,
             'Y:\ARM-HR-Dept\HR-Data\P2\HC\employee_%s.xls' % Day,
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\班別.xlsx',
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\money.xlsx',
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\排程.xlsx',
             'Y:\A32_Operation\HCP_data\INPUT\match_table.xlsx',
             'Y:\A32_Operation\FATP_CPB\ASSY\%s_ASSY.csv' % Day,
             'Y:\\A32_Operation\\FATP_CPB\\PACK\\%s_PACK.csv' % Day,
             'Y:\A32_Operation\FATP_CPB\CPB\Input\%s產量目標.xlsx' % Day,
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\MES_Transform.xlsx',
             'Y:\A32_Operation\FATP_CPB\CPB\PECC\CPB目標.xlsx',
             'Y:\A32_Operation\HCP_data\PECC OT\FATP\%s.xlsx' % Day,
             'Y:\A32_Operation\HCP_data\在職人力\在職人力_%s.xls' % Day]
    File_1 = pd.read_excel(order[0],'工作表1')
    File_2 = pd.read_excel(order[1],'工作表1')
    File_3 = pd.read_excel(order[2],'工作表1')
    File_4 = pd.read_excel(order[3],'工作表1')
    File_5 = pd.read_excel(order[4],'Shift')
    File_6 = pd.read_excel(order[5],'money')
    File_7 = pd.read_excel(order[6],'工作表1')
    File_8 = pd.read_excel(order[7],'SMT',header=2)
    File_9 = pd.read_csv(order[8])
    File_10 = pd.read_csv(order[9],encoding='big5')
    File_11 = pd.read_excel(order[10],'工作表1',header=1)
    File_12 = pd.read_excel(order[11],'Transform')
    File_13 = pd.read_excel(order[12],'Goal',header=1)
    File_14 = pd.read_excel(order[13],'Sheet1')
    File_15 = pd.read_excel(order[14],'工作表1')
    '''
    
    
    '''
    writer = pd.ExcelWriter('C:\\Users\\Nono_Wang\\Desktop\\Python\\CPB_PY.xlsx')

    File_1 = pd.DataFrame(pd.read_excel('Y:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_20180911_D2.xls'))
    File_2 = pd.DataFrame(pd.read_excel('Y:\A32_Operation\FATP_CPB\CPB\PECC\班別.xlsx'))
    File_1.to_excel(writer,'DayShift',index=0,columns=['員工工號','部門代號','年資日期','線別卡 實際刷卡線別','線別卡 系統維護線別','班別','考勤卡上班時間'])
    File_3 = pd.DataFrame(pd.read_excel('Y:\ARM-HR-Dept\HR-Data\P2\OT\\20180911PRE_OVERTIME.xls'))
    
    RD_writer = pd.read_excel(writer,'CPB')
    RD_writer.rename(columns={RD_writer.columns[0]:'工號'},inplace=True)
    RD_writer.rename(columns={RD_writer.columns[2]:'年資'},inplace=True)
    RD_writer.rename(columns={RD_writer.columns[5]:'班别'},inplace=True)

    Main_sheet = pd.DataFrame(RD_writer)
    
    Class_right = File_2,columns=['班别','出勤開始時間','出勤結束時間','休息開始時間','休息結束時間']
    Main_sheet['班别'] = Main_sheet['班别'].map(lambda x: str(x)[0:3])
    pd.merge(Main_sheet,Class_right,on='班别',how='inner',indicator=True)#班別上下班時間

    PRE_right = File_3,columns=['工號','上班時間','下班時間']
    pd.merge(Main_sheet,PRE_right,on='工號',how='inner',indicator=True)#實際出勤時間
    
    RD_writer = RD_writer['考勤卡上班時間'].dropna()#不顯示NaN行
    RD_writer['考勤卡上班時間'] = ceil_dt(pd.to_datetime(RD_writer['考勤卡上班時間']), timedelta(minutes=-30))#格式轉換以30制的上班時間
    RD_writer['年資'] = (list(map(lambda x: x.days, pd.to_datetime('today') - pd.to_datetime(RD_writer['年資日期'],format='%Y%m%d')))) / 30 #計算年資
    #轉換以30分鐘制的上下班時間
    RD_writer['上班時間'] = ceil_dt(pd.to_datetime(RD_writer['上班時間']), timedelta(minutes=-30))
    RD_writer['下班時間'] = ceil_dt(pd.to_datetime(RD_writer['下班時間']), timedelta(minutes=-30))
    print (RD_writer)
    
    a = pd.DataFrame(pd.read_excel('CPB_PY.xlsx'))
    a['班制內給薪時間'] = list(map(lambda x: x.seconds, pd.to_datetime(a['薪資結束時間'],format='%m/%d/%Y %I:%M:%S %p') - pd.to_datetime(a['薪資開始時間'],format='%m/%d/%Y %I:%M:%S %p')))
    a['班制內給薪時間'] = a['班制內給薪時間']/60
    print (a['班制內給薪時間'])
    '''
    a = pd.DataFrame(pd.read_excel('Y:\ARM-HR-Dept\HR-Data\P2\LineRecord\LineRecord_20181015_N2.xls','工作表1'))
    b = pd.DataFrame(pd.read_excel('Y:\ARM-HR-Dept\HR-Data\P2\OT\\20181015PRE_OVERTIME.xls','工作表1'))
    a = a.rename(columns={'員工工號':'工號'})
    b = b[['工號','上班時間','下班時間','加班時數估算']]
    main_a = pd.merge(a,b,on='工號',how='inner',indicator=True)
    print (pd.to_datetime(main_a['下班時間'].dropna()), timedelta(minutes=-30))
    print (ceil_dt(pd.to_datetime(main_a['下班時間'].dropna()), pd.Timedelta(minutes=-30)))
    #main_a['上班時間'] = ceil_dt(pd.to_datetime(main_a['上班時間']), timedelta(minutes=-30))
    #main_a['下班時間'] = ceil_dt(pd.to_datetime(main_a['下班時間']), timedelta(minutes=-30))

    return True

def ceil_dt(dt, delta):#30分制
    return datetime.min + math.ceil((dt - datetime.min) / delta) * delta

    

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
    
    tStart = datetime.now()#計時開始

    logger.info('start time:%s', tStart)

    if Main():
        logger.info('main ok')
    else:
        logger.info('main fail')

    tEnd  = datetime.now()#計時結束

    logger.info('end time:%s', tEnd)
          
    logger.info('system exit cost:%s' % (tEnd - tStart))
    logger.info('-' * 50)
        
    console_handler.close()
    logger.removeHandler(console_handler)
        
    file_handler.close()
    logger.removeHandler(file_handler)    
    
    sys.exit()
