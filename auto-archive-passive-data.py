# Data : 2023/8/22 17:37
# Author: Shawn Shi
# Right Reserved By COROS
import sys
from FormatingData_Fast import formatting_data
import tkinter as tk
import os
from tkinter import filedialog
import glob
import shutil
import pandas as pd
import xlwings as xw
import time


# import openpyxl

def copy_data(files, target_file):
    # delete sheet without data
    file_list = []
    for file in files:
        if '$' not in file:
            file_list.append(file.replace(os.path.dirname(file) + '\\', ''))
    if ('BT.xlsx' not in file_list) and ('CP.xlsx' in file_list) and ('LP.xlsx' not in file_list):
        print("      数据不完整，请补充数据！")
        return 0
    if ('BT.xlsx' not in file_list) and ('L1.xlsx' in file_list) and ('C1.xlsx' not in file_list):
        print("      数据不完整，请补充数据！")
        return 0
    if ('BT.xlsx' not in file_list) and ('L5.xlsx' in file_list) and ('C5.xlsx' not in file_list):
        print("      数据不完整，请补充数据！")
    if not (('BT.xlsx' in file_list) or ('CP.xlsx' in file_list) or ('C1.xlsx' in file_list) or (
            'C5.xlsx' in file_list)):
        print("      数据不完整，请补充数据！")
        return 0

    for file in files:
        file_name = file.replace(os.path.dirname(file) + '\\', '')
        # copy bluetooth data
        if 'BT.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 30)))
            # frequency.to_excel(target_file, 'BT-FS',columns=[""])
            df = pd.read_excel(file, header=3, usecols='B:AE')
            BT_efficiency = df.loc[list(i for i in range(0, 31))]
            BT_efficiency.columns = columns
            BT_efficiency.loc[:, '24'] = -BT_efficiency.loc[:, '24']
            BT_gain_2400 = df.loc[list(i for i in range(579, 592))]
            BT_gain_2400.columns = columns
            BT_gain_2440 = df.loc[list(i for i in range(783, 796))]
            BT_gain_2440.columns = columns
            BT_gain_2480 = df.loc[list(i for i in range(987, 1000))]
            BT_gain_2480.columns = columns
            # BT_efficiency.to_excel(target_file, 'BT-FS')
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                BT_efficiency.to_excel(writer, sheet_name='BT-FS', columns=list(str(i) for i in range(0, 30)),
                                       index=False, header=False, startrow=2, startcol=0)
                BT_gain_2400.to_excel(writer, sheet_name='BT-FS', columns=list(str(i) for i in range(1, 13)),
                                      index=False, header=False, startrow=38, startcol=1)
                BT_gain_2440.to_excel(writer, sheet_name='BT-FS', columns=list(str(i) for i in range(1, 13)),
                                      index=False, header=False, startrow=56, startcol=1)
                BT_gain_2480.to_excel(writer, sheet_name='BT-FS', columns=list(str(i) for i in range(1, 13)),
                                      index=False, header=False, startrow=74, startcol=1)
                time.sleep(0.5)
                # os.remove(file)
        elif 'LP.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 29)))
            df = pd.read_excel(file, header=3, usecols='B:AD')
            GPS_l1_effi = df.loc[list(i for i in range(50, 71))]
            GPS_l1_effi.columns = columns
            GPS_l1_effi.loc[:, '24'] = -GPS_l1_effi.loc[:, '24']
            GPS_l5_effi = df.loc[list(i for i in range(10, 31))]
            GPS_l5_effi.columns = columns
            GPS_l5_effi.loc[:, '24'] = -GPS_l5_effi.loc[:, '24']
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                GPS_l1_effi.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
                GPS_l5_effi.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
            time.sleep(0.5)
            # os.remove(file)
        elif 'CP.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 26)))
            df = pd.read_excel(file, header=3, usecols='B:AA')
            GPS_gain_1160r = df.loc[list(i for i in range(891, 905))]
            GPS_gain_1160r.columns = columns
            GPS_gain_1160l = df.loc[list(i for i in range(908, 922))]
            GPS_gain_1160l.columns = columns
            time.sleep(0.2)
            GPS_gain_1180r = df.loc[list(i for i in range(993, 1007))]
            GPS_gain_1180r.columns = columns
            GPS_gain_1180l = df.loc[list(i for i in range(1010, 1024))]
            GPS_gain_1180l.columns = columns
            GPS_gain_1190r = df.loc[list(i for i in range(1044, 1058))]
            GPS_gain_1190r.columns = columns
            GPS_gain_1190l = df.loc[list(i for i in range(1061, 1075))]
            GPS_gain_1190l.columns = columns
            GPS_gain_1560r = df.loc[list(i for i in range(2931, 2945))]
            GPS_gain_1560r.columns = columns
            GPS_gain_1560l = df.loc[list(i for i in range(2948, 2962))]
            GPS_gain_1560l.columns = columns
            GPS_gain_1580r = df.loc[list(i for i in range(3033, 3047))]
            GPS_gain_1580r.columns = columns
            GPS_gain_1580l = df.loc[list(i for i in range(3050, 3064))]
            GPS_gain_1580l.columns = columns
            GPS_gain_1610r = df.loc[list(i for i in range(3186, 3200))]
            GPS_gain_1610r.columns = columns
            GPS_gain_1610l = df.loc[list(i for i in range(3203, 3217))]
            GPS_gain_1610l.columns = columns
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                # load l5 data to excel
                GPS_gain_1160r.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1160l.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1180r.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1180l.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1190r.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1190l.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
                # load l1 data to excel
                GPS_gain_1560r.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1560l.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1580r.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1580l.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1610r.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1610l.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
            time.sleep(0.5)
            # os.remove(file)
        elif 'L5.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 29)))
            df = pd.read_excel(file, header=3, usecols='B:AD')
            GPS_l5_effi = df.loc[list(i for i in range(0, 21))]
            GPS_l5_effi.columns = columns
            GPS_l5_effi.loc[:, '24'] = -GPS_l5_effi.loc[:, '24']
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                GPS_l5_effi.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
            time.sleep(0.5)
            # os.remove(file)
        elif 'C5.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 26)))
            df = pd.read_excel(file, header=3, usecols='B:AA')
            GPS_gain_1160r = df.loc[list(i for i in range(331, 345))]
            GPS_gain_1160r.columns = columns
            GPS_gain_1160l = df.loc[list(i for i in range(348, 362))]
            GPS_gain_1160l.columns = columns
            GPS_gain_1180r = df.loc[list(i for i in range(433, 447))]
            GPS_gain_1180r.columns = columns
            GPS_gain_1180l = df.loc[list(i for i in range(450, 464))]
            GPS_gain_1180l.columns = columns
            GPS_gain_1190r = df.loc[list(i for i in range(484, 498))]
            GPS_gain_1190r.columns = columns
            GPS_gain_1190l = df.loc[list(i for i in range(501, 515))]
            GPS_gain_1190l.columns = columns
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                # load l5 data to excel
                GPS_gain_1160r.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1160l.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1180r.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1180l.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1190r.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1190l.to_excel(writer, sheet_name='L5-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
            time.sleep(0.5)
            # os.remove(file)
        elif 'L1.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 29)))
            df = pd.read_excel(file, header=3, usecols='B:AD')
            GPS_l1_effi = df.loc[list(i for i in range(0, 21))]
            GPS_l1_effi.columns = columns
            GPS_l1_effi.loc[:, '24'] = -GPS_l1_effi.loc[:, '24']
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                GPS_l1_effi.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
            time.sleep(0.5)
            # os.remove(file)
        elif 'C1.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 26)))
            df = pd.read_excel(file, header=3, usecols='B:AA')
            GPS_gain_1560r = df.loc[list(i for i in range(331, 345))]
            GPS_gain_1560r.columns = columns
            GPS_gain_1560l = df.loc[list(i for i in range(348, 362))]
            GPS_gain_1560l.columns = columns
            GPS_gain_1580r = df.loc[list(i for i in range(433, 447))]
            GPS_gain_1580r.columns = columns
            GPS_gain_1580l = df.loc[list(i for i in range(450, 464))]
            GPS_gain_1580l.columns = columns
            GPS_gain_1610r = df.loc[list(i for i in range(586, 600))]
            GPS_gain_1610r.columns = columns
            GPS_gain_1610l = df.loc[list(i for i in range(603, 617))]
            GPS_gain_1610l.columns = columns
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                # load l1 data to excel
                GPS_gain_1560r.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1560l.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1580r.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1580l.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1610r.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1610l.to_excel(writer, sheet_name='L1-FS', columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
            time.sleep(0.5)
            # os.remove(file)
    # delete sheet without data
    wb = xw.Book(target_file)
    if 'BT.xlsx' not in file_list:
        for sheet in wb.sheets:
            if 'BT-' in sheet.name:
                sheet.delete()
    if ('C1.xlsx' not in file_list) and ('C5.xlsx' not in file_list) and ('CP.xlsx' not in file_list):
        for sheet in wb.sheets:
            if 'L1-' in sheet.name or 'L5-' in sheet.name:
                sheet.delete()
    if ('C1.xlsx' in file_list) and ('C5.xlsx' not in file_list):
        for sheet in wb.sheets:
            if 'L5-' in sheet.name:
                sheet.delete()
    if ('C1.xlsx' not in file_list) and ('C5.xlsx' in file_list):
        for sheet in wb.sheets:
            if 'L1-' in sheet.name:
                sheet.delete()
    # rename sheets
    for sheet in wb.sheets:
        if '-' in sheet.name:
            sheet.name = sheet.name.split('-')[0] + '-' + target_file.split('/')[-1].split('.')[0]
    # return wb
    time.sleep(1.5)
    wb.save()
    for file in files:
        file_name = file.replace(os.path.dirname(file) + '\\', '')
        if file_name in ['BT.xlsx', 'LP.xlsx', 'CP.xlsx', 'L1.xlsx', 'C1.xlsx', 'L5.xlsx', 'C5.xlsx']:
            os.remove(file)
    return wb
    # wb.close()
    # wb.app.kill


def append_data(files, target_file):
    target_file_name = target_file.replace(os.path.dirname(target_file) + '/', '').split('.')[0]
    time.sleep(0.5)
    try:
        twb = xw.Book('//nas.local/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.2.xlsx')
    except:
        twb = xw.Book('//10.0.0.5/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.2.xlsx')
    # for tsheet in twb.sheets:
    #     # if wbs contains only one data set
    #     if 'L1-' in tsheet.name:
    #         # tsheet.copy(after=wb.sheets['L1-FS'], name=ssheet.name)
    #         time.sleep(1.0)
        # elif 'L5-' in ssheet.name:
        #     ssheet.copy(after=wb.sheets['L5-FS'], name=ssheet.name)
        #     time.sleep(1.0)
        # elif 'BT-' in ssheet.name:
        #     ssheet.copy(after=wb.sheets['BT-FS'], name=ssheet.name)
        #     time.sleep(1.0)
    file_list = []
    l1_sheet_name = ''
    l5_sheet_name = ''
    for file in files:
        if '$' not in file:
            file_list.append(file.replace(os.path.dirname(file) + '/', ''))
    if ('BT.xlsx' not in file_list) and ('CP.xlsx' in file_list) and ('LP.xlsx' not in file_list):
        print("      数据不完整，请补充数据！")
        return 0
    if ('BT.xlsx' not in file_list) and ('L1.xlsx' in file_list) and ('C1.xlsx' not in file_list):
        print("      数据不完整，请补充数据！")
        return 0
    if ('BT.xlsx' not in file_list) and ('L5.xlsx' in file_list) and ('C5.xlsx' not in file_list):
        print("      数据不完整，请补充数据！")
    if not (('BT.xlsx' in file_list) or ('CP.xlsx' in file_list) or ('C1.xlsx' in file_list) or (
            'C5.xlsx' in file_list)):
        print("      数据不完整，请补充数据！")
        return 0

    for file in files:
        wb = xw.Book(target_file)
        file_name = file.replace(os.path.dirname(file) + '/', '')
        # source_file_name = source_file.replace(os.path.dirname(source_file) + '/', '').split('.')[0]
        # copy bluetooth data
        if 'BT.xlsx' == file_name:
            columns = pd.Index(list(str(i) for i in range(0, 30)))
            # frequency.to_excel(target_file, 'BT-FS',columns=[""])
            df = pd.read_excel(file, header=3, usecols='B:AE', engine='openpyxl')
            BT_efficiency = df.loc[list(i for i in range(0, 31))]
            BT_efficiency.columns = columns
            BT_efficiency.loc[:, '24'] = -BT_efficiency.loc[:, '24']
            BT_gain_2400 = df.loc[list(i for i in range(579, 592))]
            BT_gain_2400.columns = columns
            BT_gain_2440 = df.loc[list(i for i in range(783, 796))]
            BT_gain_2440.columns = columns
            BT_gain_2480 = df.loc[list(i for i in range(987, 1000))]
            BT_gain_2480.columns = columns
            # BT_efficiency.to_excel(target_file, 'BT-FS')
            bt_sheet_name = 'BT-'+target_file_name
            twb.sheets['BT-FS'].copy(after=wb.sheets['References'], name=bt_sheet_name)
            wb.save()
            time.sleep(0.5)
            wb.close()
            time.sleep(0.5)
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                BT_efficiency.to_excel(writer, sheet_name=bt_sheet_name, columns=list(str(i) for i in range(0, 30)),
                                       index=False, header=False, startrow=2, startcol=0)
                BT_gain_2400.to_excel(writer, sheet_name=bt_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                      index=False, header=False, startrow=38, startcol=1)
                BT_gain_2440.to_excel(writer, sheet_name=bt_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                      index=False, header=False, startrow=56, startcol=1)
                BT_gain_2480.to_excel(writer, sheet_name=bt_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                      index=False, header=False, startrow=74, startcol=1)
                time.sleep(0.5)
                # os.remove(file)
        elif 'LP.xlsx' == file_name:
            if l1_sheet_name == '':
                l1_sheet_name = 'L1-' + target_file_name
                twb.sheets['L1-FS'].copy(after=wb.sheets['References'], name=l1_sheet_name)
                l5_sheet_name = 'L5-' + target_file_name
                twb.sheets['L5-FS'].copy(after=wb.sheets['References'], name=l5_sheet_name)
            columns = pd.Index(list(str(i) for i in range(0, 29)))
            df = pd.read_excel(file, header=3, usecols='B:AD', engine='openpyxl')
            GPS_l1_effi = df.loc[list(i for i in range(50, 71))]
            GPS_l1_effi.columns = columns
            GPS_l1_effi.loc[:, '24'] = -GPS_l1_effi.loc[:, '24']
            GPS_l5_effi = df.loc[list(i for i in range(10, 31))]
            GPS_l5_effi.columns = columns
            GPS_l5_effi.loc[:, '24'] = -GPS_l5_effi.loc[:, '24']
            wb.save()
            time.sleep(0.5)
            wb.close()
            time.sleep(0.5)
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                GPS_l1_effi.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
                GPS_l5_effi.to_excel(writer, sheet_name=l5_sheet_name , columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
            time.sleep(0.5)
            # os.remove(file)
        elif 'CP.xlsx' == file_name:
            if l1_sheet_name == '':
                l1_sheet_name = 'L1-' + target_file_name
                twb.sheets['L1-FS'].copy(after=wb.sheets['References'], name=l1_sheet_name)
                l5_sheet_name = 'L5-' + target_file_name
                twb.sheets['L5-FS'].copy(after=wb.sheets['References'], name=l5_sheet_name)
            columns = pd.Index(list(str(i) for i in range(0, 26)))
            df = pd.read_excel(file, header=3, usecols='B:AA', engine='openpyxl')
            GPS_gain_1160r = df.loc[list(i for i in range(891, 905))]
            GPS_gain_1160r.columns = columns
            GPS_gain_1160l = df.loc[list(i for i in range(908, 922))]
            GPS_gain_1160l.columns = columns
            time.sleep(0.2)
            GPS_gain_1180r = df.loc[list(i for i in range(993, 1007))]
            GPS_gain_1180r.columns = columns
            GPS_gain_1180l = df.loc[list(i for i in range(1010, 1024))]
            GPS_gain_1180l.columns = columns
            GPS_gain_1190r = df.loc[list(i for i in range(1044, 1058))]
            GPS_gain_1190r.columns = columns
            GPS_gain_1190l = df.loc[list(i for i in range(1061, 1075))]
            GPS_gain_1190l.columns = columns
            GPS_gain_1560r = df.loc[list(i for i in range(2931, 2945))]
            GPS_gain_1560r.columns = columns
            GPS_gain_1560l = df.loc[list(i for i in range(2948, 2962))]
            GPS_gain_1560l.columns = columns
            GPS_gain_1580r = df.loc[list(i for i in range(3033, 3047))]
            GPS_gain_1580r.columns = columns
            GPS_gain_1580l = df.loc[list(i for i in range(3050, 3064))]
            GPS_gain_1580l.columns = columns
            GPS_gain_1610r = df.loc[list(i for i in range(3186, 3200))]
            GPS_gain_1610r.columns = columns
            GPS_gain_1610l = df.loc[list(i for i in range(3203, 3217))]
            GPS_gain_1610l.columns = columns
            wb.save()
            time.sleep(0.5)
            wb.close()
            time.sleep(0.5)
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                # load l5 data to excel
                GPS_gain_1160r.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1160l.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1180r.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1180l.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1190r.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1190l.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
                # load l1 data to excel
                GPS_gain_1560r.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1560l.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1580r.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1580l.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1610r.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1610l.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
            time.sleep(0.5)
            # os.remove(file)
        elif 'L5.xlsx' == file_name:
            if l5_sheet_name == '':
                # l1_sheet_name = 'L1-' + target_file_name
                # twb.sheets['L1-FS'].copy(after=wb.sheets['Reference'], name=l1_sheet_name)
                l5_sheet_name = 'L5-' + target_file_name
                twb.sheets['L5-FS'].copy(after=wb.sheets['References'], name=l5_sheet_name)
            columns = pd.Index(list(str(i) for i in range(0, 29)))
            df = pd.read_excel(file, header=3, usecols='B:AD', engine='openpyxl')
            GPS_l5_effi = df.loc[list(i for i in range(0, 21))]
            GPS_l5_effi.columns = columns
            GPS_l5_effi.loc[:, '24'] = -GPS_l5_effi.loc[:, '24']
            wb.save()
            time.sleep(0.5)
            wb.close()
            time.sleep(0.5)
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                GPS_l5_effi.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
            time.sleep(0.5)
            # os.remove(file)
        elif 'C5.xlsx' == file_name:
            if l5_sheet_name == '':
                # l1_sheet_name = 'L1-' + target_file_name
                # twb.sheets['L1-FS'].copy(after=wb.sheets['Reference'], name=l1_sheet_name)
                l5_sheet_name = 'L5-' + target_file_name
                twb.sheets['L5-FS'].copy(after=wb.sheets['References'], name=l5_sheet_name)
            columns = pd.Index(list(str(i) for i in range(0, 26)))
            df = pd.read_excel(file, header=3, usecols='B:AA', engine='openpyxl')
            GPS_gain_1160r = df.loc[list(i for i in range(331, 345))]
            GPS_gain_1160r.columns = columns
            GPS_gain_1160l = df.loc[list(i for i in range(348, 362))]
            GPS_gain_1160l.columns = columns
            GPS_gain_1180r = df.loc[list(i for i in range(433, 447))]
            GPS_gain_1180r.columns = columns
            GPS_gain_1180l = df.loc[list(i for i in range(450, 464))]
            GPS_gain_1180l.columns = columns
            GPS_gain_1190r = df.loc[list(i for i in range(484, 498))]
            GPS_gain_1190r.columns = columns
            GPS_gain_1190l = df.loc[list(i for i in range(501, 515))]
            GPS_gain_1190l.columns = columns
            wb.save()
            time.sleep(0.5)
            wb.close()
            time.sleep(0.5)
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                # load l5 data to excel
                GPS_gain_1160r.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1160l.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1180r.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1180l.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1190r.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1190l.to_excel(writer, sheet_name=l5_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
            time.sleep(0.5)
            # os.remove(file)
        elif 'L1.xlsx' == file_name:
            if l1_sheet_name == '':
                l1_sheet_name = 'L1-' + target_file_name
                twb.sheets['L1-FS'].copy(after=wb.sheets['References'], name=l1_sheet_name)
                # l5_sheet_name = 'L5-' + target_file_name
                # twb.sheets['L5-FS'].copy(after=wb.sheets['Reference'], name=l5_sheet_name)
            columns = pd.Index(list(str(i) for i in range(0, 29)))
            df = pd.read_excel(file, header=3, usecols='B:AD', engine='openpyxl')
            GPS_l1_effi = df.loc[list(i for i in range(0, 21))]
            GPS_l1_effi.columns = columns
            GPS_l1_effi.loc[:, '24'] = -GPS_l1_effi.loc[:, '24']
            wb.save()
            time.sleep(0.5)
            wb.close()
            time.sleep(0.5)
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                GPS_l1_effi.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(0, 29)),
                                     index=False, header=False, startrow=2, startcol=0)
            time.sleep(0.5)
            # os.remove(file)
        elif 'C1.xlsx' == file_name:
            if l1_sheet_name == '':
                l1_sheet_name = 'L1-' + target_file_name
                twb.sheets['L1-FS'].copy(after=wb.sheets['References'], name=l1_sheet_name)
                # l5_sheet_name = 'L5-' + target_file_name
                # twb.sheets['L5-FS'].copy(after=wb.sheets['Reference'], name=l5_sheet_name)
            columns = pd.Index(list(str(i) for i in range(0, 26)))
            df = pd.read_excel(file, header=3, usecols='B:AA', engine='openpyxl')
            GPS_gain_1560r = df.loc[list(i for i in range(331, 345))]
            GPS_gain_1560r.columns = columns
            GPS_gain_1560l = df.loc[list(i for i in range(348, 362))]
            GPS_gain_1560l.columns = columns
            GPS_gain_1580r = df.loc[list(i for i in range(433, 447))]
            GPS_gain_1580r.columns = columns
            GPS_gain_1580l = df.loc[list(i for i in range(450, 464))]
            GPS_gain_1580l.columns = columns
            GPS_gain_1610r = df.loc[list(i for i in range(586, 600))]
            GPS_gain_1610r.columns = columns
            GPS_gain_1610l = df.loc[list(i for i in range(603, 617))]
            GPS_gain_1610l.columns = columns
            wb.save()
            time.sleep(0.5)
            wb.close()
            time.sleep(0.5)
            with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
                # load l1 data to excel
                GPS_gain_1560r.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=28, startcol=1)
                GPS_gain_1560l.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=45, startcol=1)
                GPS_gain_1580r.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=78, startcol=1)
                GPS_gain_1580l.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=95, startcol=1)
                GPS_gain_1610r.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=128, startcol=1)
                GPS_gain_1610l.to_excel(writer, sheet_name=l1_sheet_name, columns=list(str(i) for i in range(1, 13)),
                                        index=False, header=False, startrow=145, startcol=1)
            time.sleep(0.5)
            # os.remove(file)
    # return wb
    twb.close()
    time.sleep(1.0)
    wb = xw.Book(target_file)
    wb.save()
    for file in files:
        file_name = file.replace(os.path.dirname(file) + '/', '')
        if file_name in ['BT.xlsx', 'LP.xlsx', 'CP.xlsx', 'L1.xlsx', 'C1.xlsx', 'L5.xlsx', 'C5.xlsx']:
            os.remove(file)
    return wb
    # wb.close()
    # wb.app.kill()


def merge_files(files, target_file):
    # Determine whether source files intersect with each other.
    # if yes, delete the sheets
    wb = xw.Book(target_file)
    time.sleep(1.0)
    for file in files:
        # source file
        wbs = xw.Book(file)
        time.sleep(1.5)
        for ssheet in wbs.sheets:
            # if wbs contains only one data set
            if 'L1-' in ssheet.name:
                ssheet.copy(after=wb.sheets['L1-FS'], name=ssheet.name)
                time.sleep(1.0)
            elif 'L5-' in ssheet.name:
                ssheet.copy(after=wb.sheets['L5-FS'], name=ssheet.name)
                time.sleep(1.0)
            elif 'BT-' in ssheet.name:
                ssheet.copy(after=wb.sheets['BT-FS'], name=ssheet.name)
                time.sleep(1.0)
        wbs.close()
        time.sleep(1.5)
    for tsheet in wb.sheets:
        if 'FS' in tsheet.name:
            tsheet.delete()
    # os.wait(1)
    wb.save()
    return wb


def rename_sheet(file):
    wb = xw.Book(file)
    time.sleep(0.5)
    sheets = wb.sheets
    sheet_name_list = [sheet.name for sheet in sheets]
    # In case that one more than sheet have are started with 'BT' or 'L1' or 'L5', stop this function
    for element1 in sheet_name_list:
        for element2 in sheet_name_list:
            if element1[:3] == element2[:3] and element1 != element2:
                print('     excel文件包含至少两个DUT数据，无法重命名')
                return 0
    for sheet in sheets:
        if '-' in sheet.name:
            sheet.name = sheet.name.split('-')[0] + '-' + file.split('/')[-1].split('.')[0]
    wb.save()
    return wb


if __name__ == '__main__':
    selection = '0'
    print('><<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>')
    print('     Antenna Passive Test Data Automation Archiving Tool V5.0   ')
    print('**************All rights are reserved by COROS******************')
    print("----------------------使用指南-----------------------------------")
    print("---------------------------------------------------------------")
    print("         1. 所有测试请选择标准模板\n")
    print("         2. 蓝牙测试数据请导出为BT.xlsx文件\n")
    print("         3. 双频GPS线极化测试数据请导出为LP.xlsx\n")
    print("         4. 双频GPS圆极化测试数据请导出为CP.xlsx\n")
    print("         5. L1线极化测试数据请导出为L1.xlsx\n")
    print("         6. L1圆极化测试数据请导出为C1.xlsx\n")
    print("         7. L5线极化测试数据请导出为L5.xlsx\n")
    print("         8. L5圆极化测试数据请导出为C5.xlsx\n")
    try:
        while selection != '6':
            print("---------------------------------------------------------------")
            print("---------------------请选择一个功能-------------------------------")
            print("---------------------------------------------------------------")
            print("         1. 将导出的测试数据格式化写入到一个xlsx文件中\n")
            print("         2. 合并多个xlsx文件\n")
            print("         3. 追加数据到xlsx文件中\n")
            print("         4. 数据评分\n")
            print("         5. 重命名sheet以及文件名\n")
            print("         6. 退出程序")
            print('===============================================================')
            selection = input("请输入你的选择：")
            root = tk.Tk()
            root.withdraw()
            if selection == '1':
                print("************************************************************\n")
                print("============1. 将导出的测试数据格式化写入到一个xlsx文件中============\n")
                print('*******请选择一个源文件目录(内有包含GPS或者BT数据的xlsx文件)*******\n')
                source_file_path = filedialog.askdirectory(title='打开测试数据目录')
                if source_file_path == '':
                    print('未选择任何目录，请选择一个目录，如取消选择，将返回到主目录')
                    source_file_path = filedialog.askdirectory(title='打开测试数据目录')
                if source_file_path == '':
                    continue
                print(f"源文件目录为 {source_file_path}\n")
                excel_name = input("========请输入xlsx名称(35字符以内)========\n")
                if excel_name == '' or len(excel_name) > 35:
                    print('名称为空，请再次输入xlsx名称，如输入不合规，将返回到主目录\n')
                    excel_name = input("========请输入xlsx名称========\n")
                if excel_name == '' or len(excel_name) > 35:
                    continue
                start_time = time.perf_counter()
                print('     数据归档进行中...\n')
                target_file = f'{source_file_path}/{excel_name}.xlsx'
                # if the file has existed, delete the file
                if os.path.exists(target_file):
                    print('目标文件已存在，目标文件将会被删除并按照测试数据重新创建，请选择是否继续y/n')
                    if_delete = input()
                    if if_delete == 'n':
                        continue
                    os.remove(target_file)
                # os.popen(f'//nas.local/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.0.xlsx'
                # f' {source_file_path}/{sheet_name}.xlsx')
                try:
                    shutil.copyfile(
                        '//nas.local/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.2.xlsx',
                        target_file)
                except:
                    shutil.copyfile(
                        '//10.0.0.5/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.2.xlsx',
                        target_file)
                time.sleep(0.5)
                files = glob.glob(source_file_path + r"/*.xlsx")
                files.remove(f'{source_file_path}\\{excel_name}.xlsx')
                wb = copy_data(files, target_file)
                if wb == 0:
                    continue
                print('     数据归档完成\n')
                print('     数据评分中...\n')
                formatting_data(target_file, wb)
                print('     数据评分完成\n')
                print('总计用时: %s 秒\n' % (round((time.perf_counter() - start_time), 2)))
            elif selection == '2':
                print("***********************************************************\n")
                print("================2. 合并多个xlsx文件 =======================\n")
                print('********************请选择一个或多个源文件********************\n')
                # source_file_path = filedialog.askdirectory(title='打开源文件目录')
                source_files = filedialog.askopenfilenames(title='选择源文件（可多选）')
                if len(source_files) == 0:
                    print('未选择任何文件，请再次选择源文件，如取消选择，将返回到主目录\n')
                    source_files = filedialog.askopenfilenames(title='选择源文件（可多选）')
                if len(source_files) == 0:
                    continue
                source_file_path = os.path.dirname(source_files[0])
                print(f"源文件目录为 {source_file_path}\n")
                # path = r"\\nas.local\DATA\Wireless\Library\Components\for test"
                excel_name = input("========请输入汇总后的文件名称========\n")
                if excel_name == '':
                    print('名称为空，请再次输入xlsx名称，如输入不合规，将返回到主目录\n')
                    excel_name = input("========请输入xlsx名称========\n")
                if excel_name == '':
                    continue
                # estimate whether a file in the directory has a same name as excel_name
                target_file = f'{source_file_path}/{excel_name}.xlsx'
                # files = glob.glob(source_file_path + r"/*.xlsx")
                start_time = time.perf_counter()
                # target_file_name = target_file.replace(os.path.dirname(target_file) + '/', '').split('.')[0]
                renamed_source_file_name = excel_name + '_old'
                if target_file in source_files:
                    print('excel_name与现有文件重名，现有文件将会被重命名为*_old.xlsx\n')
                    # rename the source with the same name as the target file
                    os.rename(target_file, target_file.replace(excel_name, renamed_source_file_name))
                    source_files = [source_file.replace(
                        target_file, target_file.replace(excel_name, renamed_source_file_name))
                        for source_file in source_files]
                print('     文件合并中...\n')
                # os.popen(f'//nas.local/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.0.xlsx'
                # f' {source_file_path}/{sheet_name}.xlsx')
                try:
                    shutil.copyfile(
                        '//nas.local/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.2.xlsx',
                        target_file)
                except:
                    shutil.copyfile(
                        '//10.0.0.5/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.2.xlsx',
                        target_file)
                time.sleep(0.5)
                # dest_path = r"\\nas.local\DATA\Wireless\Library\Components\for test\Shunt\\"
                # files = glob.glob(source_file_path + r"/*.xlsx")
                # files.remove(f'{source_file_path}\\{excel_name}.xlsx')
                wb = merge_files(source_files, target_file)
                print('     文件合并完成\n')
                print('     数据评分中...\n')
                formatting_data(target_file, wb)
                print('     数据评分完成\n')
                print('总计用时: %s 秒\n' % (round((time.perf_counter() - start_time), 2)))
            elif selection == '3':
                target_file = ''
                print("***********************************************************\n")
                print("================3. 追加数据到xlsx文件中 =======================\n")
                print('********************请选择目标文件以及测试数据文件********************\n')
                # source_file_path = filedialog.askdirectory(title='打开源文件目录')
                source_files = filedialog.askopenfilenames(title='选择文件（至少选择一个目标文件和一个测试数据文件）')
                if len(source_files) == 0:
                    print('未选择任何文件，请再次选择源文件，如取消选择，将返回到主目录\n')
                    source_files = filedialog.askopenfilenames(title='选择文件（至少选择一个目标文件和一个测试数据文件）')
                if len(source_files) == 0:
                    continue
                source_file_path = os.path.dirname(source_files[0])
                print(f"源文件目录为 {source_file_path}\n")
                files = []
                for file in source_files:
                    file_name = file.replace(os.path.dirname(file)+'/', '')
                    if file_name in ['BT.xlsx', 'LP.xlsx', 'CP.xlsx', 'L1.xlsx', 'C1.xlsx', 'L5.xlsx', 'C5.xlsx']:
                        files.append(file)
                    else:
                        target_file = file
                if target_file == '':
                    print('未选择目标文件，返回主目录\n')
                    continue
                # files = glob.glob(source_file_path + r"/*.xlsx")
                start_time = time.perf_counter()

                print('     追加数据中...\n')
                # os.popen(f'//nas.local/DATA/Wireless/AntennaTest/Templates/Antenna passive test templates V7.0.xlsx'
                time.sleep(0.5)
                # dest_path = r"\\nas.local\DATA\Wireless\Library\Components\for test\Shunt\\"
                # files = glob.glob(source_file_path + r"/*.xlsx")
                # files.remove(f'{source_file_path}\\{excel_name}.xlsx')
                wb = append_data(files, target_file)
                print('     文件合并完成\n')
                print('     数据评分中...\n')
                formatting_data(target_file, wb)
                print('     数据评分完成\n')
                print('总计用时: %s 秒\n' % (round((time.perf_counter() - start_time), 2)))

            elif selection == '4':
                print("****************************************************************\n")
                print("========================3. 数据评分==============================\n")
                print('***********************请选择一个文件*****************************\n')
                source_file = filedialog.askopenfilename(title='选定一个源文件')
                if source_file == '':
                    print('      未选择任何文件，请重新选择，如取消选择，将返回到主目录\n')
                    source_file = filedialog.askopenfilename(title='选定一个源文件')
                if source_file == '':
                    continue
                print(f"源文件为 {source_file}\n")
                # path = r"\\nas.local\DATA\Wireless\Library\Components\for test"
                start_time = time.perf_counter()
                wb = xw.Book(source_file)
                print('     数据评分中...\n')
                formatting_data(source_file, wb)
                print('     数据评分完成\n')
                print('总计用时: %s 秒\n' % (round((time.perf_counter() - start_time), 2)))
            elif selection == '5':
                print("****************************************************************\n")
                print("====================4. 重命名sheet以及文件名========================\n")
                print('***********************请选择一个文件*****************************\n')
                source_file = filedialog.askopenfilename(title='选定一个源文件')
                if source_file == '':
                    print('      未选择任何文件，请重新选择，如取消选择，将返回到主目录\n')
                    source_file = filedialog.askopenfilename(title='选定一个源文件')
                if source_file == '':
                    continue
                print(f"源文件为 {source_file}\n")
                source_file_name = source_file.replace(os.path.dirname(source_file) + '/', '').split('.')[0]
                # source_file_path = os.path.dirname(source_file)
                # path = r"\\nas.local\DATA\Wireless\Library\Components\for test"
                start_time = time.perf_counter()
                # wb = xw.Book(source_file)
                excel_name = input("=======请为所选文件及其sheet输入一个新名字(35字符以内)=========\n")
                if excel_name == '' or len(excel_name) > 35:
                    print('名称为空或长度35，请再次输入xlsx名称，如输入不合规，将返回到主目录')
                    excel_name = input("==========请为所选文件及其sheet输入一个新名字(35字符以内)===========\n")
                if excel_name == '' or len(excel_name) > 35:
                    continue
                os.rename(source_file, source_file.replace(source_file_name, excel_name))
                wb = rename_sheet(source_file.replace(source_file_name, excel_name))
                time.sleep(0.5)
                if wb != 0:
                    print('     重命名成功\n')
                    print('     数据评分中...\n')
                    formatting_data(source_file.replace(source_file_name, excel_name), wb)
                    print('     数据评分完成\n')
                    # print('     数据评分完成')
                    print('总计用时: %s 秒\n' % (round((time.perf_counter() - start_time), 2)))
                else:
                    print('     sheet重命名失败,仅文件被重命名\n')
            elif selection == '6':
                print("***************************************************************\n")
                sys.exit('Exit the program，have a good day')
    except Exception as e:
        print(e)
        time.sleep(5.0)
