# Data : 2022/11/17 16:04
# Author: Shawn Shi
# Right Reserved By COROS
# Data : 2022/11/16 9:53
# Author: Shawn Shi
# Right Reserved By COROS

import xlwings as xw
import glob
import copy
import time
import math
import numpy as np
import pandas as pd
from itertools import product

dut_gain = {}
dut_effi = {}


def gain_value(xls_sheet, dict_name, freq):
    freq_list = list(dict_name.keys())
    theta_list = list(dict_name[freq_list[0]].keys())
    # print(xls_sheet.name)
    if freq == 'l1':
        for key in dict_name:
            # for sub_key in key:
            if key == '1560MHz':
                for i in range(0, 5):
                    dict_name[key][theta_list[i]] = xls_sheet.range('S' + (str(28 + i))).value
            if key == '1580MHz':
                for i in range(0, 5):
                    dict_name[key][theta_list[i]] = xls_sheet.range('S' + (str(78 + i))).value
            if key == '1610MHz':
                for i in range(0, 5):
                    dict_name[key][theta_list[i]] = xls_sheet.range('S' + (str(128 + i))).value
    if freq == 'l5':
        for key in dict_name:
            # for sub_key in key:
            if key == '1170MHz':
                for i in range(0, 5):
                    dict_name[key][theta_list[i]] = xls_sheet.range('S' + (str(28 + i))).value
            if key == '1190MHz':
                for i in range(0, 5):
                    dict_name[key][theta_list[i]] = xls_sheet.range('S' + (str(78 + i))).value
            if key == '1210MHz':
                for i in range(0, 5):
                    dict_name[key][theta_list[i]] = xls_sheet.range('S' + (str(128 + i))).value
    # return dict_name


def effi_value(xls_sheet, effi_name, freq):
    if freq in ['l1', 'l5']:
        effi_name[freq] = xls_sheet.range('B3:B25').value
        if freq == 'l1':
            xls_sheet.range('A9:AC14').color = (255, 255, 0)
        else:
            xls_sheet.range('A9:AC15').color = (255, 255, 0)
    else:
        effi_name[freq] = xls_sheet.range('B8:B28').value
        effi_name[freq].append(xls_sheet.range('B34').value)
        effi_name[freq].append(xls_sheet.range('B35').value)
        xls_sheet.range('A13:AC21').color = (255, 255, 0)


def gain_chara_coloring(filename, xls_sheet):
    gain_chara_color = [
        (0, 130, 0),
        (0, 180, 0),
        (145, 218, 0),
        (216, 254, 154),
        (255, 255, 0),
        (255, 200, 0),
        (255, 0, 0),
        (150, 0, 0)
    ]
    sheet_name = xls_sheet.name
    df = pd.read_excel(filename, sheet_name, header=27, usecols='B:N')
    df.index = pd.Index(list(i for i in range(29, len(df) + 29)))
    # print(len(df.index))
    df.columns = pd.Index(list('BCDEFGHIJKLMN'))
    df = df.loc[list(i for i in range(29, 176))]
    # df.loc[29, 'B'] = gain_chara_color.keys[0]
    # print(df, type(df), df.loc[29, 'B'])

    gain_data_row = [col for i in range(0, 101) if i % 50 == 0 for col in range(i + 29, i + 42)]
    data_col = [let for let in "BCDEFGHIJKLMN"]
    chara_data_row = [col for i in range(0, 101) if i % 50 == 0 for col in range(i + 63, i + 76)]

    for col, row in product(data_col, gain_data_row):
        if df.loc[row, col] >= -6:
            df.loc[row, col] = 0
        elif -20 <= df.loc[row, col] < -16:
            df.loc[row, col] = 6
        elif df.loc[row, col] < -20:
            df.loc[row, col] = 7
        else:
            df.loc[row, col] = math.ceil(((-df.loc[row, col]) - 6) / 2)
    for col, row in product(data_col, chara_data_row):
        if df.loc[row, col] >= 15.4:
            df.loc[row, col] = 0
        elif 9.5 <= df.loc[row, col] < 15.4:
            df.loc[row, col] = 1
        elif 5.9 <= df.loc[row, col] < 9.5:
            df.loc[row, col] = 2
        elif 2.3 <= df.loc[row, col] < 5.9:
            df.loc[row, col] = 3
        elif 0 <= df.loc[row, col] < 2.3:
            df.loc[row, col] = 4
        elif -3.3 <= df.loc[row, col] < 0:
            df.loc[row, col] = 5
        elif -9.5 <= df.loc[row, col] < -3.3:
            df.loc[row, col] = 6
        elif df.loc[row, col] < -9.5:
            df.loc[row, col] = 7
    # print(df)
    for col, row in product(data_col, chara_data_row + gain_data_row):
        # print(col,row)
        xls_sheet.range(col + str(row)).color = gain_chara_color[int(df.loc[row, col])]


def get_data(filename):
    effi_list = []
    effi_data = {
        'l1': effi_list,
        'l5': effi_list,
        'bt': effi_list
    }
    gain_data = {
        '1170MHz': {
            '30°': 0,
            '45°': 0,
            '60°': 0,
            '90°': 0,
            '120°': 0
        },
        '1190MHz': {
            '30°': 0,
            '45°': 0,
            '60°': 0,
            '90°': 0,
            '120°': 0
        },
        '1210MHz': {
            '30°': 0,
            '45°': 0,
            '60°': 0,
            '90°': 0,
            '120°': 0
        },
        '1560MHz': {
            '30°': 0,
            '45°': 0,
            '60°': 0,
            '90°': 0,
            '120°': 0
        },
        '1580MHz': {
            '30°': 0,
            '45°': 0,
            '60°': 0,
            '90°': 0,
            '120°': 0
        },
        '1610MHz': {
            '30°': 0,
            '45°': 0,
            '60°': 0,
            '90°': 0,
            '120°': 0
        },

    }
    wb = xw.Book(filename)
    gps_sheets = [i for i in wb.sheet_names if ('L1' in i or 'L5' in i)]
    l1_sheets = [i for i in gps_sheets if 'L1' in i]
    l5_sheets = [i for i in gps_sheets if 'L5' in i]
    bt_sheets = [i for i in wb.sheet_names if 'BT' in i]
    # ws = wb.sheets[0]

    # get gain data
    for l1_sheet_name in l1_sheets:
        # print(l1_sheets)
        dut_name = l1_sheet_name.replace('L1-', '')
        ds = wb.sheets[l1_sheet_name]
        # print(ds)
        gain_value(ds, gain_data, 'l1')
        effi_value(ds, effi_data, 'l1')
        for l5_sheet_name in l5_sheets:
            ds = wb.sheets[l5_sheet_name]
            if dut_name in l5_sheet_name:
                gain_value(ds, gain_data, 'l5')
                effi_value(ds, effi_data, 'l5')
                break
            else:  # set to 0
                for key in gain_data:
                    for i in gain_data[key].keys():
                        if key in ['1170MHz', '1190MHz', '1210MHz']:
                            gain_data[key][i] = 0
                effi_data['l5'] = []

        for bt_sheet_name in bt_sheets:
            ds = wb.sheets(bt_sheet_name)
            if dut_name in bt_sheet_name:
                effi_value(ds, effi_data, 'bt')
            else:
                effi_data['bt'] = []
        dut_effi[dut_name] = copy.deepcopy(effi_data)
        dut_gain[dut_name] = copy.deepcopy(gain_data)
    # Coloring gain and characteristic data
    for gps_sheet_name in gps_sheets:
        gain_chara_coloring(filename, wb.sheets[gps_sheet_name])


def write_data(filename):
    freq_color = [(172, 185, 202), (255, 255, 0), (0, 176, 80), (255, 192, 0), (0, 112, 192), (146, 208, 80)]
    wb = xw.Book(filename)
    ws = wb.sheets[0]
    # print(ws)
    ws.range('A15:J50').value = ''
    ws.range('A15:J50').color = (255, 255, 255)
    # dut_num =  len(dut_data)
    dut_list = list(dut_gain.keys())
    freq_list = list(dut_gain[dut_list[0]].keys())
    theta_list = list(dut_gain[dut_list[0]][freq_list[0]].keys())
    # print(freq_list)
    # write gain data
    theta_col = {'30°': 'C', '45°': 'D', '60°': 'E', '90°': 'F', '120°': 'G'}
    for i in range(0, 6):
        ws.range('A' + (str(15 + len(dut_list) * i))).value = freq_list[i]
        # print(str('A' + str(15 + len(dut_list) * i)+':'+'A' + str(15 + len(dut_list) * i+len(dut_list))))
        ws.range('A' + str(15 + len(dut_list) * i) + ':' + 'A' + str(15 + len(dut_list) * i + len(dut_list) - 1),
                 'A' + str(15 + len(dut_list) * i) + ':' + 'G' + str(15 + len(dut_list) * i)).color = freq_color[i]
        for j in range(0, len(dut_list)):
            ws.range('B' + str(15 + len(dut_list) * i + j)).value = dut_list[j]
            # ws.range('B' + str(15 + len(dut_list) * i + j)).color = (255, 255, 204)
            for theta in theta_list:
                ws.range(theta_col[theta] + str(15 + len(dut_list) * i + j)).value = \
                    round(dut_gain[dut_list[j]][freq_list[i]][theta], 2)

    # write efficiency and s11 data
    ws.range('M2:V71').value = ''
    ws.range('L3:V71').color = (255, 255, 255)
    dut_col = {0: 'M', 1: 'N', 2: 'O', 3: 'P', 4: 'Q', 5: 'R', 6: 'S', 7: 'T', 8: 'U', 9: 'V'}
    for m in range(len(dut_list)):
        ws.range(dut_col[m] + str(2)).value = dut_list[m]
        effi = np.array([dut_effi[dut_list[m]]['l1']])
        ws.range(dut_col[m] + str(3)).value = effi.T
        effi = np.array([dut_effi[dut_list[m]]['l5']])
        ws.range(dut_col[m] + str(26)).value = effi.T
        effi = np.array([dut_effi[dut_list[m]]['bt']])
        ws.range(dut_col[m] + str(49)).value = effi.T
    ws.range('L9:' + dut_col[len(dut_list) - 1] + str(14)).color = (255, 255, 0)
    ws.range('L32:' + dut_col[len(dut_list) - 1] + str(38)).color = (255, 255, 0)
    ws.range('L54:' + dut_col[len(dut_list) - 1] + str(62)).color = (255, 255, 0)
    ws.range('L24:' + dut_col[len(dut_list) - 1] + str(24)).color = (255, 217, 100)
    ws.range('L47:' + dut_col[len(dut_list) - 1] + str(47)).color = (255, 217, 100)
    ws.range('L70:' + dut_col[len(dut_list) - 1] + str(70)).color = (255, 217, 100)
    ws.range('L25:' + dut_col[len(dut_list) - 1] + str(25)).color = (0, 176, 80)
    ws.range('L48:' + dut_col[len(dut_list) - 1] + str(48)).color = (0, 176, 80)
    ws.range('L71:' + dut_col[len(dut_list) - 1] + str(71)).color = (0, 176, 80)


if __name__ == "__main__":
    for file in glob.glob('./*.xls'):
        # init()
        print('文件 %s 处理中，请稍后......' % file[2:])
        try:
            start_time = time.perf_counter()
            get_data(file)
            write_data(file)
        except:
            print('数据记录错误，请检查sheet名称是否正确并确认测试数据是否填充完整！！')
        print('总计用时: %s 秒' % (round((time.perf_counter() - start_time), 2)))
    input('按任意键结束...')
