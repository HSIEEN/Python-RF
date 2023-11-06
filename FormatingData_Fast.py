# Data : 2022/11/17 16:04
# Author: Shawn Shi
# Right Reserved By COROS
# Data : 2022/11/16 9:53
# Author: Shawn Shi
# Right Reserved By COROS

# import glob
import copy
import time
import math
import numpy as np
import pandas as pd
from itertools import product

dut_gain = {}
dut_effi = {}


# Read gain data from a xls sheet
def gain_value(xls_sheet, dict_name, freq):
    freq_list = list(dict_name.keys())
    theta_list = list(dict_name[freq_list[0]].keys())
    # print(xls_sheet.name)
    if freq == 'l1':
        for key in list(dict_name.keys())[3:7]:
            for i in range(0, 5):
                dict_name[key][theta_list[i]] = \
                    xls_sheet.range('S' + (str(((freq_list.index(key) - 3) * 50) + 28 + i))).value
            # dict_name[key]['d30'] = xls_sheet.range('S' + (str(((freq_list.index(key) - 3) * 50) + 36))).value
            dict_name[key]['dist'] = xls_sheet.range('R' + (str(((freq_list.index(key) - 3) * 50) + 40))).value
            # dict_name[key]['d90'] = xls_sheet.range('S' + (str(((freq_list.index(key) - 3) * 50) + 38))).value
    if freq == 'l5':
        for key in list(dict_name.keys())[0:3]:
            for i in range(0, 5):
                dict_name[key][theta_list[i]] = xls_sheet.range('S' + (str((freq_list.index(key) * 50) + 28 + i))).value
            # dict_name[key]['d30'] = xls_sheet.range('S' + (str((freq_list.index(key) * 50) + 36))).value
            # dict_name[key]['d60'] = xls_sheet.range('S' + (str((freq_list.index(key) * 50) + 37))).value
            dict_name[key]['dist'] = xls_sheet.range('R' + (str((freq_list.index(key) * 50) + 40))).value
    # return dict_name


keys = []


# initialize gain data in a specified frequency band
def gain_value_ini(dict_name, freq):
    global keys
    freq_list = list(dict_name.keys())
    theta_list = list(dict_name[freq_list[0]].keys())
    if freq == 'l1':
        keys = ['1560MHz', '1580MHz', '1610MHz']
    elif freq == 'l5':
        keys = ['1160MHz', '1180MHz', '1190MHz']
    for key in keys:
        for i in range(0, 6):
            dict_name[key][theta_list[i]] = 0


def effi_value(xls_sheet, effi_name, freq):
    if freq in ['l1', 'l5']:
        effi_name[freq] = xls_sheet.range('B3:B25').value
        color_range = 'A9:AC14' if freq == 'l1' else 'A9:AC12'
        xls_sheet.range(color_range).color = (255, 255, 0)
    else:  # read BT effi data
        effi_name[freq] = xls_sheet.range('B8:B28').value
        effi_name[freq] += [xls_sheet.range('B34').value, xls_sheet.range('B35').value]
        effi_name[freq] += [(xls_sheet.range('R38').value+xls_sheet.range('R39').value+xls_sheet.range('R40').value)/3]
        xls_sheet.range('A13:AD21').color = (255, 255, 0)


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
    if 'L1-' in sheet_name or 'L5-' in sheet_name:
        # df = pd.read_excel(filename, sheet_name, header=27, usecols='B:N')
        df = pd.read_excel(filename, sheet_name, header=27, usecols='B:N', engine='openpyxl')
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
            if 0 < df.loc[row, col] <= 3:
                df.loc[row, col] = 0
            elif 3 < df.loc[row, col] <= 6:
                df.loc[row, col] = 1
            elif 6 < df.loc[row, col] <= 10:
                df.loc[row, col] = 2
            elif 10 < df.loc[row, col] <= 18:
                df.loc[row, col] = 3
            elif df.loc[row, col] > 18:
                df.loc[row, col] = 4
            elif df.loc[row, col] <= -14:
                df.loc[row, col] = 5
            elif -14 < df.loc[row, col] <= -6:
                df.loc[row, col] = 6
            elif -6 < df.loc[row, col] < 0:
                df.loc[row, col] = 7
        # print(df)
        for col, row in product(data_col, chara_data_row + gain_data_row):
            # print(col,row)
            xls_sheet.range(col + str(row)).color = gain_chara_color[int(df.loc[row, col])]
    elif 'BT-' in sheet_name:
        df = pd.read_excel(filename, sheet_name, header=37, usecols='B:N', engine='openpyxl')
        df.index = pd.Index(list(i for i in range(39, len(df) + 39)))
        # print(len(df.index))
        df.columns = pd.Index(list('BCDEFGHIJKLMN'))
        df = df.loc[list(i for i in range(39, 89))]
        # df.loc[29, 'B'] = gain_chara_color.keys[0]
        # print(df, type(df), df.loc[29, 'B'])

        gain_data_row = [col for i in range(0, 37) if i % 18 == 0 for col in range(i + 39, i + 52)]
        data_col = [let for let in "BCDEFGHIJKLMN"]
        # chara_data_row = [col for i in range(0, 101) if i % 50 == 0 for col in range(i + 63, i + 76)]

        for col, row in product(data_col, gain_data_row):
            if df.loc[row, col] >= -6:
                df.loc[row, col] = 0
            elif -20 <= df.loc[row, col] < -16:
                df.loc[row, col] = 6
            elif df.loc[row, col] < -20:
                df.loc[row, col] = 7
            else:
                df.loc[row, col] = math.ceil(((-df.loc[row, col]) - 6) / 2)
        for col, row in product(data_col, gain_data_row):
            # print(col,row)
            xls_sheet.range(col + str(row)).color = gain_chara_color[int(df.loc[row, col])]


def get_data(filename, wb):
    effi_list = []
    effi_data = {
        'l1': effi_list,
        'l5': effi_list,
        'bt': effi_list
    }

    frequencies = ['1160MHz', '1180MHz', '1190MHz', '1560MHz', '1580MHz', '1610MHz']
    # angles = ['30°', '45°', '60°', '90°', '120°', 'd30', 'd60', 'd90']
    angles = ['30°', '45°', '60°', '90°', '120°', 'dist']
    gain_data = {}
    for freq in frequencies:
        gain_data[freq] = {}
        for angle in angles:
            gain_data[freq][angle] = 0

    dut_effi.clear()
    dut_gain.clear()
    # dut_gain = {}
    # dut_name = []
    # wb = xw.Book(filename)
    # app = xw.App()
    # wb = xw.Book(filename)
    gps_sheets = [i for i in wb.sheet_names if ('L1-' in i or 'L5-' in i)]
    l1_sheets = [i for i in gps_sheets if 'L1-' in i]
    l5_sheets = [i for i in gps_sheets if 'L5-' in i]
    bt_sheets = [i for i in wb.sheet_names if 'BT-' in i]

    # Collect the dut names
    l1_dut_name = [i.replace('L1-', '') for i in l1_sheets]
    l5_dut_name = [i.replace('L5-', '') for i in l5_sheets]
    bt_dut_name = [i.replace('BT-', '') for i in bt_sheets]
    dut_name = list(set(l1_dut_name + l5_dut_name + bt_dut_name))

    # get gain and efficiency data
    for dut in dut_name:
        if dut in l1_dut_name:
            ds = wb.sheets['L1-' + dut]
            gain_value(ds, gain_data, 'l1')
            effi_value(ds, effi_data, 'l1')
        else:
            effi_data['l1'] = []
            gain_value_ini(gain_data, 'l1')
        if dut in l5_dut_name:
            ds = wb.sheets['L5-' + dut]
            gain_value(ds, gain_data, 'l5')
            effi_value(ds, effi_data, 'l5')
        else:
            effi_data['l5'] = []
            gain_value_ini(gain_data, 'l5')
        if dut in bt_dut_name:
            effi_data['bt'] = []
            ds = wb.sheets['BT-' + dut]
            effi_value(ds, effi_data, 'bt')
        else:
            effi_data['bt'] = []
        dut_effi[dut] = copy.deepcopy(effi_data)
        dut_gain[dut] = copy.deepcopy(gain_data)

    # Coloring gain and characteristic data
    for sheet in gps_sheets + bt_sheets:
        gain_chara_coloring(filename, wb.sheets[sheet])

    # wb.save()

    # return wb


def write_data(wb):
    freq_color = [(172, 185, 202), (255, 255, 0), (0, 176, 80), (255, 192, 0), (0, 112, 192), (146, 208, 80)]
    # wb = xw.App().books(filename)
    ws = wb.sheets['Conclusion']
    # ws.range('A1:I1').column_width = 20
    ws.range('A15:A100').api.HorizontalAlignment = -4108  # center
    ws.range('C15:J100').api.HorizontalAlignment = -4108
    ws.range('B15:B100').api.HorizontalAlignment = -4131  # Left
    ws.range('A15:J100').api.Font.Bold = True
    ws.range('I15:I100').api.Font.ColorIndex = 3
    ws.range('I15:I100').api.NumberFormat = "0.0"
    ws.range('M72:Z72').api.NumberFormat = "0.0"
    ws.range('M72:Z72').api.Font.ColorIndex = 3
    ws.range('M72:Z72').api.HorizontalAlignment = -4108
    # print(ws)
    ws.range('A15:J100').value = ''
    ws.range('A15:J100').color = (255, 255, 255)
    # dut_num =  len(dut_data)
    dut_list = list(dut_gain.keys())
    freq_list = list(dut_gain[dut_list[0]].keys())
    theta_list = list(dut_gain[dut_list[0]][freq_list[0]].keys())
    # print(freq_list)
    # write gain data
    # theta_col = {'30°': 'C', '45°': 'D', '60°': 'E', '90°': 'F', '120°': 'G', 'd30': 'H', 'd60': 'I', 'd90': 'J'}
    theta_col = {'30°': 'C', '45°': 'D', '60°': 'E', '90°': 'F', '120°': 'G', 'dist': 'I'}
    # record if there are L5 data, no in default.
    # min_freq_index
    min_freq_index = 3
    # dut_omitted is the number of DUT without gain data
    dut_omitted = 0
    max_freq_index = 3
    for dut in dut_list:
        # there is l5
        if dut_gain[dut][freq_list[0]][theta_list[0]] != 0:
            min_freq_index = 0
        # no l1 and l5
        if dut_gain[dut][freq_list[3]][theta_list[0]] == 0 and dut_gain[dut][freq_list[0]][theta_list[0]] == 0:
            dut_omitted = dut_omitted + 1
        # there is l1
        if dut_gain[dut][freq_list[3]][theta_list[0]] != 0:
            max_freq_index = 6
    # write gain data
    if len(dut_list) - dut_omitted > 0:
        for i in range(min_freq_index, max_freq_index):
            ws.range('A' + (str(15 + (len(dut_list) - dut_omitted) * (i - min_freq_index)))).value = freq_list[i]
            # print(str('A' + str(15 + len(dut_list) * i)+':'+'A' + str(15 + len(dut_list) * i+len(dut_list))))
            ws.range(
                'A' + str(15 + (len(dut_list) - dut_omitted) * (i - min_freq_index)) + ':' +
                'A' + str(
                    15 + (len(dut_list) - dut_omitted) * (i - min_freq_index) + (len(dut_list) - dut_omitted) - 1),
                'A' + str(15 + (len(dut_list) - dut_omitted) * (i - min_freq_index)) + ':' +
                'J' + str(15 + (len(dut_list) - dut_omitted) * (i - min_freq_index))).color = freq_color[
                i]
            # k= 0  # the number of line to be ignored
            k = 0
            for j in range(0, len(dut_list)):
                if dut_gain[dut_list[j]][freq_list[0]][theta_list[0]] != 0 or \
                        dut_gain[dut_list[j]][freq_list[3]][theta_list[0]] != 0:
                    ws.range('B' + str(15 + (len(dut_list) - dut_omitted) * (i - min_freq_index) + j - k)).value = \
                        dut_list[j]
                    # ws.range('B' + str(15 + len(dut_list) * i + j)).color = (255, 255, 204)
                    for theta in theta_list:
                        ws.range(
                            theta_col[theta] + str(15 + (len(dut_list)
                                                         - dut_omitted) * (i - min_freq_index) + j - k)).value = \
                            round(dut_gain[dut_list[j]][freq_list[i]][theta], 2)
                else:
                    k = k + 1

    # write efficiency data
    ws.range('M2:Z71').value = ''
    ws.range('L3:Z71').color = (255, 255, 255)
    ws.range('M2:Z71').api.HorizontalAlignment = -4108  # center
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
    ws.range('L32:' + dut_col[len(dut_list) - 1] + str(35)).color = (255, 255, 0)
    ws.range('L54:' + dut_col[len(dut_list) - 1] + str(62)).color = (255, 255, 0)
    ws.range('L24:' + dut_col[len(dut_list) - 1] + str(24)).color = (255, 217, 100)
    ws.range('L47:' + dut_col[len(dut_list) - 1] + str(47)).color = (255, 217, 100)
    ws.range('L70:' + dut_col[len(dut_list) - 1] + str(70)).color = (255, 217, 100)
    ws.range('L25:' + dut_col[len(dut_list) - 1] + str(25)).color = (0, 176, 80)
    ws.range('L48:' + dut_col[len(dut_list) - 1] + str(48)).color = (0, 176, 80)
    ws.range('L71:' + dut_col[len(dut_list) - 1] + str(71)).color = (0, 176, 80)
    # wb.save()
    # wb.close()
    # app.kill()
    # xw.App().kill()


def formatting_data(file, wb):
    # if_exit = 'Y'
    # while if_exit != 'N' and if_exit != 'n' and (if_exit == 'Y' or if_exit == 'y'):
    # root = tk.Tk()
    # root.withdraw()
    # print('---------Format antenna gain data_version 6.3-----------')
    # print('-----------All rights are reserved by COROS------------')
    # 'file_or_directory = input('============请选择文件=============')
    # print('*****************请选择一个文件******************')
    # file = filedialog.askopenfile()
    # name = file.name
    # try:
    # app = xw.App(visible=True, add_book=False)
    # app.display_alerts = True
    # app.screen_updating = True

    get_data(file, wb)
    time.sleep(0.5)
    write_data(wb)
    time.sleep(0.5)
    wb.save()
    # except:
    # print('数据记录错误，请检查sheet名称是否正确并确认测试数据是否填充完整！！')
