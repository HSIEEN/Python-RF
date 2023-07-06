# Data : 2023/6/30 15:32
# Author: Shawn Shi
# Right Reserved By COROS
import os.path
import sys

import skrf as rf
import numpy as np
import glob
import matplotlib.pyplot as plt
import datetime
import glob
import tkinter as tk
from tkinter import filedialog


def write_self_resonance_frequency(snp_file, line_number):
    """
    param snp_file: Murata part file with frequency unit of Hz
    :return: Null
    """
    snp_data = rf.Network(snp_file)
    if 'GRM' in snp_file:
        # If a capacitor, get min of S11
        freq_at_min = snp_data.f[np.argmin(snp_data.s_mag[:, 0, 0])]
    else:
        # If an inductor, get min of S21
        freq_at_min = snp_data.f[np.argmin(snp_data.s_mag[:, 1, 0])]
    # tmp = snp_data.f.min
    with open(snp_file, "r") as f:
        contents = f.readlines()
    if snp_data.f.min() < freq_at_min < snp_data.f.max():
        contents.insert(line_number, f'!The self-resonance frequency is {freq_at_min / 1e9} GHz \n')
        print(f'Self-resonance frequency of {snp_file} is {freq_at_min / 1e9} GHz')
    else:
        contents.insert(line_number, f'!No self-resonance frequency found within the frequency range\n')
        print(f'!No self-resonance frequency found within the frequency range')
    with open(snp_file, "w") as f:
        contents = "".join(contents)
        f.write(contents)


def remove_line(snp_file, line_number):
    """

    :param snp_file: snp file full path
    :param line_number: line number at which a new line is inserted, if -1 remove all comment lines
    :return: Non
    """
    if line_number != -1:
        with open(snp_file, "r") as f:
            contents = f.readlines()
            contents.pop(line_number)
        with open(snp_file, "w") as f:
            contents = "".join(contents)
            f.write(contents)
    else:
        with open(snp_file, "r") as f:
            contents = f.readlines()
            line_number = 0
            while contents[line_number][0] == '!':
                contents.pop(line_number)

        with open(snp_file, "w") as f:
            contents = "".join(contents)
            f.write(contents)


def shunt2series(s2p_file, dest_path):
    s2p_shunt = rf.Network(s2p_file)
    # plt.figure()
    s2p_shunt.s11.plot_s_db()
    s11 = s2p_shunt.s_re[:, 0, 0] + 1j * s2p_shunt.s_im[:, 0, 0]
    zc1 = s2p_shunt.z0[:, 0]
    zc2 = s2p_shunt.z0[:, 1]
    r = zc2 / zc1
    y = []
    for i in range(0, len(r)):
        y.append((1 - (1 / r[i] + 1) * s11[i] - 1 / r[i]) / (1 + s11[i]) / zc1[i])
    y = np.array(y)
    z = 1 / y
    A = []
    for i in range(0, len(y)):
        A.append([[1, z[i]], [0, 1]])
    A = np.array(A)
    frequency = rf.Frequency.from_f(s2p_shunt.f, unit='Hz')
    s2p_series = rf.Network(
        frequency=frequency, a=A, name='Shunt S-parameter',
        comments="Created by COROS at " +
                 str(datetime.datetime.now()) + '\nPartNumber: ' + s2p_file[s2p_file.rfind("\\") + 1:].replace(
            "_shunt.s2p", ""))
    s2p_series.write_touchstone(dest_path + s2p_file[s2p_file.rfind("\\") + 1:].replace("shunt", "series"))


def series2shunt(s2p_file, dest_path):
    # Obtain S-parameter of initial series s2p file
    s2p_series = rf.Network(s2p_file)

    s2p_series.s11.plot_s_db()
    # s11 of s-parameters at all frequencies
    s11 = s2p_series.s_re[:, 0, 0] + 1j * s2p_series.s_im[:, 0, 0]
    # Characteristic impedance at port 1
    zc1 = s2p_series.z0[:, 0]
    # Characteristic impedance at port 2
    zc2 = s2p_series.z0[:, 1]
    r = zc2 / zc1
    # Z at all frequencies, it is not a Z-matrix
    z = []
    # Get Z from s11 at all frequencies
    for i in range(0, len(r)):
        z.append((1 + (r[i] + 1) * s11[i] - r[i]) / (1 - s11[i]) * zc1[i])
    z = np.array(z)
    # Y, not Y-matrix
    y = 1 / z
    # ABCD-matrix
    A = []

    for i in range(0, len(y)):
        A.append([[1, 0], [y[i], 1]])
    A = np.array(A)
    # Construct a network from ABCD matrix
    frequency = rf.Frequency.from_f(s2p_series.f, unit='Hz')
    s2p_shunt = rf.Network(
        frequency=frequency, a=A, name='Series S-parameter',
        comments="Created by COROS at " + str(datetime.datetime.now()) + '\nPartNumber: ' + s2p_file[
                                                                                            s2p_file.rfind(
                                                                                                "\\") + 1:].replace(
            "_series.s2p", ""))
    s2p_shunt.write_touchstone(dest_path + s2p_file[s2p_file.rfind("\\") + 1:].replace("series", "shunt"))


def cascaded_s2p_generator(source_s2p1, source_s2p2, destination_directory):
    s2p_file_name1 = source_s2p1[source_s2p1.rfind("\\") + 1:]
    s2p_file_name2 = source_s2p2[source_s2p1.rfind("\\") + 1:]
    # Determine whether the component is shunt or series
    if 'shunt' in s2p_file_name1:
        sh_or_se1 = 'P'
    elif 'series' in s2p_file_name1:
        sh_or_se1 = 'S'
    else:
        sys.exit("No available data!")
    if 'shunt' in s2p_file_name2:
        sh_or_se2 = 'P'
    elif 'series' in s2p_file_name2:
        sh_or_se2 = 'S'
    else:
        sys.exit("No available data!")
    # Determine whether the component is a capacitor or an inductor, and subsequently obtain its value.
    value1 = s2p_file_name1[s2p_file_name1.find('_') + 1:s2p_file_name1.rfind('_')]
    if 'p' in value1:
        C_or_L1 = 'C'
    else:
        C_or_L1 = 'L'
    value2 = s2p_file_name2[s2p_file_name2.find('_') + 1:s2p_file_name2.rfind('_')]
    if 'p' in value2:
        C_or_L2 = 'C'
    else:
        C_or_L2 = 'L'
    s2p_data1 = rf.Network(source_s2p1)
    s2p_data2 = rf.Network(source_s2p2)
    # rf.cascade()
    # Determine whether the frequencies of two s2p files are identical; if not, extract a subset and resample it.
    if not np.array_equal(s2p_data1.f, s2p_data2.f):

        fmin = max(s2p_data1.f.min(), s2p_data2.f.min())
        fmax = min(s2p_data1.f.max(), s2p_data2.f.max())
        s2p_data1 = s2p_data1[str(fmin) + '-' + str(fmax) + 'hz']
        s2p_data2 = s2p_data2[str(fmin) + '-' + str(fmax) + 'hz']
        if s2p_data2.f.size > s2p_data1.f.size:
            s2p_data1 = s2p_data1.interpolate(s2p_data2.frequency)
        elif s2p_data2.f.size < s2p_data1.f.size:
            s2p_data2 = s2p_data2.interpolate(s2p_data1.frequency)
        # pass
    # Ignor PCPC, PLPL, SCSC, SLSL
    if not ((sh_or_se1 == sh_or_se2) and (C_or_L1 == C_or_L2)):
        # cascaded_data = s2p_data1**s2p_data2
        cascaded_data = rf.cascade(s2p_data1, s2p_data2)
        file_name = sh_or_se1 + value1 + sh_or_se2 + value2 + '.s2p'
        dir_name = sh_or_se1 + sh_or_se2 + '/' + sh_or_se1 + C_or_L1 + sh_or_se2 + C_or_L2
        if not os.path.exists(destination_directory + dir_name):
            os.makedirs(destination_directory + dir_name)
        cascaded_data.write_touchstone(
            file_name, destination_directory + dir_name)
        # remove all comments lines
        remove_line(
            destination_directory + dir_name + '/' + file_name, -1)
        print('=======================================')
        print(f"s2p file {file_name} generated!!")


# ====================================================================================================
if __name__ == '__main__':
    # remove_line(r"\\nas.local\DATA\Wireless\Library\Components\temporary use\d\S0p2S0p2.s2p", -1)
    print("==============请选择一个功能：=============")
    print("1. s2p文件自谐振频率计算")
    print("2. 串联s2p文件转换为并联s2p文件")
    print("3. 并联s2p文件转换为串联s2p文件")
    print("4. s2p文件级联")
    print("5. snp文件作图")
    print('=========================================')
    selection = input("请输入你的选择：")
    root = tk.Tk()
    root.withdraw()
    if selection == '1':
        print("============s2p文件自谐振频率计算与写入=============")
        print('************请选择一个源文件目录***************')
        source_file_path = filedialog.askdirectory()
        print(f"Source path is {source_file_path}")
        line_number = input("======请输入要写入自谐振频率的行号，如不确定，请输入-1======")
        if line_number == '-1':
            line = 0
        else:
            line = int(line_number)
        files = glob.glob(source_file_path + r"\*.s2p")
        for file in files:
            write_self_resonance_frequency(file, line)
    elif selection == '2':
        print("============将串联s2p文件转换为并联s2p文件并写入指定文件夹=============")
        print('************请选择一个源文件目录***************')
        source_file_path = filedialog.askdirectory()
        print(f"Source path is {source_file_path}")
        # path = r"\\nas.local\DATA\Wireless\Library\Components\for test"
        print('************请选择一个目标文件目录***************')
        # By default series to shunt
        dest_path = filedialog.askdirectory()
        print(f"Destination path is {dest_path}")
        # dest_path = r"\\nas.local\DATA\Wireless\Library\Components\for test\Shunt\\"
        files = glob.glob(source_file_path + r"\*.s2p")
        for file in files:
            series2shunt(file, dest_path + "\\")
    elif selection == '3':
        print("============将并联s2p文件转换为串联s2p文件并写入指定文件夹=============")
        print('************请选择一个源文件目录***************')
        source_file_path = filedialog.askdirectory()
        print(f"Source path is {source_file_path}")
        # path = r"\\nas.local\DATA\Wireless\Library\Components\for test"
        print('************请选择一个目标文件目录***************')
        # By default series to shunt
        dest_path = filedialog.askdirectory()
        print(f"Destination path is {dest_path}")
        # dest_path = r"\\nas.local\DATA\Wireless\Library\Components\for test\Shunt\\"
        files = glob.glob(source_file_path + r"\*.s2p")
        for file in files:
            shunt2series(file, dest_path + "\\")
    elif selection == '4':
        print("============将两个源目录中的s2p文件进行级联(s2p文件来自于Murata)，并写入指定目录中=============")
        print("*************请选择第一个源文件目录******************")
        source_file_path1 = filedialog.askdirectory()
        print(f"Source path1 is {source_file_path1}")
        print("*************请选择第二个源文件目录******************")
        source_file_path2 = filedialog.askdirectory()
        print(f"Source path2 is {source_file_path2}")
        print("*************请选择目标文件目录******************")
        dest_file_path = filedialog.askdirectory()
        print(f"Destination path is {dest_file_path}")
        source_files1 = glob.glob(source_file_path1 + r"\*s2p")
        source_files2 = glob.glob(source_file_path2 + r"\*s2p")
        # start = False
        for file1 in source_files1:
            # if '3n5' in file1 and 'shunt' in file1:
            #     start = True
            # if start:
            for file2 in source_files2:
                cascaded_s2p_generator(file1, file2, dest_file_path + "/")
                # remove_line()
    elif selection == '5':
        print("============snp文件作图=============")
        print("*************请选择一个snp文件******************")
        source_file = filedialog.askopenfilename()
        print(f"Source file is {source_file}")
        snp_data = rf.Network(source_file)
        rf.stylely()
        plt.figure()
        snp_data.plot_s_smith()
        plt.figure()
        snp_data.plot_s_db()
        plt.figure()
        snp_data.plot_s_deg()
        plt.show(block=True)
