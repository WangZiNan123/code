import pandas as pd
import copy
import openpyxl
from datetime import datetime
import numpy as np
import os

# ================================================= #
# 2024_3_29 版本更新：2024.3.29
# 更新内容：
# 1.读取excel表格里面所有页数据，如白石，楼下机房

# 2024_3_29_A 版本更新：2024.3.29
# 1.增加外置液位，内置液位毫米（mm）的计算 ，当 外置液位，内置液位以升（L）为单位的值为空时，改用毫米（mm）计算液位消耗，如白石，楼下机房
# 此时 燃料消耗率(L.kWh -1) 不参与计算，值等于0。因为单位是升（L），毫米不参与计算

# 2024_3_30_A 版本更新：2024.3.30
# 1.对最后文件保存做出判断，如果检测到没有发电数据，不保存文件。'总发电功率(W)': everytime_power列表,如果全部都等于0，则没有发电
# 2.修复发电量为负数的情况，如果小于0，则‘发电量(kw/h)’=0。

# 2024_5_31_A 版本更新：2024.5.31
# 增加发电时 电堆A1平均温度（Statem1）, 电堆A2平均温度（Statem2）, 电堆B平均温度（FcB_StackT）

# ================================================= #

new_df = []
# 打印行号和列的数据
A_Power_values = []
B_Power_values = []
power_values = []  # 储存发电时的功率值
IC_value = []  # 储存发电时的芯片温度值
Topgen_value = []  # 储存每次发电，开始/结束的发电量值
Once_Topgen_value = []  # 储存，每次发电的发电量。用于算出总发电量
Time_value = []  # 储存每次发电，开始/结束时间的值
Time_diffs = []  # 储存，每次发电的时间的时长。用于算出总发电时间
differences = []
total_sum = 0
fuel_levels = []
last_fuel_levels = []
S_RemFuelIn_value = []
positive_differences = []
calculate_positive_differences = []
Once_S_RemFuelIn = []
B_StackV_value = []
A_StackV_value = []
B_List = []
A_List = []
last_A_List = []
last_B_List = []
HGretem_value = []  # 发电时，储存 重整室温度的值到列表 HGretem_value
Hfetem_value = []  # 发电时，储存 重整室温度的值到列表 Hfetem_value
HGretem_list = []
Hfetem_list = []
last_HGretem_list = []
last_Hfetem_list = []
start_datatime = []
end_datatime = []
start_S_RemFuelIn = []
end_S_RemFuelIn = []
start_Topgen = []
end_Topgen = []
start_S_RemFuelOut = []
end_S_RemFuelOut = []
Stwtims = []
Fuel_consumption = None

current_voltage = []
current_voltage_value = []
everytime_current_voltage = []
current_voltage_List_value = []
last_current_voltage_List_value = []

everytime_Topgen = []
everytime_A_power = []
everytime_B_power = []
everytime_power = []
everytime_IC = []
everytime_A_StackV = []
everytime_B_StackV = []
everytime_max_HGretem = []
everytime_min_HGretem = []
everytime_max_Hfetem = []
everytime_min_Hfetem = []
everytime_Fuel_consumption = []

copy_everytime_A_StackV = []
copy_everytime_B_StackV = []
copys_everytime_A_StackV = []
copys_everytime_B_StackV = []
copysS_everytime_A_StackV = []
copysS_everytime_B_StackV = []
modified_A_StackV = []
modified_B_StackV = []
count_end_datatime = []
fuel_List_value = []
last_A_power_value_list = []
last_B_power_value_list = []
last_power_value_list = []
power_list = []
start_time = None
second_start_time = None
copy_start_datatime = []
copy_end_datatime = []
count_datatime = []  # 开始时间+结束时间，放入一个列表里面。除以2余0.证明当天发电，开始和结束成一对。用于计算当天没有结束时的判断
first_start_datatime = 0
second_end_datatime = 0
df_list = []

true_LiqlelL = []  # 外置液位mm
true_LiqlelM = []  # 内置液位mm

start_LiqlelL = []
end_LiqlelL = []

start_LiqlelM = []
end_LiqlelM = []

A1_Stack_Temp_Value = []  # 电堆A1温度
A2_Stack_Temp_Value  = []  # 电堆A2温度
B_Stack_Temp_Value  = []  # 电堆B温度
everytime_A1_Stack_Temp = []       # 储存电堆A1温度的列表
everytime_A2_Stack_Temp = []      # 储存电堆A2温度的列表
everytime_B_Stack_Temp = []      # 储存电堆B温度的列表

b1 = '2024_4月管委会发电数据 test2'  # 储存 EXCEL表格 的文件名称
# adress2 = 'C:/Users/FCK/Desktop/12/test/%s.xlsx' % b1
adress3 = f"E:/远程下载数据/管委会/{b1}.xlsx"  # 储存 EXCEL表格文件 的路径
#  EXCEL格式为“某某 年，某某 月，某某 日” ，例如：”2023.10.1“这种格式.。"  年 . 月  . 日  "
year = 2024  # 年，表格的年
month = 4  # 月，表格的月

for i in range(12, 15):  # 遍历所有数据  i=8  range=31.   取值范围：8<= i <31
    # a1 = '2023.9.%s' % i
    # b1 = '2023_11_%s_test数据' %i
    a1 = '%d.%d.%d' % (year, month, i)  # 这个指令将会使用 year、month 和 i 的值来创建一个类似于 "XXXX.XX.XX" 格式的字符串，并将其存储在变量 a1 中。
    a1 = a1.strip()  # 这个指令会将变量 a1 中的字符串去掉开头和结尾的空白字符
    # 读取Excel文件中的数据
    adress1 = f'E:/远程下载数据/管委会/4月/{a1}.xlsx'  # 读取 EXCEL表格文件 的路径

    if os.path.exists(adress1):  # 检查文件（文件名，文件路径是对得上）是否存在，不存在则结束程序
        try:
            # 在这里进行对数据的处理和分析
            # df = pd.read_excel(adress1)
            # df['电堆总功率'] = df['Stapow'] + df['FcB_StackP']
            # 创建Series对象并使用NaN值填充不同长度的列数据，然后将这些Series对象传递给DataFrame构造函数

            # xl = pd.ExcelFile(adress1)  # 使用 pd.ExcelFile() 方法打开 Excel 文件
            # # # sheet_name = xl.sheet_names  # 获取文件中的所有工作表名称
            # # # 如果页面超过2页，把页面合并成一页
            # # # df = []
            # # # if len(sheet_name) >= 2:
            # for sheet_name in xl.sheet_names:  # 遍历文件中的所有 sheet
            #     one_sheet = xl.parse(sheet_name)  # 读取当前 sheet 的数据
            #     df.append(one_sheet)  # 将读取的数据合并到 all_data 中
            # # # 使用 pd.concat() 方法将所有数据框连接成一个
            # df = pd.concat(df, ignore_index=True)

            # # 如果页面只有1页，就直接读取第一页
            # else:
            #     one_sheet = xl.parse(sheet_name)  # 读取当前 sheet 的数据
            #     df = one_sheet  # 将读取的数据合并到 all_data 中

            # xl = pd.ExcelFile(adress1)  # 使用 pd.ExcelFile() 方法打开 Excel 文件
            df_list = []  # 初始化一个空的DataFrame列表，用于存储每个工作表的数据
            # 使用 'with' 语句打开Excel文件
            with pd.ExcelFile(adress1) as xl:
                for sheet_name in xl.sheet_names:  # 遍历文件中的所有 sheet
                    one_sheet = xl.parse(sheet_name)  # 读取当前 sheet 的数据
                    df_list.append(one_sheet)  # 将读取的数据添加到df_list中

            df = pd.concat(df_list, ignore_index=True)  # 使用 pd.concat() 方法将所有数据框连接成一个

            # 使用fillna()方法来替换DataFrame中的NaN值。如果你想要将所有的NaN值替换为0，可以直接调用方法 fillna(0)
            df.fillna(0, inplace=True)

            df['电堆总功率'] = df['Stapow'] + df['FcB_StackP']

            # 选择要读取的列名
            MSw = 'MSw'  # 开关状态
            DateTime = 'DateTime'  # 时间
            S_RemFuelIn = 'S_RemFuelIn'  # 内置水箱液位
            S_RemFuelOut = 'S_RemFuelOut'  # 外置水箱液位
            Topgen = 'Topgen'  # 发电量
            IC_Temp = 'Chiptem'  # 芯片温度
            A_Power = 'Stapow'  #
            B_Power = 'FcB_StackP'
            Power = '电堆总功率'
            prev_row = None
            B_StackV = 'FcB_StackV'  # 电堆B电压
            A_StackV = 'StaV'  # 电堆A电压
            HGretem = 'HGretem'  # 重整室温度
            Hfetem = 'Hfetem'  # 提纯器温度
            Stwtim = 'Stwtim'  # 发电次数
            S_CurVol = 'S_CurVol'  # 母线电压

            LiqlelL = 'LiqlelL'  # 外置液位（mm）
            LiqlelM = 'LiqlelM'  # 内置液位（mm）

            A1_Stack_Temp = 'Statem1'  # 电堆A1温度
            A2_Stack_Temp = 'Statem2'   # 电堆A2温度
            B_Stack_Temp = 'FcB_StackT'    # 电堆B温度

            #   打印有多少行
            # print('==========电堆电压', df['StaV'])

            #   如果电压小于85，则跳过当天计算
            if any(df['StaV'] >= 60):
                second_row = df.iloc[1]  # 这行代码将DataFrame中的第二行数据存储在变量second_row中，以便后续对第二行数据进行操作和分析
                last_row = df.iloc[-1]  # 这行代码将DataFrame中的最后一行数据存储在变量last_row中，以便后续对最后一行数据进行操作和分析


                # #  !!!  如果计算对象是 “众宇电堆” 筛选范围选择：  ９２ ＜＝ Ｘ ＜ １２５
                # #  !!!  如果计算对象是 “攀业电堆” 筛选范围选择：  ７５ ＜＝ Ｘ ＜ １２０
                # 对电堆电压算平均值 。
                def calculate_filtered_average(data):
                    filtered_data = [x for x in data if 75 <= x < 125]  # 设置筛选范围
                    average = sum(filtered_data) / len(filtered_data) if len(filtered_data) > 0 else 0  # 计算平均值
                    return average


                # 对发电功率算平均值,计算列表元素十个最大值平均值
                def calculate_average(input_list):
                    # 去掉小于100的元素并重新生成列表
                    new_list = [x for x in input_list if x >= 100]

                    if len(new_list) > 10:  # 如果新列表元素个数大于10
                        top_values = sorted(new_list, reverse=True)[:10]  # 找出新列表元素十个最大值
                        average = sum(top_values) / 10  # 计算平均值
                        return average
                    elif len(set(new_list)) == 1:  # 如果所有元素都相等
                        return new_list[0]  # 返回任意一个元素的值作为平均值
                    elif 0 < len(new_list) <= 10:  # 如果新列表元素个数小于等于10且不为空
                        average = sum(new_list) / len(new_list)  # 计算所有元素的平均值
                        return average
                    else:
                        if len(new_list) == 0:  # 如果新列表为空
                            return 0


                print('\n ————————————————    一天计算开始    ————————————————    \n')

                for index, row in df.iterrows():  # 这段代码会遍历 DataFrame df 中的每一行数据。

                    if prev_row is not None:  # 这段代码检查变量 prev_row 是否为非空值。

                        if row[MSw] == True:  # 如果MSW=TRUE，发电时，储存发电时间段内某列的数据
                            A_Power_values.append(round(row[A_Power], 1))
                            B_Power_values.append(round(row[B_Power], 1))
                            power_values.append(round(row[Power], 1))  # 发电时，储存 功率 的值到列表 power_values
                            IC_value.append(round(row[IC_Temp], 1))  # 发电时，储存 芯片温度 的值到列表 power_values

                            S_RemFuelIn_value.append(
                                round(row[S_RemFuelIn], 1))  # 发电时，储存 内置水箱剩余燃料(L) 的值到列表 S_RemFuelIn_value

                            B_StackV_value.append(round(row[B_StackV], 1))  # 发电时，储存 电堆B电压 的值到列表 B_StackV_value
                            A_StackV_value.append(round(row[A_StackV], 1))  # 发电时，储存 电堆A电压 的值到列表 A_StackV_value
                            HGretem_value.append(round(row[HGretem], 1))  # 发电时，储存 重整室温度的值到列表 HGretem_value
                            Hfetem_value.append(round(row[Hfetem], 1))  # 发电时，储存 提纯室温度的值到列表 Hfetem_value

                            current_voltage_value.append(round(row[S_CurVol], 1))  # 发电时，储存 母线电压的值到 current_voltage

                            true_LiqlelM.append(round(row[LiqlelM], 2))  # 发电时，储存 内置水箱剩余燃料(mm) 的值到列表 true_LiqlelM
                            true_LiqlelL.append(round(row[LiqlelL], 2))  # 发电时，储存 外置水箱剩余燃料(mm) 的值到列表 true_LiqlelL

                            A1_Stack_Temp_Value.append(round(row[A1_Stack_Temp], 1))  # 储存电堆A1的温度
                            A2_Stack_Temp_Value.append(round(row[A2_Stack_Temp], 1))  # 储存电堆A2的温度
                            B_Stack_Temp_Value.append(round(row[B_Stack_Temp], 1))  # 储存电堆B的温度


                        if prev_row[MSw] == False and row[MSw] == True:  # 开始发电时间 。 如果MSW的上一个值=false,并且当前的值=true
                            print(f"\n第一有开始 ###############\n")
                            print(  # 在控制台上打印，显示
                                f"开始发电时间：{row[DateTime]}      "
                                f"内置水箱剩余燃料(L): {round(row[S_RemFuelIn], 1)}     "
                                f"外置水箱剩余燃料(L): {round(row[S_RemFuelOut], 1)}  "

                                f"内置水箱剩余燃料(mm): {round(row[LiqlelM], 1)} "
                                f"外置水箱剩余燃料(mm): {round(row[LiqlelL], 1)} "
                                f"总发电量:{round(row[Topgen], 1)}      ")
                            Topgen_value.append(round(row[Topgen], 1))
                            Time_value.append(row[DateTime])
                            count_end_datatime.append(row[DateTime])
                            second_start_time = row[DateTime]  # 用于后面当天发电缺少“开始发电”的判断

                            # 创建列表用于储存输出到excel表格和数据
                            start_datatime.append(row[DateTime])

                            # copy_start_datatime.clear()
                            # copy_start_datatime.append(row[DateTime])
                            # first_start_datatime = len(copy_start_datatime)
                            # print(f"计数 copy_start_datatime 》》》》》：{copy_start_datatime}")
                            # print(f"个数 first_start_datatime 》》》》》：{first_start_datatime}")

                            start_S_RemFuelIn.append(round(row[S_RemFuelIn], 1))
                            start_Topgen.append(round(row[Topgen], 1))
                            start_S_RemFuelOut.append(round(row[S_RemFuelOut], 1))

                            start_LiqlelL.append(round(row[LiqlelL], 1))
                            start_LiqlelM.append(round(row[LiqlelM], 1))

                        else:

                            if second_start_time is None and second_row[MSw] == True:  #
                                print(f"\n第二没有开始 ************\n")
                                print(
                                    f"开始发电时间：{second_row[DateTime]}     "
                                    f" 内置水箱剩余燃料(L): {round(second_row[S_RemFuelIn], 1)}    "
                                    f" 外置水箱剩余燃料(L): {round(second_row[S_RemFuelOut], 1)}"

                                    f"内置水箱剩余燃料(mm): {round(row[LiqlelM], 1)} "
                                    f"外置水箱剩余燃料(mm): {round(row[LiqlelL], 1)} "
                                    f"    总发电量:{round(second_row[Topgen], 1)}      ")
                                Topgen_value.append(round(second_row[Topgen], 1))
                                Time_value.append(second_row[DateTime])
                                second_start_time = second_row[DateTime]
                                count_end_datatime.append(second_row[DateTime])
                                # 创建列表用于储存输出到excel表格和数据
                                start_datatime.append(second_row[DateTime])
                                copy_start_datatime.append(second_row[DateTime])
                                first_start_datatime = len(copy_start_datatime)
                                # print(f"计数 count_datatime 》》》》》：{copy_start_datatime}")
                                # print(f"个数 count_datatime 》》》》》：{first_start_datatime}")

                                start_S_RemFuelIn.append(round(second_row[S_RemFuelIn], 1))
                                start_Topgen.append(round(second_row[Topgen], 1))
                                start_S_RemFuelOut.append(round(second_row[S_RemFuelOut], 1))

                                start_LiqlelL.append(round(row[LiqlelL], 1))
                                start_LiqlelM.append(round(row[LiqlelM], 1))
                        if prev_row[MSw] == True and row[MSw] == False:  # 结束发电时间。如果MSW的上一个值=true,并且当前的值=false

                            print(
                                f"结束发电时间：{prev_row[DateTime]}      "
                                f"内置水箱剩余燃料(L): {round(prev_row[S_RemFuelIn], 1)}     "
                                f"外置水箱剩余燃料(L): {round(prev_row[S_RemFuelOut], 1)}   "
                                f"内置水箱剩余燃料(mm): {round(prev_row[LiqlelM], 1)} "
                                f"外置水箱剩余燃料(mm): {round(prev_row[LiqlelL], 1)}"
                                f"总发电量:{round(prev_row[Topgen], 1)}    ")

                            print(len(count_end_datatime))  # 计算当天发电次数
                            Topgen_value.append(round(prev_row[Topgen], 1))
                            Time_value.append(prev_row[DateTime])
                            start_time = prev_row[DateTime]  # 用于后面当天发电缺少“结束发电”的判断

                            # 创建列表用于储存输出到excel表格和数据
                            end_datatime.append(prev_row[DateTime])

                            # count_datatime，用于计数。一天当天“开始+结束”的
                            # count_datatime.clear()
                            # count_datatime.append(prev_row[DateTime])

                            # second_end_datatime = len(count_datatime)
                            # copy_start_datatime.clear()
                            # print(f"计数 count_datatime 》》》》》：{count_datatime}")
                            # print(f"个数 count_datatime 》》》》》：{second_end_datatime}")

                            end_S_RemFuelIn.append(round(prev_row[S_RemFuelIn], 1))
                            end_Topgen.append(round(prev_row[Topgen], 1))
                            end_S_RemFuelOut.append(round(prev_row[S_RemFuelOut], 1))

                            end_LiqlelL.append(round(prev_row[LiqlelL], 1))
                            end_LiqlelM.append(round(prev_row[LiqlelM], 1))

                            Once_Topgen = round(Topgen_value[-1] - Topgen_value[-2], 3)
                            if Once_Topgen < 0:
                                Once_Topgen = 0
                            print(f"每次发电量(kw/h)：{Once_Topgen}")
                            Once_Topgen_value.append(Once_Topgen)

                            Stwtims.append(row[Stwtim])
                            print(f"发电次数：{row[Stwtim]}")

                            Time_diff = round(
                                (pd.to_datetime(Time_value[-1]) - pd.to_datetime(Time_value[-2])).total_seconds() / 60,
                                2)
                            Time_diffs.append(Time_diff)
                            print(f"每次发电时长(min)：{Time_diff}")

                            mean_IC = round(sum(IC_value) / len(IC_value), 2)
                            everytime_IC.append(mean_IC)
                            print(f'芯片平均温度(℃):{mean_IC}')

                            mean_A1_Stack_Temp = round(sum(A1_Stack_Temp_Value) / len(A1_Stack_Temp_Value), 2)
                            everytime_A1_Stack_Temp.append(mean_A1_Stack_Temp)
                            print(f'电堆A1平均温度(℃):{mean_A1_Stack_Temp}')

                            mean_A2_Stack_Temp = round(sum(A2_Stack_Temp_Value) / len(A2_Stack_Temp_Value), 2)
                            everytime_A2_Stack_Temp.append(mean_A2_Stack_Temp)
                            print(f'电堆A2平均温度(℃):{mean_A2_Stack_Temp}')

                            mean_B_Stack_Temp = round(sum(B_Stack_Temp_Value) / len(B_Stack_Temp_Value), 2)
                            everytime_B_Stack_Temp.append(mean_B_Stack_Temp)
                            print(f'电堆B平均温度(℃):{mean_B_Stack_Temp}')

                            Once_RemFuelIn = 0

                            # 一天只发一次电时，执行下面程序
                            if len(count_end_datatime) == 1:

                                current_voltage = round(sum(current_voltage_value) / len(current_voltage_value), 1)
                                everytime_current_voltage.append(current_voltage)
                                current_voltage_value.clear()
                                print(f'母线电压平均值(W)：{current_voltage}')

                                calculate_A_power = round(calculate_average(A_Power_values), 1)
                                everytime_A_power.append(calculate_A_power)
                                A_Power_values.clear()
                                print(f'A堆功率平均值(W)：{calculate_A_power}')

                                calculate_B_power = round(calculate_average(B_Power_values), 1)
                                everytime_B_power.append(calculate_B_power)
                                B_Power_values.clear()
                                print(f'B堆功率平均值(W)：{calculate_B_power}')

                                calculate_power = round(calculate_average(power_values), 1)
                                everytime_power.append(calculate_power)
                                power_values.clear()
                                print(f'总功率平均值(W)：{calculate_power}')

                                print(f'S_RemFuelIn_value[0]：{S_RemFuelIn_value[0]}')
                                if S_RemFuelIn_value[0] > 0:
                                    differences = [S_RemFuelIn_value[i] - S_RemFuelIn_value[i + 1] for i in
                                                   range(len(S_RemFuelIn_value) - 1)]
                                    positive_differences = [x for x in differences if x > 0]
                                    Once_RemFuelIn = round(sum(positive_differences), 2)
                                    if Once_RemFuelIn == 0:
                                        Once_RemFuelIn = 0.3
                                    Once_S_RemFuelIn.append(Once_RemFuelIn)
                                    print(f'每次发电消耗燃料（L）:{Once_RemFuelIn}')
                                    S_RemFuelIn_value.clear()  # 用完S_RemFuelIn_value列表后，要把列表清空，不然会叠加列表
                                else:
                                    differences = round(start_LiqlelM[-1] - end_LiqlelM[-1], 1)
                                    if differences < 0:
                                        differences = 0
                                    Once_S_RemFuelIn.append(differences)
                                    print(f'每次发电消耗燃料（mm）:{differences}')
                                    # print(f'液位(mm)******** ：{differences}')

                                # 计算发电过程中，A电堆电压平均值（过滤小于90和大于130的值）
                                average_A_StackV = round(calculate_filtered_average(A_StackV_value), 1)
                                everytime_A_StackV.append(average_A_StackV)
                                #### 2023.1.16新增
                                copy_everytime_A_StackV = copy.deepcopy(everytime_A_StackV)
                                copys_everytime_A_StackV.append(copy_everytime_A_StackV)
                                modified_A_StackV = [item[0] for item in copys_everytime_A_StackV]
                                ######
                                print(f'A电堆平均电压(V):{average_A_StackV}', end="        ")
                                # print(f'A电堆平均电压  -------- (V):{A_StackV_value}')
                                A_StackV_value.clear()  # 用完A_StackV_value列表后，要把列表清空，不然会叠加列表
                                everytime_A_StackV.clear()  # everytime_A_StackV 用于计算平均值。每次算完后列表清零

                                # 计算发电过程中，B电堆电压平均值（过滤小于90和大于130的值）
                                # everytime_B_StackV 用于计算平均值。每次算完后列表清零
                                average_B_StackV = round(calculate_filtered_average(B_StackV_value), 1)
                                everytime_B_StackV.append(average_B_StackV)
                                #### 2023.1.16新增
                                copy_everytime_B_StackV = copy.deepcopy(everytime_B_StackV)
                                copys_everytime_B_StackV.append(copy_everytime_B_StackV)
                                modified_B_StackV = [item[0] for item in copys_everytime_B_StackV]
                                ######
                                print(f'B电堆平均电压(V):{average_B_StackV}')

                                # print(f'B电堆平均电压  -------- (V):{B_StackV_value}')

                                B_StackV_value.clear()  # 用完B_StackV_value列表后，要把列表清空，不然会叠加列表
                                everytime_B_StackV.clear()  # everytime_B_StackV 用于计算平均值。每次算完后列表清零

                                if all(item == 0 for item in HGretem_value) and all(item == 0 for item in Hfetem_value):

                                    max_HGretem = 0
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = 0
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')

                                    # print(f'重整室最列表温度^^^^^^^^^^^^5(℃)：{HGretem_value}')
                                    HGretem_value = []  # 用完HGretem_value列表后，要把列表清空，不然会叠加列表

                                    max_Hfetem = 0
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = 0
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                    # print(f'提纯器温度列表^^^^^^^^^^^^^^^6(℃)：{Hfetem_value}')
                                    Hfetem_value = []

                                else:
                                    #   使用列表推导式过滤了列表 HGretem_value 中值为 0 的元素，并将结果重新赋值给 HGretem_value
                                    HGretem_value = [x for x in HGretem_value if x != 0]
                                    max_HGretem = round(max(HGretem_value), 1)
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = round(min(HGretem_value), 1)
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')

                                    # print(f'重整室最列表温度^^^^^^^^^^^^5(℃)：{HGretem_value}')
                                    HGretem_value = []  # 用完HGretem_value列表后，要把列表清空，不然会叠加列表

                                    #   使用列表推导式过滤了列表 Hfetem_value 中值为 0 的元素，并将结果重新赋值给 Hfetem_value
                                    Hfetem_value = [x for x in Hfetem_value if x != 0]
                                    max_Hfetem = round(max(Hfetem_value), 1)
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = round(min(Hfetem_value), 1)
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                    # print(f'提纯器温度列表^^^^^^^^^^^^^^^6(℃)：{Hfetem_value}')
                                    Hfetem_value = []  # 用完Hfetem_value列表后，要把列表清空，不然会叠加列表

                                # 燃料耗率 / L.kWh - 1
                                if Once_Topgen != 0:
                                    Fuel_consumption = round((Once_RemFuelIn / Once_Topgen), 1)
                                else:
                                    Fuel_consumption = 0
                                everytime_Fuel_consumption.append(Fuel_consumption)
                                print(f'燃料消耗率列表 ：{Fuel_consumption}')

                            # 一天发一次电以上，执行下面程序
                            if len(count_end_datatime) > 1:

                                # 找出每次发电期间（内置水箱剩余燃料）fuel_List_value 的所有值
                                # 求出两个列表长度不同的部分。这段代码使用了 Python 中的切片操作。我们知道，对一个列表进行切片操作时，
                                # 可以指定起始位置和结束位置，如果只有一个位置（索引），则表示从那个位置到列表末尾。 在这里，fuel_levels[len(last_fuel_levels):] 表示从
                                # fuel_levels 列表中的索引 len(last_fuel_levels) 开始， 一直取到末尾，即取出 fuel_levels
                                # 计算液位。 last_fuel_levels 多出来的部分元素。
                                fuel_List_value = S_RemFuelIn_value[len(last_fuel_levels):]

                                # 计算电压重整室温度。每次发电期间 HGretem_List_value 重整室温度的值
                                HGretem_List_value = HGretem_value[len(last_HGretem_list):]
                                # 计算提纯器温度。每次发电期间 Hfetem_List_value 提纯器温度的值
                                Hfetem_List_value = Hfetem_value[len(last_Hfetem_list):]

                                # 找出每次发电期间（A 电堆电压）A_List_value 的所有值
                                A_List_value = A_StackV_value[len(last_A_List):]
                                # 找出每次发电期间（B 电堆电压）A_List_value 的所有值
                                B_List_value = B_StackV_value[len(last_B_List):]
                                # 找出每次发电期间（发电功率）last_power_value_list 的所有值
                                power_value_list = power_values[len(last_power_value_list):]
                                power_A_value_list = A_Power_values[len(last_A_power_value_list):]
                                power_B_value_list = B_Power_values[len(last_B_power_value_list):]
                                # 找出每次发电期间（母线电压）current_voltage_List_value 的所有值
                                current_voltage_List_value = current_voltage_value[
                                                             len(last_current_voltage_List_value):]
                                # print(f'母线电压平均值current_voltage_value-----------(W)：{current_voltage_value}')
                                # print(
                                #     f'母线电压平均值last_current_voltage_List_value-----------(W)：{last_current_voltage_List_value}')
                                # print(f'母线电压平均值-----------(W)：{current_voltage_List_value}')

                                current_voltage = round(
                                    sum(current_voltage_List_value) / len(current_voltage_List_value), 1)
                                everytime_current_voltage.append(current_voltage)
                                current_voltage_List_value.clear()
                                print(f'母线电压平均值(W)：{current_voltage}')

                                # print(f'总功率平均值>>>>>>>>>>>>：{power_value_list}')
                                calculate_A_power = round(calculate_average(power_A_value_list), 1)
                                everytime_A_power.append(calculate_A_power)
                                power_A_value_list.clear()
                                print(f'A堆功率平均值(W)：{calculate_A_power}')

                                calculate_B_power = round(calculate_average(power_B_value_list), 1)
                                everytime_B_power.append(calculate_B_power)
                                power_B_value_list.clear()
                                print(f'B堆功率平均值(W)：{calculate_B_power}')

                                # print(f'总功率平均值>>>>>>>>>>>>：{power_value_list}')
                                calculate_power = round(calculate_average(power_value_list), 1)
                                everytime_power.append(calculate_power)
                                power_value_list.clear()
                                print(f'总功率平均值(W)：{calculate_power}')

                                if S_RemFuelIn_value[0] > 0:
                                    # 计算燃料使用，计算列表中两两元素的差,大于等于0的部分存到一个新的列表中
                                    differences = [fuel_List_value[i] - fuel_List_value[i + 1] for i in
                                                   range(len(fuel_List_value) - 1)]
                                    positive_differences = [x for x in differences if x > 0]
                                    Once_RemFuelIn = round(sum(positive_differences), 2)
                                    if Once_RemFuelIn == 0:
                                        Once_RemFuelIn = 0.3
                                    Once_S_RemFuelIn.append(Once_RemFuelIn)
                                    print(f'每次发电消耗燃料（L）:{Once_RemFuelIn}')
                                else:
                                    differences = round(start_LiqlelM[-1] - end_LiqlelM[-1], 1)
                                    if differences < 0:
                                        differences = 0
                                    Once_S_RemFuelIn.append(differences)
                                    print(f'每次发电消耗燃料（mm）:{differences}')
                                    # print(f'液位(mm)******** ：{differences}')

                                # 计算发电过程中，A电堆电压平均值（过滤小于90和大于130的值）
                                average_A_StackV = round(calculate_filtered_average(A_List_value), 1)
                                everytime_A_StackV.append(average_A_StackV)
                                #### 2023.1.16新增
                                copy_everytime_A_StackV = copy.deepcopy(everytime_A_StackV)
                                copys_everytime_A_StackV.append(copy_everytime_A_StackV)
                                modified_A_StackV = [item[0] for item in copys_everytime_A_StackV]
                                everytime_A_StackV.clear()
                                ######
                                print(f'A电堆平均电压(V):{average_A_StackV}', end="        ")

                                # 计算发电过程中，B电堆电压平均值（过滤小于90和大于130的值）
                                average_B_StackV = round(calculate_filtered_average(B_List_value), 1)
                                everytime_B_StackV.append(average_B_StackV)
                                #### 2023.1.16新增
                                copy_everytime_B_StackV = copy.deepcopy(everytime_B_StackV)
                                copys_everytime_B_StackV.append(copy_everytime_B_StackV)
                                modified_B_StackV = [item[0] for item in copys_everytime_B_StackV]
                                everytime_B_StackV.clear()
                                ######
                                print(f'B电堆平均电压(V):{average_B_StackV}')

                                # print(f'重整室温度 HGretem_List_value ///////////// (℃) ：{HGretem_List_value}')
                                if all(item == 0 for item in HGretem_List_value) and all(
                                        item == 0 for item in Hfetem_List_value):

                                    max_HGretem = 0
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = 0
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')

                                    # print(f'重整室最列表温度^^^^^^^^^^^^5(℃)：{HGretem_value}')
                                    HGretem_value = []  # 用完HGretem_value列表后，要把列表清空，不然会叠加列表

                                    max_Hfetem = 0
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = 0
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                    # print(f'提纯器温度列表^^^^^^^^^^^^^^^6(℃)：{Hfetem_value}')
                                    Hfetem_value = []

                                else:
                                    # print(f'重整室温度列表(℃)>>>>>>>>>>>>>>>>>：{HGretem_List_value}\n')
                                    # print(f'提纯器温度列表(℃)>>>>>>>>>>>>>>>>>：{Hfetem_List_value}\n')

                                    #   使用列表推导式过滤了列表 HGretem_value 中值为 0 的元素，并将结果重新赋值给 HGretem_value
                                    HGretem_List_value = [x for x in HGretem_List_value if x != 0]
                                    max_HGretem = round(max(HGretem_List_value), 1)
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = round(min(HGretem_List_value), 1)
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')
                                    # print(f'重整室最温度列表 00000000  (℃)：{HGretem_List_value}')
                                    # print(f'重整室最小温度 HGretem_List_value |||||||||||  (℃)：{HGretem_List_value}')

                                    #   使用列表推导式过滤了列表 Hfetem_value 中值为 0 的元素，并将结果重新赋值给 Hfetem_value
                                    Hfetem_List_value = [x for x in Hfetem_List_value if x != 0]
                                    max_Hfetem = round(max(Hfetem_List_value), 1)
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = round(min(Hfetem_List_value), 1)
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                # 初始化,上一个的列表
                                last_fuel_levels.clear()
                                last_A_List.clear()
                                last_B_List.clear()
                                last_Hfetem_list.clear()
                                last_HGretem_list.clear()
                                last_power_value_list.clear()
                                last_A_power_value_list.clear()
                                last_B_power_value_list.clear()
                                last_current_voltage_List_value.clear()

                                # 在每次迭代结束后，将 fuel_levels 的值复制给 last_fuel_levels
                                # 使用 copy 模块中的 deepcopy 函数来创建一个深层副本，确保每个元素都是独立的
                                # 赋值，将当前列表的值赋于另一个列表，使另一个列表成为上一个列表的值

                                last_fuel_levels = copy.deepcopy(S_RemFuelIn_value)

                                last_A_List = copy.deepcopy(A_StackV_value)
                                last_B_List = copy.deepcopy(B_StackV_value)
                                last_HGretem_list = copy.deepcopy(HGretem_value)
                                last_Hfetem_list = copy.deepcopy(Hfetem_value)
                                last_power_value_list = copy.deepcopy(power_values)
                                last_A_power_value_list = copy.deepcopy(A_Power_values)
                                last_B_power_value_list = copy.deepcopy(B_Power_values)
                                last_current_voltage_List_value = copy.deepcopy(current_voltage_value)

                                # 燃料耗率 / L.kWh - 1
                                if Once_Topgen != 0:
                                    Fuel_consumption = round((Once_RemFuelIn / Once_Topgen), 1)
                                else:
                                    Fuel_consumption = 0
                                everytime_Fuel_consumption.append(Fuel_consumption)
                                print(f'燃料消耗率列表 ：{Fuel_consumption}')

                            print('=============     本次发电结束      ==================')

                        else:
                            Once_RemFuelIn = 0
                            if start_time is None and (index == len(df) - 1) == True and last_row[MSw] == True and len(
                                    count_end_datatime) == 1:  # 有开始发电时间并且到列的最后一行，把最后一行的数值添加进去
                                print(
                                    f"结束发电时间：{row[DateTime]}      "
                                    f"内置水箱剩余燃料: {round(row[S_RemFuelIn], 2)}    "
                                    f" 外置水箱剩余燃料: {round(row[S_RemFuelOut], 2)}    "
                                    f"内置水箱剩余燃料(mm): {round(row[LiqlelM], 1)} "
                                    f"外置水箱剩余燃料(mm): {round(row[LiqlelL], 1)}"
                                    f"总发电量:{row[Topgen]}    ")

                                print(len(count_end_datatime))  # 计算当天发电次数
                                Time_value.append(row[DateTime])
                                end_datatime.append(row[DateTime])
                                Topgen_value.append(row[Topgen])
                                # 创建列表用于储存输出到excel表格和数据

                                end_LiqlelL.append(round(row[LiqlelL], 1))
                                end_LiqlelM.append(round(row[LiqlelM], 1))

                                # 创建列表count_end_datatime，用于计数。一天发了多少次电

                                end_S_RemFuelIn.append(round(row[S_RemFuelIn], 1))
                                end_Topgen.append(round(row[Topgen], 1))
                                end_S_RemFuelOut.append(round(row[S_RemFuelOut], 1))

                                Once_Topgen = round(Topgen_value[-1] - Topgen_value[-2], 3)
                                if Once_Topgen < 0:
                                    Once_Topgen = 0
                                print(f"每次发电量(kw/h)：{Once_Topgen}")
                                Once_Topgen_value.append(Once_Topgen)

                                Stwtims.append(row[Stwtim])
                                print(f"发电次数：{row[Stwtim]}")

                                Time_diff = round(
                                    (pd.to_datetime(Time_value[-1]) - pd.to_datetime(
                                        Time_value[-2])).total_seconds() / 60,
                                    2)
                                Time_diffs.append(Time_diff)
                                print(f"每次发电时长(min)：{Time_diff}")

                                mean_IC = round(sum(IC_value) / len(IC_value), 2)
                                everytime_IC.append(mean_IC)
                                print(f'芯片平均温度(℃):{mean_IC}')

                                mean_A1_Stack_Temp = round(sum(A1_Stack_Temp_Value) / len(A1_Stack_Temp_Value), 2)
                                everytime_A1_Stack_Temp.append(mean_A1_Stack_Temp)
                                print(f'电堆A1平均温度(℃):{mean_A1_Stack_Temp}')

                                mean_A2_Stack_Temp = round(sum(A2_Stack_Temp_Value) / len(A2_Stack_Temp_Value), 2)
                                everytime_A2_Stack_Temp.append(mean_A2_Stack_Temp)
                                print(f'电堆A2平均温度(℃):{mean_A2_Stack_Temp}')

                                mean_B_Stack_Temp = round(sum(B_Stack_Temp_Value) / len(B_Stack_Temp_Value), 2)
                                everytime_B_Stack_Temp.append(mean_B_Stack_Temp)
                                print(f'电堆B平均温度(℃):{mean_B_Stack_Temp}')

                                # 计算液位。 last_fuel_levels 多出来的部分元素。
                                fuel_List_value = S_RemFuelIn_value[len(last_fuel_levels):]
                                # print(f'最后一次液位 >>>>>>>>>>>>：{fuel_List_value} \n')
                                # 计算电压重整室温度。每次发电期间 HGretem_List_value 重整室温度的值
                                HGretem_List_value = HGretem_value[len(last_HGretem_list):]
                                # print(f'最后一次重整室温度 >>>>>>>>>>>>：{HGretem_List_value} \n')
                                # 计算提纯器温度。每次发电期间 Hfetem_List_value 提纯器温度的值
                                Hfetem_List_value = Hfetem_value[len(last_Hfetem_list):]
                                # print(f'最后一次提纯器温度 >>>>>>>>>>>>：{Hfetem_List_value} \n')
                                # 找出每次发电期间（A 电堆电压）A_List_value 的所有值
                                A_List_value = A_StackV_value[len(last_A_List):]
                                # print(f'最后一次 A 电堆电压 >>>>>>>>>>>>：{A_List_value} \n')
                                # 找出每次发电期间（B 电堆电压）A_List_value 的所有值
                                B_List_value = B_StackV_value[len(last_B_List):]
                                # print(f'最后一次 B 电堆电压 >>>>>>>>>>>>：{B_List_value} \n')
                                # 找出每次发电期间（发电功率）last_power_value_list 的所有值
                                power_value_list = power_values[len(last_power_value_list):]
                                # print(f'总功率平均值>>>>>>>>>>>>：{power_value_list} \n')

                                power_A_value_list = A_Power_values[len(last_A_power_value_list):]
                                power_B_value_list = B_Power_values[len(last_B_power_value_list):]

                                current_voltage_List_value = current_voltage_value[
                                                             len(last_current_voltage_List_value):]

                                current_voltage = round(
                                    sum(current_voltage_List_value) / len(current_voltage_List_value), 1)
                                everytime_current_voltage.append(current_voltage)
                                current_voltage_List_value.clear()
                                print(f'母线电压平均值(W)：{current_voltage}')

                                # print(f'总功率平均值>>>>>>>>>>>>：{power_value_list}')
                                calculate_A_power = round(calculate_average(power_A_value_list), 1)
                                everytime_A_power.append(calculate_A_power)
                                power_A_value_list.clear()
                                print(f'A堆功率平均值(W)：{calculate_A_power}')

                                calculate_B_power = round(calculate_average(power_B_value_list), 1)
                                everytime_B_power.append(calculate_B_power)
                                power_B_value_list.clear()
                                print(f'B堆功率平均值(W)：{calculate_B_power}')

                                calculate_power = round(calculate_average(power_value_list), 1)
                                everytime_power.append(calculate_power)
                                power_value_list.clear()
                                print(f'总功率平均值(W)：{calculate_power}')

                                if S_RemFuelIn_value[0] > 0:
                                    # 计算燃料使用，计算列表中两两元素的差,大于等于0的部分存到一个新的列表中
                                    differences = [fuel_List_value[i] - fuel_List_value[i + 1] for i in
                                                   range(len(fuel_List_value) - 1)]
                                    positive_differences = [x for x in differences if x > 0]
                                    Once_RemFuelIn = round(sum(positive_differences), 2)
                                    if Once_RemFuelIn == 0:
                                        Once_RemFuelIn = 0.3
                                    Once_S_RemFuelIn.append(Once_RemFuelIn)
                                    print(f'每次发电消耗燃料（L）:{Once_RemFuelIn}')
                                else:
                                    differences = round(start_LiqlelM[-1] - end_LiqlelM[-1], 1)
                                    if differences < 0:
                                        differences = 0
                                    Once_S_RemFuelIn.append(differences)
                                    print(f'每次发电消耗燃料（mm）:{differences}')

                                # 计算发电过程中，A电堆电压平均值（过滤小于90和大于130的值）
                                average_A_StackV = round(calculate_filtered_average(A_List_value), 1)
                                everytime_A_StackV.append(average_A_StackV)
                                #### 2023.1.16新增
                                copy_everytime_A_StackV = copy.deepcopy(everytime_A_StackV)
                                copys_everytime_A_StackV.append(copy_everytime_A_StackV)
                                modified_A_StackV = [item[0] for item in copys_everytime_A_StackV]
                                everytime_A_StackV.clear()
                                ######
                                print(f'A电堆平均电压(V):{average_A_StackV}', end="        ")

                                # 计算发电过程中，B电堆电压平均值（过滤小于90和大于130的值）
                                average_B_StackV = round(calculate_filtered_average(B_List_value), 1)
                                everytime_B_StackV.append(average_B_StackV)
                                #### 2023.1.16新增
                                copy_everytime_B_StackV = copy.deepcopy(everytime_B_StackV)
                                copys_everytime_B_StackV.append(copy_everytime_B_StackV)
                                modified_B_StackV = [item[0] for item in copys_everytime_B_StackV]
                                everytime_B_StackV.clear()
                                ######
                                print(f'B电堆平均电压(V):{average_B_StackV}')

                                # print(f'重整室温度 HGretem_List_value ///////////// (℃) ：{HGretem_List_value}')
                                if all(item == 0 for item in HGretem_List_value) and all(
                                        item == 0 for item in Hfetem_List_value):

                                    max_HGretem = 0
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = 0
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')

                                    # print(f'重整室最列表温度^^^^^^^^^^^^5(℃)：{HGretem_value}')
                                    HGretem_value = []  # 用完HGretem_value列表后，要把列表清空，不然会叠加列表

                                    max_Hfetem = 0
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = 0
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                    # print(f'提纯器温度列表^^^^^^^^^^^^^^^6(℃)：{Hfetem_value}')
                                    Hfetem_value = []

                                else:
                                    # print(f'重整室温度列表(℃)>>>>>>>>>>>>>>>>>：{HGretem_List_value}\n')
                                    # print(f'提纯器温度列表(℃)>>>>>>>>>>>>>>>>>：{Hfetem_List_value}\n')

                                    #   使用列表推导式过滤了列表 HGretem_value 中值为 0 的元素，并将结果重新赋值给 HGretem_value
                                    HGretem_List_value = [x for x in HGretem_List_value if x != 0]
                                    max_HGretem = round(max(HGretem_List_value), 1)
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = round(min(HGretem_List_value), 1)
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')
                                    # print(f'重整室最温度列表 00000000  (℃)：{HGretem_List_value}')
                                    # print(f'重整室最小温度 HGretem_List_value |||||||||||  (℃)：{HGretem_List_value}')

                                    #   使用列表推导式过滤了列表 Hfetem_value 中值为 0 的元素，并将结果重新赋值给 Hfetem_value
                                    Hfetem_List_value = [x for x in Hfetem_List_value if x != 0]
                                    max_Hfetem = round(max(Hfetem_List_value), 1)
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = round(min(Hfetem_List_value), 1)
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                # 燃料耗率 / L.kWh - 1
                                if Once_Topgen != 0:
                                    Fuel_consumption = round((Once_RemFuelIn / Once_Topgen), 1)
                                else:
                                    Fuel_consumption = 0
                                everytime_Fuel_consumption.append(Fuel_consumption)
                                print(f'燃料消耗率 ：{Fuel_consumption}')

                                # 初始化,上一个的列表
                                last_fuel_levels.clear()
                                last_A_List.clear()
                                last_B_List.clear()
                                last_Hfetem_list.clear()
                                last_HGretem_list.clear()
                                last_power_value_list.clear()
                                last_A_power_value_list.clear()
                                last_B_power_value_list.clear()
                                last_current_voltage_List_value.clear()

                                # 在每次迭代结束后，将 fuel_levels 的值复制给 last_fuel_levels
                                # 使用 copy 模块中的 deepcopy 函数来创建一个深层副本，确保每个元素都是独立的
                                # 赋值，将当前列表的值赋于另一个列表，使另一个列表成为上一个列表的值
                                last_fuel_levels = copy.deepcopy(S_RemFuelIn_value)
                                last_A_List = copy.deepcopy(A_StackV_value)
                                last_B_List = copy.deepcopy(B_StackV_value)
                                last_HGretem_list = copy.deepcopy(HGretem_value)
                                last_Hfetem_list = copy.deepcopy(Hfetem_value)
                                last_power_value_list = copy.deepcopy(power_values)
                                last_A_power_value_list = copy.deepcopy(A_Power_values)
                                last_B_power_value_list = copy.deepcopy(B_Power_values)
                                last_current_voltage_List_value = copy.deepcopy(current_voltage_value)

                            if start_time is None and (index == len(df) - 1) == True and last_row[MSw] == True and len(
                                    count_end_datatime) > 1:
                                print(
                                    f"结束发电时间：{last_row[DateTime]}      "
                                    f"内置水箱剩余燃料: {round(last_row[S_RemFuelIn], 2)}     "
                                    f"外置水箱剩余燃料: {round(last_row[S_RemFuelOut], 2)}    "
                                    f"内置水箱剩余燃料(mm): {round(prev_row[LiqlelM], 1)} "
                                    f"外置水箱剩余燃料(mm): {round(prev_row[LiqlelL], 1)}"
                                    f"总发电量:{last_row[Topgen]}    ")

                                print(len(count_end_datatime))  # 计算当天发电次数
                                Time_value.append(last_row[DateTime])
                                end_datatime.append(last_row[DateTime])
                                Topgen_value.append(last_row[Topgen])
                                # 创建列表用于储存输出到excel表格和数据

                                end_LiqlelL.append(round(prev_row[LiqlelL], 1))
                                end_LiqlelM.append(round(prev_row[LiqlelM], 1))

                                # 创建列表count_end_datatime，用于计数。一天发了多少次电

                                end_S_RemFuelIn.append(round(last_row[S_RemFuelIn], 1))
                                end_Topgen.append(round(last_row[Topgen], 1))
                                end_S_RemFuelOut.append(round(last_row[S_RemFuelOut], 1))

                                Once_Topgen = round(Topgen_value[-1] - Topgen_value[-2], 3)
                                if Once_Topgen < 0:
                                    Once_Topgen = 0
                                print(f"每次发电量(kw/h)：{Once_Topgen}")
                                Once_Topgen_value.append(Once_Topgen)

                                Stwtims.append(last_row[Stwtim])
                                print(f"发电次数：{row[Stwtim]}")

                                Time_diff = round(
                                    (pd.to_datetime(Time_value[-1]) - pd.to_datetime(
                                        Time_value[-2])).total_seconds() / 60,
                                    2)
                                Time_diffs.append(Time_diff)
                                print(f"每次发电时长(min)：{Time_diff}")

                                mean_IC = round(sum(IC_value) / len(IC_value), 2)
                                everytime_IC.append(mean_IC)
                                print(f'芯片平均温度(℃):{mean_IC}')

                                mean_A1_Stack_Temp = round(sum(A1_Stack_Temp_Value) / len(A1_Stack_Temp_Value), 2)
                                everytime_A1_Stack_Temp.append(mean_A1_Stack_Temp)
                                print(f'电堆A1平均温度(℃):{mean_A1_Stack_Temp}')

                                mean_A2_Stack_Temp = round(sum(A2_Stack_Temp_Value) / len(A2_Stack_Temp_Value), 2)
                                everytime_A2_Stack_Temp.append(mean_A2_Stack_Temp)
                                print(f'电堆A2平均温度(℃):{mean_A2_Stack_Temp}')

                                mean_B_Stack_Temp = round(sum(B_Stack_Temp_Value) / len(B_Stack_Temp_Value), 2)
                                everytime_B_Stack_Temp.append(mean_B_Stack_Temp)
                                print(f'电堆B平均温度(℃):{mean_B_Stack_Temp}')

                                # 计算液位。 last_fuel_levels 多出来的部分元素。
                                fuel_List_value = S_RemFuelIn_value[len(last_fuel_levels):]
                                # print(f'最后一次液位 >>>>>>>>>>>>：{fuel_List_value} \n')
                                # 计算电压重整室温度。每次发电期间 HGretem_List_value 重整室温度的值
                                HGretem_List_value = HGretem_value[len(last_HGretem_list):]
                                # print(f'最后一次重整室温度 >>>>>>>>>>>>：{HGretem_List_value} \n')
                                # 计算提纯器温度。每次发电期间 Hfetem_List_value 提纯器温度的值
                                Hfetem_List_value = Hfetem_value[len(last_Hfetem_list):]
                                # print(f'最后一次提纯器温度 >>>>>>>>>>>>：{Hfetem_List_value} \n')
                                # 找出每次发电期间（A 电堆电压）A_List_value 的所有值
                                A_List_value = A_StackV_value[len(last_A_List):]
                                # print(f'最后一次 A 电堆电压 >>>>>>>>>>>>：{A_List_value} \n')
                                # 找出每次发电期间（B 电堆电压）A_List_value 的所有值
                                B_List_value = B_StackV_value[len(last_B_List):]
                                # print(f'最后一次 B 电堆电压 >>>>>>>>>>>>：{B_List_value} \n')
                                # 找出每次发电期间（发电功率）last_power_value_list 的所有值
                                power_value_list = power_values[len(last_power_value_list):]
                                # print(f'总功率平均值>>>>>>>>>>>>：{power_value_list} \n')

                                power_A_value_list = A_Power_values[len(last_A_power_value_list):]
                                power_B_value_list = B_Power_values[len(last_B_power_value_list):]

                                current_voltage_List_value = current_voltage_value[
                                                             len(last_current_voltage_List_value):]

                                current_voltage = round(
                                    sum(current_voltage_List_value) / len(current_voltage_List_value), 1)
                                everytime_current_voltage.append(current_voltage)
                                current_voltage_List_value.clear()
                                print(f'母线电压平均值(W)：{current_voltage}')

                                # print(f'总功率平均值>>>>>>>>>>>>：{power_value_list}')
                                calculate_A_power = round(calculate_average(power_A_value_list), 1)
                                everytime_A_power.append(calculate_A_power)
                                power_A_value_list.clear()
                                print(f'A堆功率平均值(W)：{calculate_A_power}')

                                calculate_B_power = round(calculate_average(power_B_value_list), 1)
                                everytime_B_power.append(calculate_B_power)
                                power_B_value_list.clear()
                                print(f'B堆功率平均值(W)：{calculate_B_power}')

                                calculate_power = round(calculate_average(power_value_list), 1)
                                everytime_power.append(calculate_power)
                                power_value_list.clear()
                                print(f'总功率平均值(W)：{calculate_power}')

                                if S_RemFuelIn_value[0] > 0:
                                    # 计算燃料使用，计算列表中两两元素的差,大于等于0的部分存到一个新的列表中
                                    differences = [fuel_List_value[i] - fuel_List_value[i + 1] for i in
                                                   range(len(fuel_List_value) - 1)]
                                    positive_differences = [x for x in differences if x > 0]
                                    Once_RemFuelIn = round(sum(positive_differences), 2)
                                    if Once_RemFuelIn == 0:
                                        Once_RemFuelIn = 0.3
                                    Once_S_RemFuelIn.append(Once_RemFuelIn)
                                    print(f'每次发电消耗燃料（L）:{Once_RemFuelIn}')
                                else:
                                    differences = round(start_LiqlelM[-1] - end_LiqlelM[-1], 1)
                                    if differences < 0:
                                        differences = 0
                                    Once_S_RemFuelIn.append(differences)
                                    print(f'每次发电消耗燃料（mm）:{differences}')
                                    # print(f'液位(mm)******** ：{differences}')

                                # 计算发电过程中，A电堆电压平均值（过滤小于90和大于130的值）
                                average_A_StackV = round(calculate_filtered_average(A_List_value), 1)
                                everytime_A_StackV.append(average_A_StackV)
                                #### 2023.1.16新增
                                copy_everytime_A_StackV = copy.deepcopy(everytime_A_StackV)
                                copys_everytime_A_StackV.append(copy_everytime_A_StackV)
                                modified_A_StackV = [item[0] for item in copys_everytime_A_StackV]
                                everytime_A_StackV.clear()
                                ######
                                print(f'A电堆平均电压(V):{average_A_StackV}', end="        ")

                                # 计算发电过程中，B电堆电压平均值（过滤小于90和大于130的值）
                                average_B_StackV = round(calculate_filtered_average(B_List_value), 1)
                                everytime_B_StackV.append(average_B_StackV)
                                #### 2023.1.16新增
                                copy_everytime_B_StackV = copy.deepcopy(everytime_B_StackV)
                                copys_everytime_B_StackV.append(copy_everytime_B_StackV)
                                modified_B_StackV = [item[0] for item in copys_everytime_B_StackV]
                                everytime_B_StackV.clear()
                                ######
                                print(f'B电堆平均电压(V):{average_B_StackV}')

                                # print(f'重整室温度 HGretem_List_value ///////////// (℃) ：{HGretem_List_value}')
                                if all(item == 0 for item in HGretem_List_value) and all(
                                        item == 0 for item in Hfetem_List_value):

                                    max_HGretem = 0
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = 0
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')

                                    # print(f'重整室最列表温度^^^^^^^^^^^^5(℃)：{HGretem_value}')
                                    HGretem_value = []  # 用完HGretem_value列表后，要把列表清空，不然会叠加列表

                                    max_Hfetem = 0
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = 0
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                    # print(f'提纯器温度列表^^^^^^^^^^^^^^^6(℃)：{Hfetem_value}')
                                    Hfetem_value = []

                                else:
                                    # print(f'重整室温度列表(℃)>>>>>>>>>>>>>>>>>：{HGretem_List_value}\n')
                                    # print(f'提纯器温度列表(℃)>>>>>>>>>>>>>>>>>：{Hfetem_List_value}\n')

                                    #   使用列表推导式过滤了列表 HGretem_value 中值为 0 的元素，并将结果重新赋值给 HGretem_value
                                    HGretem_List_value = [x for x in HGretem_List_value if x != 0]
                                    max_HGretem = round(max(HGretem_List_value), 1)
                                    everytime_max_HGretem.append(max_HGretem)
                                    print(f'重整室最大温度(℃)：{max_HGretem}', end="      ")

                                    min_HGretem = round(min(HGretem_List_value), 1)
                                    everytime_min_HGretem.append(min_HGretem)
                                    print(f'重整室最小温度(℃)：{min_HGretem}')
                                    # print(f'重整室最温度列表 00000000  (℃)：{HGretem_List_value}')
                                    # print(f'重整室最小温度 HGretem_List_value |||||||||||  (℃)：{HGretem_List_value}')

                                    #   使用列表推导式过滤了列表 Hfetem_value 中值为 0 的元素，并将结果重新赋值给 Hfetem_value
                                    Hfetem_List_value = [x for x in Hfetem_List_value if x != 0]
                                    max_Hfetem = round(max(Hfetem_List_value), 1)
                                    everytime_max_Hfetem.append(max_Hfetem)
                                    print(f'提纯器最大温度(℃)：{max_Hfetem}', end="      ")

                                    min_Hfetem = round(min(Hfetem_List_value), 1)
                                    everytime_min_Hfetem.append(min_Hfetem)
                                    print(f'提纯器最小温度(℃)：{min_Hfetem}')

                                # 燃料耗率 / L.kWh - 1
                                if Once_Topgen != 0:
                                    Fuel_consumption = round((Once_RemFuelIn / Once_Topgen), 1)
                                else:
                                    Fuel_consumption = 0
                                everytime_Fuel_consumption.append(Fuel_consumption)
                                print(f'燃料消耗率 ：{Fuel_consumption}')

                                # 初始化,上一个的列表
                                last_fuel_levels.clear()
                                last_A_List.clear()
                                last_B_List.clear()
                                last_Hfetem_list.clear()
                                last_HGretem_list.clear()
                                last_power_value_list.clear()
                                last_A_power_value_list.clear()
                                last_B_power_value_list.clear()
                                last_current_voltage_List_value.clear()

                                # 在每次迭代结束后，将 fuel_levels 的值复制给 last_fuel_levels
                                # 使用 copy 模块中的 deepcopy 函数来创建一个深层副本，确保每个元素都是独立的
                                # 赋值，将当前列表的值赋于另一个列表，使另一个列表成为上一个列表的值
                                last_fuel_levels = copy.deepcopy(S_RemFuelIn_value)
                                last_A_List = copy.deepcopy(A_StackV_value)
                                last_B_List = copy.deepcopy(B_StackV_value)
                                last_HGretem_list = copy.deepcopy(HGretem_value)
                                last_Hfetem_list = copy.deepcopy(Hfetem_value)
                                last_power_value_list = copy.deepcopy(power_values)
                                last_A_power_value_list = copy.deepcopy(A_Power_values)
                                last_B_power_value_list = copy.deepcopy(B_Power_values)
                                last_current_voltage_List_value = copy.deepcopy(current_voltage_value)

                            start_time = None
                    prev_row = row

                Sum_Topgen = round(sum(Once_Topgen_value), 2)
                Sum_S_RemFuelIn = sum(Once_S_RemFuelIn)
                Sum_Time_min = round(sum(Time_diffs), 2)

                print(f"总发电量(kw/h)：{Sum_Topgen}")
                print(f"总发电时间(min.s)：{Sum_Time_min}")

                if start_S_RemFuelIn[0] > 0:
                    print(f"总燃料消耗(L)：{Sum_S_RemFuelIn}")
                else:
                    print(f"总燃料消耗(mm)：{Sum_S_RemFuelIn}")

                # 计数清零，用于计算有多少个【'结束发电时间': end_datatime】。来判断一天里面发了多少次电
                count_end_datatime.clear()
                S_RemFuelIn_value.clear()
                A_StackV_value.clear()
                B_StackV_value.clear()
                current_voltage_value.clear()

                A_Power_values = []
                B_Power_values = []
                power_values = []
                HGretem_value = []
                Hfetem_value = []
                last_HGretem_list = []  # 确保在每次循环开始时重置为空列表
                last_Hfetem_list = []  # 确保在每次循环开始时重置为空列表
                last_B_List = []
                last_A_List = []
                last_fuel_levels = []
                last_A_power_value_list = []
                last_B_power_value_list = []
                last_power_value_list = []
                last_current_voltage_List_value = []
                count_datatime = []
                start_time = None
                second_start_time = None
                second_row = None
                first_start_datatime = 0
                second_end_datatime = 0

                A1_Stack_Temp_Value.clear()
                A2_Stack_Temp_Value.clear()
                B_Stack_Temp_Value.clear()
                IC_value.clear()

                print(f"\n开始发电时间 长度：{len(start_datatime)}")
                print(f"结束发电时间 长度：{len(end_datatime)}")
                print(f"开始外置水箱剩余燃料 长度：{len(start_S_RemFuelOut)}")
                print(f"结束外置水箱剩余燃料 长度：{len(end_S_RemFuelOut)}")
                print(f"开始内置水箱剩余燃料 长度：{len(start_S_RemFuelIn)}")
                print(f"结束内置水箱剩余燃料 长度：{len(end_S_RemFuelIn)}")
                print(f"开始总发电量 长度：{len(start_Topgen)}")
                print(f"结束总发电量 长度：{len(end_Topgen)}")
                print(f"发电功率 长度：{len(everytime_power)}")
                print(f"芯片温度 长度：{len(everytime_IC)}")
                print(f"A电堆电压 长度：{len(modified_A_StackV)}")
                print(f"B电堆电压 长度：{len(modified_B_StackV)}")
                print(f"重整室最高温度 长度：{len(everytime_max_HGretem)}")
                print(f"重整室最低温度 长度：{len(everytime_min_HGretem)}")
                print(f"提纯器最高温度 长度：{len(everytime_max_Hfetem)}")
                print(f"提纯器最低温度 长度：{len(everytime_min_Hfetem)}")
                print(f"发电运行时间 长度：{len(Time_diffs)}")
                print(f"消耗燃料 长度：{len(Once_S_RemFuelIn)}")
                print(f"发电量 长度：{len(Once_Topgen_value)}")
                print(f"发电次数 长度：{len(Stwtims)}")
                print(f"燃料消耗率 长度：{len(everytime_Fuel_consumption)}\n")
                print(f"母线电压 长度：{len(everytime_current_voltage)}\n")

                print(f"开始外置水箱剩余燃料(mm) 长度：{len(start_LiqlelL)}")
                print(f"结束外置水箱剩余燃料(mm) 长度：{len(end_LiqlelL)}")
                print(f"开始内置水箱剩余燃料(mm) 长度：{len(start_LiqlelM)}")
                print(f"结束内置水箱剩余燃料(mm) 长度：{len(end_LiqlelM)}")

                print(f"电堆A1温度 长度：{len(everytime_A1_Stack_Temp)}")
                print(f"电堆A2温度 长度：{len(everytime_A2_Stack_Temp)}")
                print(f"电堆B温度 长度：{len(everytime_B_Stack_Temp)}")

                print(f'\n++++++++++++++   一天的计算结束   ++++++++++++++++++++++++\n')

            else:
                print(f'\n++++++++++++++   {a1}    当天没有发电     ++++++++++++++++++++++++\n')

            # # 当你完成对Excel文件的操作后，应该关闭文件以释放资源。使用ExcelFile对象的close方法来实现这一点。
            # xl.close()

        except FileNotFoundError:
            print(f"文件 {adress1} 不存在，已跳过")



    else:
        print(f"文件 {adress1} 不存在，已跳过")

# 在控制台上打印，显示每列的长度(元素个数) ，如果长度(元素个数)不一样，会报错“输出的列长不一样”
print(f"开始发电时间 长度：{len(start_datatime)}")
print(f"结束发电时间 长度：{len(end_datatime)}")
print(f"开始外置水箱剩余燃料(L) 长度：{len(start_S_RemFuelOut)}")
print(f"结束外置水箱剩余燃料(L) 长度：{len(end_S_RemFuelOut)}")
print(f"开始内置水箱剩余燃料(L) 长度：{len(start_S_RemFuelIn)}")
print(f"结束内置水箱剩余燃料(L) 长度：{len(end_S_RemFuelIn)}")
print(f"开始总发电量 长度：{len(start_Topgen)}")
print(f"结束总发电量 长度：{len(end_Topgen)}")
print(f"总发电功率 长度：{len(everytime_power)}")
print(f"A电堆功率 长度：{len(everytime_A_power)}")
print(f"B电堆功率 长度：{len(everytime_B_power)}")
print(f"芯片温度 长度：{len(everytime_IC)}")
print(f"A电堆电压 长度：{len(modified_A_StackV)}")
print(f"B电堆电压 长度：{len(modified_B_StackV)}")
print(f"重整室最高温度 长度：{len(everytime_max_HGretem)}")
print(f"重整室最低温度 长度：{len(everytime_min_HGretem)}")
print(f"提纯器最高温度 长度：{len(everytime_max_Hfetem)}")
print(f"提纯器最低温度 长度：{len(everytime_min_Hfetem)}")
print(f"发电运行时间 长度：{len(Time_diffs)}")
print(f"消耗燃料 长度：{len(Once_S_RemFuelIn)}")
print(f"发电量 长度：{len(Once_Topgen_value)}")
print(f"发电次数 长度：{len(Stwtims)}")
print(f"燃料消耗率 长度：{len(everytime_Fuel_consumption)}")
print(f"母线电压 长度：{len(everytime_current_voltage)}\n")

print(f"开始外置水箱剩余燃料(mm) 长度：{len(start_LiqlelL)}")
print(f"结束外置水箱剩余燃料(mm) 长度：{len(end_LiqlelL)}")
print(f"开始内置水箱剩余燃料(mm) 长度：{len(start_LiqlelM)}")
print(f"结束内置水箱剩余燃料(mm) 长度：{len(end_LiqlelM)}")

print(f"电堆A1温度 长度：{len(everytime_A1_Stack_Temp)}")
print(f"电堆A2温度 长度：{len(everytime_A2_Stack_Temp)}")
print(f"电堆B温度 长度：{len(everytime_B_Stack_Temp)}")

count = 0
if any(value > 0 for value in start_S_RemFuelIn) and any(value > 0 for value in end_S_RemFuelIn):
    count = 1
    # 以下'消耗燃料'保存格式为 L
    # 将新的DataFrame保存到新的Excel文件中
    new_df = pd.DataFrame(
        {
            '开始发电时间': start_datatime,
            '结束发电时间': end_datatime,

            '开始外置水箱剩余燃料(mm)': start_LiqlelL,
            '结束外置水箱剩余燃料(mm)': end_LiqlelL,
            '开始内置水箱剩余燃料(mm)': start_LiqlelM,
            '结束内置水箱剩余燃料(mm)': end_LiqlelM,

            '开始外置水箱剩余燃料(L)': start_S_RemFuelOut,
            '结束外置水箱剩余燃料(L)': end_S_RemFuelOut,
            '开始内置水箱剩余燃料(L)': start_S_RemFuelIn,
            '结束内置水箱剩余燃料(L)': end_S_RemFuelIn,
            '开始总发电量(kw/h)': start_Topgen,
            '结束总发电量(kw/h)': end_Topgen,
            '母线电压(V)': everytime_current_voltage,
            '总发电功率(W)': everytime_power,
            'A电堆功率(W)': everytime_A_power,
            'B电堆功率(W)': everytime_B_power,
            '芯片温度(℃)': everytime_IC,

            '电堆A1温度(℃)': everytime_A1_Stack_Temp,
            '电堆A2温度(℃)': everytime_A2_Stack_Temp,
            '电堆B温度(℃)': everytime_B_Stack_Temp,

            'A电堆电压(V)': modified_A_StackV,
            'B电堆电压(V)': modified_B_StackV,
            '重整室最高温度(℃)': everytime_max_HGretem,
            '重整室最低温度(℃)': everytime_min_HGretem,
            '提纯器最高温度(℃)': everytime_max_Hfetem,
            '提纯器最低温度(℃)': everytime_min_Hfetem,
            '发电运行时间(min.s)': Time_diffs,
            '消耗燃料(L)': Once_S_RemFuelIn,
            '发电量(kw/h)': Once_Topgen_value,
            '发电次数': Stwtims,
            '燃料消耗率(L.kWh -1)': everytime_Fuel_consumption

        })

else:
    count = 2
    # 以下'消耗燃料'保存格式为 mm
    # 将新的DataFrame保存到新的Excel文件中
    new_df = pd.DataFrame(
        {
            '开始发电时间': start_datatime,
            '结束发电时间': end_datatime,

            '开始外置水箱剩余燃料(mm)': start_LiqlelL,
            '结束外置水箱剩余燃料(mm)': end_LiqlelL,
            '开始内置水箱剩余燃料(mm)': start_LiqlelM,
            '结束内置水箱剩余燃料(mm)': end_LiqlelM,

            '开始外置水箱剩余燃料(L)': start_S_RemFuelOut,
            '结束外置水箱剩余燃料(L)': end_S_RemFuelOut,
            '开始内置水箱剩余燃料(L)': start_S_RemFuelIn,
            '结束内置水箱剩余燃料(L)': end_S_RemFuelIn,
            '开始总发电量(kw/h)': start_Topgen,
            '结束总发电量(kw/h)': end_Topgen,
            '母线电压(V)': everytime_current_voltage,
            '总发电功率(W)': everytime_power,
            'A电堆功率(W)': everytime_A_power,
            'B电堆功率(W)': everytime_B_power,
            '芯片温度(℃)': everytime_IC,

            '电堆A1温度(℃)': everytime_A1_Stack_Temp,
            '电堆A2温度(℃)': everytime_A2_Stack_Temp,
            '电堆B温度(℃)': everytime_B_Stack_Temp,

            'A电堆电压(V)': modified_A_StackV,
            'B电堆电压(V)': modified_B_StackV,
            '重整室最高温度(℃)': everytime_max_HGretem,
            '重整室最低温度(℃)': everytime_min_HGretem,
            '提纯器最高温度(℃)': everytime_max_Hfetem,
            '提纯器最低温度(℃)': everytime_min_Hfetem,
            '发电运行时间(min.s)': Time_diffs,
            '消耗燃料(mm)': Once_S_RemFuelIn,
            '发电量(kw/h)': Once_Topgen_value,
            '发电次数': Stwtims,
            '燃料消耗率(L.kWh -1)': everytime_Fuel_consumption

        })

file_path = adress3
new_df.to_excel(file_path, index=False)
# 打开现有的Excel文件
workbook = openpyxl.load_workbook(file_path)
# 选择第一个工作表
sheet = workbook.active
# 设置第一行的行高
sheet.row_dimensions[1].height = 50
# 设置第一列和第二列的宽度为 25
sheet.column_dimensions['A'].width = 23  # 第一列
sheet.column_dimensions['B'].width = 23  # 第二列
# 设置其余列的宽度为 10
for col in sheet.columns:
    if col[0].column_letter not in ['A', 'B']:
        sheet.column_dimensions[col[0].column_letter].width = 12
# 遍历第一行的所有单元格，并为每个单元格对象同时设置自动换行、水平居中和垂直居中。
for cell in sheet[1]:
    cell_obj = cell
    cell_obj.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')

if any(value > 0 for value in everytime_power):
    workbook.save(file_path)
    print(f"\n文件保存成功 ！! ! ")
    print(f"文件保存路径 ：{file_path}")
else:
    print(f"\n文件保存不成功 ！! ! ")
    print(f"\n所读取的文件里面没有发电数据  ！! ! ")
