import pandas as pd
import copy
import openpyxl
from datetime import datetime
import numpy as np
import os

# ================================================= #
# 版本更新：2024.3.27
# 更新内容：
# 1.新增‘平均产氢间隔时间’   2.新增’备注‘
# 3.修改液位选择，如果以"升"为单位的液位为空，则赋于默认值0.计算以“毫米”为单位的液位，“白石”“楼下机房”内置液位S_RemFuelIn，S_RemFuelOut等为空

#  2024_3_29 版本更新：   更新时间 2024.3.29
# 1.新增‘待机条件’判断：如果电堆电压‘StaV’全部为0，则为待机待机状态，没有发电。（原本条件：整机开关’MSw‘全部都是False，则为待机状态）
# 2.新增备注条件：如果数据量（总行数）小于3500，则给备注加上注释。数据量（总行数）小于多少。因为数据太少，算出来的值不准确
# ================================================= #

# 打印行号和列的数据
differences = []

fuel_levels = []
last_fuel_levels = []
S_RemFuelIn_value = []
positive_differences = []
calculate_positive_differences = []
Once_S_RemFuelIn = []
start_datatime = []
end_datatime = []

start_Topgen = []
end_Topgen = []

No_S_RemFuelIn_value = []  # 不发电时，储存内置液位的值
# NO_differences = []
NO_positive_differences = []
NO_Once_S_RemFuelIn = []
New_MSW = []
all_Sum_S_RemFuelIn = []
Timer_RemFuelIn = []

start_S_RemFuelIn = []
end_S_RemFuelIn = []

start_S_RemFuelOut = []
end_S_RemFuelOut = []
No_S_RemFuelOut_value = []

No_LiqlelL = []
No_LiqlelM = []

start_No_LiqlelL = []
end_No_LiqlelL = []

start_No_LiqlelM = []
end_No_LiqlelM = []

NO_DateTime = []
One_DateTime = []

No_HGHpre = []
No_HGHpre_Count = []
No_HGHpre_SumCount = []

q = []
df = []
df_list = []

No_HGHpre_time_list = []
No_HGHpre_time_average = []  # 平均产氢时间

remark = []  # 备注
New_StaV=[]     #电堆电压列表


b1 = '2024_1月白石待机燃料消耗数据，备注小于3500'  # 储存 EXCEL表格 的文件名称202
# adress2 = 'C:/Users/FCK/Desktop/12/test/%s.xlsx' % b1
adress3 = f"C:/Users/FCK/Desktop/12/数据处理/{b1}.xlsx"  # 储存 EXCEL表格文件 的路径
#  EXCEL格式为“某某 年，某某 月，某某 日” ，例如：”2023.10.1“这种格式.。"  年 . 月  . 日  "
year = 2024  # 年，表格的年
month = 1  # 月，表格的月

for i in range(1, 33):  # 遍历所有数据  i=8  range=31.   取值范围：8<= i <31
    # a1 = '2023.9.%s' % i
    # b1 = '2023_11_%s_test数据' %i
    a1 = '%d.%d.%d' % (year, month, i)  # 这个指令将会使用 year、month 和 i 的值来创建一个类似于 "XXXX.XX.XX" 格式的字符串，并将其存储在变量 a1 中。
    a1 = a1.strip()  # 这个指令会将变量 a1 中的字符串去掉开头和结尾的空白字符
    # 读取Excel文件中的数据
    adress1 = f'C:/Users/FCK/Desktop/12/白石10/1月/{a1}.xlsx'  # 读取 EXCEL表格文件 的路径

    if os.path.exists(adress1):  # 检查文件（文件名，文件路径是对得上）是否存在，不存在则结束程序
        try:
            # 在这里进行对数据的处理和分析
            # xl = pd.read_excel(adress1)
            xl = pd.ExcelFile(adress1)  # 使用 pd.ExcelFile() 方法打开 Excel 文件
            for sheet_name in xl.sheet_names:  # 遍历文件中的所有 sheet
                one_sheet = xl.parse(sheet_name)  # 读取当前 sheet 的数据
                df_list.append(one_sheet)  # 将读取的数据合并到 all_data 中
            # 使用 pd.concat() 方法将所有数据框连接成一个
            df = pd.concat(df_list, ignore_index=True)
            # 现在 all_data 包含了所有 sheet 的数据

            # 选择要读取的列名
            MSw = 'MSw'  # 开关状态
            DateTime = 'DateTime'  # 时间
            S_RemFuelIn = 'S_RemFuelIn'  # 内置水箱液位
            S_RemFuelOut = 'S_RemFuelOut'  # 外置水箱液位
            HGHpre = 'HGHpre'  # 氢气压力

            New_MSW = df['MSw'].tolist()
            max_index = df.index.max()

            New_StaV = df['StaV'].tolist()

            print(f'索引最大值：{max_index}')
            # prev_row = None

            LiqlelL = 'LiqlelL'  # 外置液位（mm）
            LiqlelM = 'LiqlelM'  # 内置液位（mm）

            #   如果电压小于85，则跳过当天计算
            if any(df['StaV'] >= 0):
                # second_row = df.iloc[1]  # 这行代码将DataFrame中的第二行数据存储在变量second_row中，以便后续对第二行数据进行操作和分析
                # last_row = df.iloc[-1]  # 这行代码将DataFrame中的最后一行数据存储在变量last_row中，以便后续对最后一行数据进行操作和分析

                NO_DateTime = df['DateTime'].tolist()
                # 分割时间为{年-月-日  ， 时-分-秒}
                date_only = NO_DateTime[1].split(" ")
                # 日期 ：年-月-日

                print(f'\n ————————————————  {date_only[0]}   一天计算开始    ————————————————    \n')

                # 获取 'MSw' 列的所有数据，并存储到列表 New_MSW 中

                # 使用 all() 函数检查 'MSw' 列中的所有值是否都为 False
                # 如果MSW=FALSE，不发电时，储存发电时间段内某列的数据，或者 电堆电压StaV全部等于0时，
                if all(value == False for value in New_MSW) or all(value == 0 for value in New_StaV):
                    for index, row in df.iterrows():  # 这段代码会遍历 DataFrame df 中的每一行数据。

                        # print("row[S_RemFuelIn] 值类型:", type(row[S_RemFuelIn]), f' 值 = {row[S_RemFuelIn]}')

                        # 检查 row[S_RemFuelIn] 是否为 NaN 的数据类型为float
                        if np.isnan(row[S_RemFuelIn]):
                            No_S_RemFuelIn_value.append(0)  # 不发电时，储存 内置水箱剩余燃料 的值到列表 S_RemFuelIn_value
                            # print('row[S_RemFuelIn]值为Nan,重新赋值等于0')
                        else:
                            No_S_RemFuelIn_value.append(
                                round(row[S_RemFuelIn], 1))  # 不发电时，储存 内置水箱剩余燃料 的值到列表 S_RemFuelIn_value

                        # 检查 row[S_RemFuelOut] 是否为 NaN 的数据类型为float
                        if np.isnan(row[S_RemFuelOut]):
                            No_S_RemFuelOut_value.append(0)
                        else:
                            No_S_RemFuelOut_value.append(round(row[S_RemFuelOut], 1))

                        No_LiqlelL.append(round(row[LiqlelL], 1))
                        No_LiqlelM.append(round(row[LiqlelM], 1))
                        No_HGHpre.append(round(row[HGHpre], 1))

                    One_DateTime.append(date_only[0])
                    # print(f'时间：{NO_DateTime}')

                    # 检查 No_S_RemFuelIn_value 是否为 NaN 的数据类型为float
                    print("内置液位 值类型:", type(No_S_RemFuelIn_value[0]))

                    # 内置液位(L)
                    start_S_RemFuelIn.append(No_S_RemFuelIn_value[1])
                    end_S_RemFuelIn.append(No_S_RemFuelIn_value[-1])
                    # 外置液位(L)
                    start_S_RemFuelOut.append(No_S_RemFuelOut_value[1])
                    end_S_RemFuelOut.append(No_S_RemFuelOut_value[-1])

                    # 外置液位(mm)
                    start_No_LiqlelL.append(No_LiqlelL[1])
                    end_No_LiqlelL.append(No_LiqlelL[-1])
                    # 内置液位(mm)
                    start_No_LiqlelM.append(No_LiqlelM[1])
                    end_No_LiqlelM.append(No_LiqlelM[-1])

                    HGHpre_time = 0
                    last_HGHpre_time = 0
                    current_HGHpre_time = 0

                    # 计算产氢次数
                    # 遍历列表中的元素
                    i = 0
                    while i < len(No_HGHpre) - 1:
                        differences = No_HGHpre[i] - No_HGHpre[i + 1]
                        if differences < -1.5 and No_HGHpre[i + 1] > 22.5:
                            No_HGHpre_Count.append(No_HGHpre[i + 1])
                            index_time = i + 1

                            # 计算产氢时间
                            current_HGHpre_time = datetime.strptime(NO_DateTime[index_time], '%Y-%m-%d %H:%M:%S')

                            if current_HGHpre_time and last_HGHpre_time:
                                HGHpre_time = round((current_HGHpre_time - last_HGHpre_time).total_seconds() / 60, 2)
                                print(f'当前时间点：{current_HGHpre_time}  ====  上个时间点：{last_HGHpre_time}')
                            last_HGHpre_time = current_HGHpre_time
                            if HGHpre_time:
                                # 储存平均产氢时间差的值到列表No_HGHpre_time_list
                                No_HGHpre_time_list.append(HGHpre_time)
                                print(f"时间差：{HGHpre_time} 分钟")
                            if max_index > 15000:
                                i += 3000
                            elif max_index > 10000:
                                i += 1000
                            elif max_index > 5000:
                                i += 500
                            elif max_index > 3000:
                                i += 250
                            else:
                                # 如果条件满足，跳过接下来的200个元素
                                i += 100  # 增加i的值，确保跳过200个元素

                        else:
                            # 如果条件不满足，正常递增i
                            i += 1  # 正常递增i

                    #     q.append(i)
                    # print(f"循环列表：======{q}")
                    if len(No_HGHpre_time_list) > 2:
                        average = round(sum(No_HGHpre_time_list) / len(No_HGHpre_time_list), 2)
                    else:
                        average = 0
                    No_HGHpre_time_average.append(average)
                    # print(f'平均')
                    print(f'平均产氢时间：{average}')
                    print(f"计算产气次数 ：{len(No_HGHpre_Count)}")

                    No_HGHpre_SumCount.append(len(No_HGHpre_Count))
                    # print(f'时间  列表：{One_DateTime}')
                    No_HGHpre_Count.clear()

                    print(f"时间 ：{date_only[0]}")
                    print(f'开始时内置液位(mm)：{start_No_LiqlelL[-1]}')
                    print(f'结束时内置液位(mm)：{end_No_LiqlelL[-1]}')
                    print(f'开始时外置液位(mm)：{start_No_LiqlelM[-1]}')
                    print(f'结束时外置液位(mm)：{end_No_LiqlelM[-1]}')

                    print(f'开始时内置液位(L)：{start_S_RemFuelIn[-1]}')
                    print(f'结束时内置液位(L)：{end_S_RemFuelIn[-1]}')
                    print(f'开始时外置液位(L)：{start_S_RemFuelOut[-1]}')
                    print(f'结束时外置液位(L)：{end_S_RemFuelOut[-1]}')

                    Max_Msw = max(No_S_RemFuelIn_value)
                    print(f"最大值*（L）:{Max_Msw}")

                    Max_Msw_mm = max(No_LiqlelM)
                    print(f"最大值*（mm）:{Max_Msw_mm}")

                    Min_Msw_mm = min(No_LiqlelM)
                    print(f"最小值*（mm）:{Min_Msw_mm}")

                    # print(f'燃料值（L）:{No_S_RemFuelIn_value}')
                    print(f"++++++液位（L）的列表:{start_S_RemFuelIn}")
                    print(f"------液位（MM）的列表:{start_No_LiqlelL}")

                    NO_differences = 0
                    print('@@@@@@@@@@@@', No_S_RemFuelIn_value[1], No_S_RemFuelOut_value[1])
                    # 如果列表不为空。说明有数据，执行下面计算
                    if No_S_RemFuelIn_value[1] > 0 and No_S_RemFuelOut_value[1] > 0:
                        # print('@@@@@@@@@@@@@@@@@@@@')
                        # 如果一天中有加液，找出最大值去减第一项，大于1。说明当天有加液
                        if (Max_Msw - No_S_RemFuelIn_value[1]) > 1:
                            first_RemFuelIn = No_S_RemFuelIn_value[1] - 15
                            second_RemFuelIn = Max_Msw - No_S_RemFuelIn_value[-1]
                            NO_differences = round(first_RemFuelIn + second_RemFuelIn, 2)
                            # print(f'燃料值的差====（L）:{NO_differences}')
                            # print(f"最大值-第一个:{Max_Msw - No_S_RemFuelIn_value[1]}")
                            # 如果计算出来的NO_differences液位消耗值小于0，则等于0
                            if NO_differences <= 0:
                                NO_differences = 0
                            # NO_Once_RemFuelIn = round(sum(NO_differences), 2)
                            NO_Once_S_RemFuelIn.append(NO_differences)
                            print(f'当天有加液 + 不发电消耗燃料（L）:{NO_differences} 内部')

                        else:

                            NO_differences = round(No_S_RemFuelIn_value[1] - No_S_RemFuelIn_value[-1], 2)
                            # 如果计算出来的NO_differences液位消耗值小于0，则等于0
                            if NO_differences <= 0:
                                NO_differences = 0
                            NO_Once_S_RemFuelIn.append(NO_differences)
                            # print('当天没有加液  ， 计算 L')
                            print(f'当天没有加液 + 不发电消耗燃料（L）:{NO_differences}  ')
                    # 如果列表为空。说明没有数据，执行下面计算
                    else:
                        # 如果一天中有加液，找出最大值去减第一项，大于1。说明当天有加液
                        if (Max_Msw_mm - No_LiqlelM[1]) > 50:
                            first_RemFuelIn = No_LiqlelM[1] - Min_Msw_mm
                            second_RemFuelIn = Max_Msw - No_LiqlelM[-1]
                            NO_differences = round(first_RemFuelIn + second_RemFuelIn, 2)
                            # print(f'燃料值的差====（L）:{NO_differences}')
                            # print(f"最大值-第一个:{Max_Msw - No_S_RemFuelIn_value[1]}")
                            # 如果计算出来的NO_differences液位消耗值小于0，则等于0
                            if NO_differences <= 0:
                                NO_differences = 0
                            # NO_Once_RemFuelIn = round(sum(NO_differences), 2)
                            NO_Once_S_RemFuelIn.append(NO_differences)
                            print(f'当天有加液 + 不发电消耗燃料（mm）:{NO_differences} 内部')

                        else:

                            NO_differences = round(No_LiqlelM[1] - No_LiqlelM[-1], 2)
                            # 如果计算出来的NO_differences液位消耗值小于0，则等于0
                            if NO_differences <= 0:
                                NO_differences = 0
                            NO_Once_S_RemFuelIn.append(NO_differences)
                            # print('当天没有加液  ， 计算 mm')
                            print(f'当天没有加液 + 不发电消耗燃料（mm）:{NO_differences}  ')

                    # 将待机时的 a1时间 添加到 Timer_RemFuelIn 数组里面。里面只包含待机时间数据
                    Timer_RemFuelIn.append(date_only[0])
                    # 将待机时的 NO_Once_S_RemFuelIn 液位消耗 求出总和
                    Sum_S_RemFuelIn = sum(NO_Once_S_RemFuelIn)
                    all_Sum_S_RemFuelIn.append(Sum_S_RemFuelIn)

                    # 如果没有故障，’备注‘里面写0 ,如果数据量小于3500，开始做备注
                    if 3500 >= max_index > 3000:
                        remark.append('当天数据缺失，数据总量小于3500行')
                    elif 3000 >= max_index > 2500:
                        remark.append('当天数据缺失，数据总量小于3000行')
                    elif 2500 >= max_index > 2000:
                        remark.append('当天数据缺失，数据总量小于2500行')
                    elif 2000 >= max_index > 1500:
                        remark.append('当天数据缺失，数据总量小于2000行')
                    elif 1500 >= max_index > 1000:
                        remark.append('当天数据缺失，数据总量小于1500行')
                    elif 1000 >= max_index > 500:
                        remark.append('当天数据缺失，数据总量小于1000行')
                    elif 500 > max_index:
                        remark.append('当天数据缺失，数据总量小于500行')
                    else:
                        remark.append(0)

                    print(f'备注：{remark}')

                    print(f'\n===========   {date_only[0]} 当天待机燃料消耗   ==========\n')
                else:
                    b1 = f'当天有发电，不计算待机燃料消耗'
                    o1 = 0
                    start_No_LiqlelL.append(o1)
                    end_No_LiqlelL.append(o1)
                    start_No_LiqlelM.append(o1)
                    end_No_LiqlelM.append(o1)
                    start_S_RemFuelOut.append(o1)
                    end_S_RemFuelOut.append(o1)
                    start_S_RemFuelIn.append(o1)
                    end_S_RemFuelIn.append(o1)
                    No_HGHpre_SumCount.append(o1)
                    all_Sum_S_RemFuelIn.append(o1)
                    One_DateTime.append(date_only[0])
                    No_HGHpre_time_average.append(o1)
                    remark.append(b1)
                    print(f'\n===========   {date_only[0]} 当天有发电，不计算燃料消耗   ==========\n')

                # print(f"总燃料消耗(L)：{Sum_S_RemFuelIn}")

                NO_Once_S_RemFuelIn.clear()
                No_S_RemFuelIn_value.clear()
                No_S_RemFuelOut_value.clear()
                No_LiqlelL.clear()
                No_LiqlelM.clear()
                NO_DateTime.clear()
                No_HGHpre.clear()
                q.clear()
                df_list.clear()
                No_HGHpre_time_list.clear()
                # 在控制台上打印，显示每列的长度(元素个数) ，如果长度(元素个数)不一样，会报错“输出的列长不一样”

                print(f"\n时间 长度：{len(One_DateTime)}")
                print(f"消耗燃料 长度：{len(all_Sum_S_RemFuelIn)}")

                print(f"内置-开始时液位(L) 长度：{len(start_S_RemFuelIn)}")
                print(f"内置-结束时液位(L) 长度：{len(end_S_RemFuelIn)}")
                print(f"外置-开始时液位(L） 长度：{len(start_S_RemFuelOut)}")
                print(f"外置-结束时液位(L) 长度：{len(end_S_RemFuelOut)}")
                print(f"外置-结束时液位(MM) 长度：{len(start_No_LiqlelL)}")
                print(f"外置-结束时液位(MM) 长度：{len(end_No_LiqlelL)}")
                print(f"内置-结束时液位(MM) 长度：{len(start_No_LiqlelM)}")
                print(f"内置-结束时液位(MM) 长度：{len(end_No_LiqlelM)}\n")
                print(f"产氢次数 长度：{len(No_HGHpre_SumCount)}")
                print(f"平均产氢时间 长度：{len(No_HGHpre_time_average)}")
                print(f"备注 长度：{len(remark)}")

                print(f'\n++++++++++++++  {date_only[0]} 一天的计算结束   ++++++++++++++++++++++++\n')

                # 储存 a1时间点 到 Timer_RemFuelIn列表 里面，用于在excel表格打印

            else:
                b1 = f' 当天没有数据，下载数据为空 ！！！'
                o1 = 0

                One_DateTime.append(One_DateTime[-1])

                start_No_LiqlelL.append(o1)
                end_No_LiqlelL.append(o1)
                start_No_LiqlelM.append(o1)
                end_No_LiqlelM.append(o1)
                start_S_RemFuelOut.append(o1)
                end_S_RemFuelOut.append(o1)
                start_S_RemFuelIn.append(o1)
                end_S_RemFuelIn.append(o1)
                all_Sum_S_RemFuelIn.append(o1)
                No_HGHpre_SumCount.append(o1)
                No_HGHpre_time_average.append(o1)

                remark.append(b1)
                print(
                    f'\n++++++++++++++   {One_DateTime[-1]}    当天没有数据，下载数据为空 ！！！    ++++++++++++++++++++++++\n')

        except FileNotFoundError:
            print(f"文件 {adress1} 不存在，已跳过")
    else:
        print(f"文件 {adress1} 不存在，已跳过")

A_all_Sum_S_RemFuelIn = sum(all_Sum_S_RemFuelIn)

print(f"总燃料消耗(L)：{A_all_Sum_S_RemFuelIn}\n")

# print(f"时间：{Timer_RemFuelIn}\n")

# 在控制台上打印，显示每列的长度(元素个数) ，如果长度(元素个数)不一样，会报错“输出的列长不一样”
print(f"时间 长度：{len(One_DateTime)}")
# print(f"时间 长度：{len(Timer_RemFuelIn)}")
print(f"消耗燃料 长度：{len(all_Sum_S_RemFuelIn)}")

print(f"内置-开始时液位(L) 长度：{len(start_S_RemFuelIn)}")
print(f"内置-结束时液位(L) 长度：{len(end_S_RemFuelIn)}")
print(f"外置-开始时液位(L） 长度：{len(start_S_RemFuelOut)}")
print(f"外置-结束时液位(L) 长度：{len(end_S_RemFuelOut)}")
print(f"外置-结束时液位(MM) 长度：{len(start_No_LiqlelL)}")
print(f"外置-结束时液位(MM) 长度：{len(end_No_LiqlelL)}")
print(f"内置-结束时液位(MM) 长度：{len(start_No_LiqlelM)}")
print(f"内置-结束时液位(MM) 长度：{len(end_No_LiqlelM)}")
print(f"产氢次数 长度：{len(No_HGHpre_SumCount)}")
print(f"平均产氢时间 长度：{len(No_HGHpre_time_average)}")
print(f"备注 长度：{len(remark)}")

print(f'燃料的值…………………………： {all_Sum_S_RemFuelIn}')

if start_S_RemFuelOut[1] > 0 and end_S_RemFuelOut[1] > 0:
    # 将新的DataFrame保存到新的Excel文件中
    new_df = pd.DataFrame(
        {
            '时间': One_DateTime,
            # '时间': Timer_RemFuelIn,
            '开始外置水箱剩余燃料(mm)': start_No_LiqlelL,
            '结束外置水箱剩余燃料(mm)': end_No_LiqlelL,
            '开始内置水箱剩余燃料(mm)': start_No_LiqlelM,
            '结束内置水箱剩余燃料(mm)': end_No_LiqlelM,
            '开始外置水箱剩余燃料(L)': start_S_RemFuelOut,
            '结束外置水箱剩余燃料(L)': end_S_RemFuelOut,
            '开始内置水箱剩余燃料(L)': start_S_RemFuelIn,
            '结束内置水箱剩余燃料(L)': end_S_RemFuelIn,

            '待机消耗燃料(L)': all_Sum_S_RemFuelIn,
            '产氢计数（次）': No_HGHpre_SumCount,
            '平均产氢时间（min）': No_HGHpre_time_average,
            '备注': remark,

        })

else:

    # 将新的DataFrame保存到新的Excel文件中
    new_df = pd.DataFrame(
        {
            '时间': One_DateTime,
            # '时间': Timer_RemFuelIn,
            '开始外置水箱剩余燃料(mm)': start_No_LiqlelL,
            '结束外置水箱剩余燃料(mm)': end_No_LiqlelL,
            '开始内置水箱剩余燃料(mm)': start_No_LiqlelM,
            '结束内置水箱剩余燃料(mm)': end_No_LiqlelM,
            '开始外置水箱剩余燃料(L)': start_S_RemFuelOut,
            '结束外置水箱剩余燃料(L)': end_S_RemFuelOut,
            '开始内置水箱剩余燃料(L)': start_S_RemFuelIn,
            '结束内置水箱剩余燃料(L)': end_S_RemFuelIn,

            '待机消耗燃料(mm)': all_Sum_S_RemFuelIn,
            '产氢计数（次）': No_HGHpre_SumCount,
            '平均产氢时间（min）': No_HGHpre_time_average,
            '备注': remark,

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
sheet.column_dimensions['A'].width = 21  # 第一列
# sheet.column_dimensions['B'].width = 21  # 第二列
# 设置其余列的宽度为 10
for col in sheet.columns:
    if col[0].column_letter not in ['A']:
        sheet.column_dimensions[col[0].column_letter].width = 15
# 遍历第一行的所有单元格，并为每个单元格对象同时设置自动换行、水平居中和垂直居中。
for cell in sheet[1]:
    cell_obj = cell
    cell_obj.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')

workbook.save(file_path)
print(f"\n文件保存成功 ！! ! ")
print(f"文件保存路径 ：{file_path}")
