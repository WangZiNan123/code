import mysql.connector
from mysql.connector import Error
import pandas as pd
import numpy as np

"""
版本：2024_6_27            版本时间：2024.6.27   
更新内容：抓取自定义数据，存储到服务器数据库里面.数据库表头格式如下：

    【  日期时间,设备编号,设备名称,设备网络状态,设备运行状态,设备母线电压(V),
        电池1_Soc,电池2_Soc,外置燃料(L),内置燃料(L),内置燃料(mm),A_制氢机状态,
        A_氢气压力(Psi),A_鼓风机温度(℃),A_提纯器温度(℃),A_重整室温度(℃),B_制氢机状态,
        B_氢气压力(Psi),B_鼓风机温度(℃),B_提纯器温度(℃),B_重整室温度(℃),A_电堆状态,A_电堆电压(V),
        A_电堆电流(A),A_电堆功率(W),A1_电堆堆心温度(℃),A2_电堆堆心温度(℃),A1_电堆顶部温度(℃),A2_电堆顶部温度(℃),
        B_电堆状态,B_电堆电压(V),B_电堆电流(A),B_电堆功率(W),B_电堆堆心温度(℃),B1_电堆顶部温度(℃),B2_电堆顶部温度(℃),备注
    】

    自动获取数据表“备注”里面的最后17行，17-34行数据，进行对比。
    
版本：2024_6_28            版本时间：2024.6.28   
更新内容：优化最后输出的逻辑 ，分新旧故障输出，对比最后17-34个的值与最后17个的值，如果相同，则输出旧故障，不同输出新故障


"""

new_alarm = []  # 新的故障告警
old_alarm = []  # 旧的故障告警
same_alarm = []  # 相同的故障

# 配置数据库连接参数
db_config = {
    'host': '8.138.136.163',
    'user': 'root',
    'password': '123456wang',
    'database': '网页爬虫',
    'raise_on_warnings': True
}

# 连接到MySQL数据库
try:
    conn = mysql.connector.connect(**db_config)

    if conn.is_connected():
        print("数据库连接成功。")

        # # 定义要插入的数据。。测试
        # users = [("1231", "Doe", "john@example.com", "12", "12", "13", "123", "123", "12312412",
        #           "1231", "12", "12", "13", "32", "1231", "12", "12", "13", "32",
        #           "1231", "12", "12", "13", "32", "1231", "12", "12",
        #           "1231", "12", "12", "13", "32", "1231", "12", "12")]

        # 接下来的代码（数据库连接、创建表、插入数据等）保持不变
        cursor = conn.cursor()
        cursor.execute("SHOW TABLES LIKE %s", ('COWIN_爬虫数据库',))
        table_exists = cursor.fetchone()
        if table_exists:
            print(f"表 'COWIN_爬虫数据库' 已存在。")
        else:
            print(f"创建一个新表 'COWIN_爬虫数据库'")

            cursor.execute("""     CREATE TABLE `COWIN_爬虫数据库`  (
                                  `id` INT NOT NULL AUTO_INCREMENT,
                                  `日期时间` VARCHAR(100),
                                  `设备编号` VARCHAR(100),
                                  `设备名称` VARCHAR(100),
                                  `设备网络状态` VARCHAR(100) ,
                                  `设备运行状态` VARCHAR(100) ,
                                  `设备母线电压(V)` DOUBLE,
                                  `电池1_Soc` DOUBLE,
                                  `电池2_Soc` DOUBLE,
                                  `外置燃料(L)` DOUBLE,
                                  `内置燃料(L)` DOUBLE,
                                  `内置燃料(mm)` DOUBLE,

                                  `A_制氢机状态` VARCHAR(100),
                                  `A_氢气压力(Psi)` DOUBLE,
                                  `A_鼓风机温度(℃)` DOUBLE,
                                  `A_提纯器温度(℃)` DOUBLE,
                                  `A_重整室温度(℃)` DOUBLE,

                                  `B_制氢机状态` VARCHAR(100),
                                  `B_氢气压力(Psi)` DOUBLE,
                                  `B_鼓风机温度(℃)` DOUBLE,
                                  `B_提纯器温度(℃)` DOUBLE,
                                  `B_重整室温度(℃)` DOUBLE,

                                  `A_电堆状态` VARCHAR(100),
                                  `A_电堆电压(V)` DOUBLE,
                                  `A_电堆电流(A)` DOUBLE,
                                  `A_电堆功率(W)` DOUBLE,
                                  `A1_电堆堆心温度(℃)` DOUBLE,
                                  `A2_电堆堆心温度(℃)` DOUBLE,
                                  `A1_电堆顶部温度(℃)` DOUBLE,
                                  `A2_电堆顶部温度(℃)` DOUBLE,

                                  `B_电堆状态` VARCHAR(100),
                                  `B_电堆电压(V)` DOUBLE,
                                  `B_电堆电流(A)` DOUBLE,
                                  `B_电堆功率(W)` DOUBLE,
                                  `B_电堆堆心温度(℃)` DOUBLE,
                                  `B1_电堆顶部温度(℃)` DOUBLE,
                                  `B2_电堆顶部温度(℃)` DOUBLE,

                                  `备注` TEXT,

                                  PRIMARY KEY (`id`)
                                )   """)
        sql = ("INSERT INTO COWIN_爬虫数据库 ("
               "`日期时间`, "
               "`设备编号`, "
               "`设备名称`, "
               "`设备网络状态`, "
               "`设备运行状态`,"
               "`设备母线电压(V)`,"
               "`电池1_Soc`,"
               "`电池2_Soc`,"
               "`外置燃料(L)`,"
               "`内置燃料(L)`,"
               "`内置燃料(mm)`,"

               "`A_制氢机状态`,"
               "`A_氢气压力(Psi)`,"
               "`A_鼓风机温度(℃)`,"
               "`A_提纯器温度(℃)`,"
               "`A_重整室温度(℃)`,"

               "`B_制氢机状态`,"
               "`B_氢气压力(Psi)`,"
               "`B_鼓风机温度(℃)`,"
               "`B_提纯器温度(℃)`,"
               "`B_重整室温度(℃)`,"

               "`A_电堆状态`,"
               "`A_电堆电压(V)`,"
               "`A_电堆电流(A)`,"
               "`A_电堆功率(W)`,"
               "`A1_电堆堆心温度(℃)`,"
               "`A2_电堆堆心温度(℃)`,"
               "`A1_电堆顶部温度(℃)`,"
               "`A2_电堆顶部温度(℃)`,"

               "`B_电堆状态`,"
               "`B_电堆电压(V)`,"
               "`B_电堆电流(A)`,"
               "`B_电堆功率(W)`,"
               "`B_电堆堆心温度(℃)`,"
               "`B1_电堆顶部温度(℃)`,"
               "`B2_电堆顶部温度(℃)`,"

               "`备注`"

               ")"
               "VALUES (%s, %s, %s, %s, %s,%s,%s,%s,%s,"
               "%s, %s, %s, %s, %s,%s, %s, %s, %s, %s,"
               "%s, %s, %s, %s, %s,%s, %s, %s,"
               "%s, %s, %s, %s, %s,%s, %s, %s, %s, %s)")

        temp = pd.read_excel('D:/爬虫数据/网页采集数据_7.xlsx')

        # for i in range(len(temp.columns)):
        #     print(temp.columns[i],end=',')
        #
        # print(len(temp.columns),end='\n\n')

        # 创建一个空列表来存储元组
        tuples_list = []
        temp = temp.fillna(0)
        # 使用iterrows()遍历DataFrame中的行
        for index, row in temp.iterrows():
            # 获取当前行的值作为数组，然后转换成元组
            row_tuple = tuple(row.values)
            # 将元组添加到列表中
            tuples_list.append(row_tuple)

        # 打印结果
        # print(tuples_list)

        # print(len(tuples_list[0]))  # 打印第一个元组，检查值的数量和类型
        cursor.executemany(sql, tuples_list)

        # 提交事务
        conn.commit()

        tuples_list.clear()

        print("数据插入成功。")

        table_name = 'COWIN_爬虫数据库'
        column_name = '备注'
        primary_key = 'id'
        value17_34 = 'id'
        # 编写SQL查询语句
        # query = f"SELECT {column_name} FROM {table_name};"
        # 编写SQL查询语句
        # 编写SQL查询语句
        query17_34 = f"SELECT {column_name} FROM {table_name} ORDER BY {value17_34} DESC LIMIT 17 OFFSET 17;"
        query = f"SELECT {column_name} FROM {table_name} ORDER BY {primary_key} DESC LIMIT 17;"
        # 执行查询
        cursor.execute(query)
        # 获取查询结果
        results = cursor.fetchall()
        # 执行查询
        cursor.execute(query17_34)
        # 获取查询结果
        results17_34 = cursor.fetchall()

        results_list17_34 = []
        results_list = []
        # 打印结果
        for result in results:
            if result[0] != '0':
                results_list.append(result[0])
        # print(f'最后插入的17个数故障表：{results_list}')

        # 打印结果
        for result1 in results17_34:
            if result1[0] != '0':
                results_list17_34.append(result1[0])

        new_alarm.clear()
        old_alarm.clear()

        print(f'results_list元素：:{results_list}')
        print(f'results_list17_34元素：{results_list17_34}')

        # # 检查 '备注' 里面是否有故障
        # if len(results_list) > 0:
        #     # 将 results_list17_34 中的元素添加到一个临时列表中，用于后续比较
        #     temp_list = results_list17_34.copy()
        #
        #     # 遍历 results_list 中的每个元素
        #     for fault in results_list:
        #         # 检查故障是否在 results_list17_34 中
        #         if fault in temp_list:
        #             # 如果是旧故障，添加到 old_alarm 列表
        #             old_alarm.append(fault)
        #             # 从 temp_list 中移除该故障，以避免重复判断
        #             temp_list.remove(fault)
        #         else:
        #             # 如果是新故障，添加到 new_alarm 列表
        #             new_alarm.append(fault)
        #
        # # 使用列表推导式去除 old_alarm 中的重复项
        # old_alarm = [old_alarm[i] for n, i in enumerate(old_alarm) if old_alarm.index(old_alarm[i]) == n]
        #
        #     print(f'旧故障：{old_alarm}')
        #     print(f'新故障：{new_alarm}')

        # 检查 '备注' 里面是否有故障
        if len(results_list) > 0:
            # 检查是否有非 '0.0' 的元素
            if any(element != '0.0' for element in results_list):
                # 将 results_list17_34 中的所有元素添加到一个集合中，以提高查找效率
                set_results_list17_34 = set(results_list17_34)

                # 遍历 results_list 中的每个元素
                for fault in results_list:
                    # 如果该元素不在 results_list17_34 中，说明是新故障
                    if fault not in set_results_list17_34:
                        new_alarm.append(fault)
                    else:
                        # 如果该元素在 results_list17_34 中，说明是旧故障
                        old_alarm.append(fault)

                # 去除 old_alarm 中的重复项 。将列表转换成集合（set），自动去除重复元素，然后再将结果转换回列表。
                # 集合是一个无序的数据结构，它不允许有重复的元素。
                # old_alarm = list(set(old_alarm))
                # 打印结果
                print(f'旧故障：{old_alarm}')
                print(f'新故障：{new_alarm}')
            else:
                # 所有元素都是 '0.0'，没有故障，所以不执行任何操作
                print("没有故障信息，无需计算。")
        else:

            print(f'没有故障发生 ......')

        results_list17_34.clear()
        results_list.clear()

except Error as e:
    print("数据库操作出错：", e)

finally:
    # 关闭游标和连接
    if conn.is_connected():
        cursor.close()
        conn.close()
        print("MySQL连接已关闭。")
