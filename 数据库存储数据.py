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

"""
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

        temp = pd.read_excel('E:/网页爬虫数据/网页采集数据_7.xlsx')

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

except Error as e:
    print("数据库操作出错：", e)

finally:
    # 关闭游标和连接
    if conn.is_connected():
        cursor.close()
        conn.close()
        print("MySQL连接已关闭。")
