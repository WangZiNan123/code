import mysql.connector
from mysql.connector import Error
import pandas as pd
import numpy as np

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

        # 定义要插入的数据
        users = [("1231", "Doe", "john@example.com", "12", "12", "13", "14")
                 ]

        # 接下来的代码（数据库连接、创建表、插入数据等）保持不变
        cursor = conn.cursor()



        cursor.execute("""     CREATE TABLE `爬虫测试`  (
                              `id` INT NOT NULL AUTO_INCREMENT,
                              `日期时间` VARCHAR(100),
                              `设备名称` VARCHAR(100),
                              `设备网络状态` VARCHAR(100) ,
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

        sql = ("INSERT INTO 爬虫测试 ("
               "`日期时间`, "
               "`设备名称`, "
               "`设备网络状态`,"
               "`设备母线电压(V)`,"
               "`电池1_Soc`,"
               "`电池2_Soc`,"
               "`外置燃料(L)`"

               ") "
               "VALUES (%s, %s, %s, %s, %s, %s, %s)")

        print(len(users[0]))  # 打印第一个元组，检查值的数量和类型
        cursor.executemany(sql, users)

        # 提交事务
        conn.commit()

        print("数据插入成功。")

except Error as e:
    print("数据库操作出错：", e)

finally:
    # 关闭游标和连接
    if conn.is_connected():
        cursor.close()
        conn.close()
        print("MySQL连接已关闭。")
