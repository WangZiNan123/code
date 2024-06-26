import re
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import openpyxl
import os

'''
================================================= 

版本更新：2024_6_18   更新时间2024.6.18
网页爬虫 COWIN 数据，目前测试抓取’5G汇聚机房1‘ 的数据 ’记录时间，设备编号（Serial_No），设备名称（Remark），氢气压力（H2_Pressure）
提纯器温度（Purifier_temperature），重整室温度（Reformer_Temperature），鼓风机温度（Blower_temperature），电堆电压（Stack_voltage）
电堆电流（Stack_current），电堆功率（Stack_power），A电堆堆心温度（Stack_temperature）‘

版本更新：2024_6_22   更新时间2024.6.22
更新内容：新增 A1电堆顶部温度（发电仓温度(℃):），A2电堆顶部温度（环境温度(℃):），B1电堆顶部温度（环境湿度(%):），B2电堆顶部温度（电堆风机馈速(%): ）
        系统状态（System: ），母线电压（Current Voltage(V)：）

版本更新：2024_6_23   更新时间2024.6.23
更新内容：新增 优化代码格式  ，新增跳转到第二页 ，处理第二页的数据  。
        新增 A2电堆堆心温度（电堆温度2(℃):）
        新增 文件保存，将读取的数据保存为excel格式

版本更新：2024_6_24   更新时间2024.6.24
更新内容：新增 内置燃料值（Remaining Fuel(LIn)：），‘备注’可以将有异常的故障展示出来 ，设备网络状态（Off-line/On-line）  。
        新增 B制氢机参数（Hydrogen production module B）：
            氢气压力（H2_Pressure），鼓风机温度（Blower_temperature），提纯器温度（Purifier_temperature），重整室温度（Reformer_Temperature），制氢机运行状态（Module status: ）
        新增 B电堆参数（Power generation module B）：    
            电堆电压（Stack_voltage），电堆电流（Stack_current），电堆功率（Stack_power），B电堆堆心温度（Stack_temperature），电堆运行状态（Module status: ）


版本更新：2024_6_25   更新时间2024.6.25
更新内容：新增 外置燃料值（Remaining Fuel(LOut)：）  ，内置燃料值（液位小水箱(L):）
        优化故障处理逻辑


版本更新：2024_6_26   更新时间2024.6.26
更新内容：优化制氢机“关机”状态下，故障处理逻辑
        新增 电池1Soc（安培秒累加值:） ，电池2Soc（电堆排气次数:）
        
================================================= 

设备编号 ： 设备名称 

CW-10KW-0007 : 管委会10KW
CW-MFC6000-0001 : 5G汇聚机房1
CW-MFC6000-0002 : 5G汇聚机房2
CW-MFC6000-0008 : 江门电信白石机房1
CW-MFC6000-0010 : 江门电信白石机房2
MFC6kD480012 : 江门新会洋美接入网机房
MFC6kD480013 : 江门新会红关接入网机房
MFC6kD480014 : 江门台山墩寨综合机房
MFC6kD480015 : 江门开平潭溪综合机房
MFC6kD480016 : 江门台山华安接入网机房
MFC6kD480017 : 江门开平新美接入网机房
MFC6kD480018 : 江门鹤山升平接入网机房
MFC6kD480019 : 江门恩平平石接入网机房
MFC6kD480020 : 江门恩平三联接入网机房
MFC6kD480021 : 江门鹤山三堡接入网机房
MFC6kD480022 : 台山川岛长堤
MFC6kD480023 : 江油太平唐僧


================================================= 
'''

time_localtime_list = []
Serial_No_list = []
machine_name_list = []
A_H2_Pressure_list = []
A_Purifier_temperature_list = []
A_Reformer_Temperature_list = []
A_Blower_temperature_list = []
A_Stack_voltage_list = []
A1_Stack_temperature_list = []
A2_Stack_temperature_list = []

A_Stack_current_list = []
A_Stack_power_list = []

A1_Stack_top_temperature_list = []  #
A2_Stack_top_temperature_list = []
B1_Stack_top_temperature_list = []
B2_Stack_top_temperature_list = []

A_HG_Module_status_list = []  # 制氢机状态
A_Stack_Module_status_list = []  # 电堆状态
System_status_list = []  # 系统状态
Current_Voltage_list = []  # 母线电压

remark = []  # 备注
remark_set = []  # 备注

network_state_list = []  # 设备网络状态
Remaining_Fuel_list = []  # 内置水箱液位（L）

B_HG_Module_status_list = []
B_H2_Pressure_list = []
B_Blower_temperature_list = []
B_Purifier_temperature_list = []
B_Reformer_Temperature_list = []

B_Stack_voltage_list = []
B_Stack_current_list = []
B_Stack_power_list = []
B_Stack_Module_status_list = []
B_Stack_temperature_list = []

Out_Remaining_Fuel_list = []  # 外置液位（L）
In_Remaining_Fuel_mm_list = []  # 内置液位（mm）

battery_1_Soc_list = []  # 电池 1 容量
battery_2_Soc_list = []  # 电池 2 容量


def Program_Init(driver_path, url, loginName, passWord):
    """

    :param driver_path: 谷歌驱动路径
    :param url: 网页链接
    :param loginName: 网页登录账号名称
    :param passWord: 网页登录账号密码
    :return:
    """
    # 创建Service对象，指定ChromeDriver路径
    service = Service(executable_path=driver_path)

    # 使用Service对象作为服务启动Chrome
    driver = webdriver.Chrome(service=service)

    # 使用Selenium打开网页
    driver.get(url)

    # 等待登录页面加载完成
    wait = WebDriverWait(driver, 60)

    # 定位账号和密码输入框，以及登录按钮，并输入账号密码
    username_input = wait.until(EC.presence_of_element_located((By.ID, 'loginName')))
    password_input = wait.until(EC.presence_of_element_located((By.ID, 'password')))
    # 使用CSS选择器定位登录按钮
    login_button = driver.find_element(By.CSS_SELECTOR, 'button.login-button.ant-btn.ant-btn-primary.ant-btn-lg')

    username_input.send_keys(loginName)
    password_input.send_keys(passWord)

    time.sleep(1)
    # 提交登录信息
    login_button.click()

    time.sleep(1)

    return driver, wait


def click_Equipment_List(wait, driver):
    """
    点击 ‘ Equipment_List ’，跳转页面到设备选项页面

    参数:
    - driver:   从外部传入driver ，否则无法使用driver方法
    - wait: 从外部传入wait ，否则无法使用wait方法

    :return:
    """
    time.sleep(1.5)
    # 使用更具体的CSS选择器，确保选中的是可点击的元素
    submenu_title = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-menu-submenu-title')))

    # 点击下拉菜单标题以展开菜单
    submenu_title.click()

    time.sleep(1.5)
    # 定位并点击“Equipment List”列表项
    # 如果菜单项是一个<a>标签包裹<span>，确保XPath正确地定位到这个<a>标签
    equipment_list_item = wait.until(EC.element_to_be_clickable((By.XPATH, '//a/span[text()="Equipment List"]')))
    # 使用 JavaScript 执行点击操作

    driver.execute_script("arguments[0].click();", equipment_list_item)


def click_find_target_Details(driver, row_key):
    """
    找到目标行，并点击 "Details"

    参数:

    - driver: 从外部传入driver，否则无法使用‘driver’

    - row_key：目标表格行 ‘<tr> data-row-key’

    :return:
    """

    # 等待表格体元素加载完成
    wait = WebDriverWait(driver, 30)
    table_body_selector = ".ant-table-body"
    table_body_element = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, table_body_selector))

    # 使用JavaScript滚动到最右侧
    # 获取元素的滚动宽度
    scroll_width = driver.execute_script("return arguments[0].scrollWidth;", table_body_element)
    # 获取元素当前已经滚动的距离
    current_scroll = driver.execute_script("return window.pageXOffset || document.documentElement.scrollLeft;")
    # 计算需要滚动的距离
    scroll_to = scroll_width - current_scroll
    # 滚动到最右侧
    driver.execute_script(f"arguments[0].scrollLeft = {scroll_to};", table_body_element)
    time.sleep(1)
    # 等待表格行元素加载完成

    row_css_selector = f".ant-table-row[data-row-key='{row_key}']"
    row_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, row_css_selector)))

    # 使用 JavaScript 滚动到表格行，并使其位于视窗底部
    # 这里我们传递了 true 作为第二个参数给 scrollIntoView 方法
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'end'});", row_element)
    time.sleep(1.5)
    # < tbody >  ->   <tr> ->  "Details"
    # 定位到 <tbody> 元素
    tbody_selector = ".ant-table-tbody"
    tbody_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, tbody_selector)))

    # 在 <tbody> 中找到具有特定 data-row-key 的 <tr> 元素
    target_row_css_selector = f".ant-table-row[data-row-key='{row_key}']"
    # 正确使用 find_element 在 <tbody> 的上下文中查找 <tr>
    target_row_element = tbody_element.find_element(By.CSS_SELECTOR, target_row_css_selector)

    # 使用类名'.ant-badge-status-text'寻找具有相同类名的元素。提取网络状态
    Css_network_state = '.ant-badge-status-text'
    network_state = target_row_element.find_element(By.CSS_SELECTOR, Css_network_state).text
    network_state_list.append(network_state)

    # 在找到的 <tr> 元素中定位到 "Details" 链接
    # 假设 "Details" 链接具有 data-v-11b2bf7e 属性
    time.sleep(1)
    details_link_selector = "a[data-v-11b2bf7e]"
    details_link = target_row_element.find_element(By.CSS_SELECTOR, details_link_selector)
    # print(type(details_link))
    # 点击 "Details"

    details_link.click()
    # 使用 JavaScript 滚动到页面顶部
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, 0);")

    time.sleep(5)


def split_text_by_colon(wait, Xpath, split_located):
    """
    找出对应的抓取目标的值。
    使用正则表达式在英文冒号或中文冒号处分割文本，并返回最后一个分割的部分。

    参数:

    - wait: 从外部传入wait ，否则无法使用wait方法

    - Xpath:  元素的绝对路径

    - text: 要分割的文本字符串。

    返回:
    - 分割后去除首尾空白的最后一个部分。

    """
    text = wait.until(EC.presence_of_element_located((By.XPATH, Xpath))).text

    pattern = re.compile(r'[:：]')
    parts = re.split(pattern, text)[split_located].strip()
    if len(parts) == 0:
        parts = 0
    return parts


#   跳转函数
def JMP(driver, wait, row_key):
    """
    跳转页面函数，找到对应设备的参数页面

    :param wait: 从外部传入
    :param row_key: 从外部传入，指定要跳转到哪行
    :param driver: 从外部传入
    :return:
    """
    # 等待下拉菜单标题加载完成
    # 点击 ‘ Equipment_List ’，跳转页面到设备选项页面
    click_Equipment_List(wait, driver)

    # 找到目标行，并点击"Details"
    click_find_target_Details(driver, row_key)

    # # 这里添加10秒的等待时间
    # time.sleep(5)


def jum_page_2(driver, wait):
    """
    跳转到设备列表第2页

    :param wait:    从外部传入
    :param driver: 从外部传入
    :return:
    """
    # 回到设备选择页面
    click_Equipment_List(wait, driver)

    time.sleep(2)

    # 使用 JavaScript 滚动到页面底部
    driver.execute_script("window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });")
    time.sleep(2)
    # details_link_selector = "ant-pagination-item ant-pagination-item-2"
    # 使用 title 属性定位分页项
    page2 = wait.until(EC.element_to_be_clickable((By.XPATH, '//li[@title="2"]')))
    # details_link = driver.find_element(By.CSS_SELECTOR, details_link_selector)
    # details_link.click()
    driver.execute_script("arguments[0].click();", page2)


def data_processing(driver, wait):
    #   设备编号 绝对地址
    Serial_No_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/span[1]/span'
    #   设备名称 绝对地址
    Remark_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div[2]/div[1]'
    #   系统状态 绝对地址
    System_status_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div[9]/div[3]'
    #   母线电压 绝对地址
    Current_Voltage_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div[8]/div[2]'
    #   内置燃料剩余液位(L)
    Remaining_Fuel_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div[4]/div[3]'
    #  外置燃烧剩余液位（L）
    Out_Remaining_Fuel_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div[5]/div[1]'
    #   内置燃料剩余液位(mm)
    In_Remaining_Fuel_mm_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[6]/div[1]/div[2]/div[8]/div[2]'

    #   A制氢机状态 绝对地址
    HG_Module_status_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[4]/div[2]'
    #   A重整室温度 绝对地址
    Reformer_Temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[3]/div[2]'
    #   A鼓风机温度 绝对地址
    Blower_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[1]/div[3]'
    # A提纯器温度。使用XPath定位元素
    Purifier_temperature_XPath = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[3]/div[1]"
    # A缓冲罐氢气压力。使用提供的XPath定位元素
    H2_Pressure_XPath = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[1]/div[1]"

    #   A电堆状态 绝对地址
    Stack_Module_status_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[4]/div[2]'
    #   A电堆电流 绝对地址
    Stack_current_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[2]'
    #   A电堆功率 绝对地址
    Stack_power_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[3]'
    #  A电堆电压 绝对地址
    Stack_voltage_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]'
    #   A1电堆堆心温度 绝对地址
    A1_Stack_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[2]/div[1]'
    #   A2电堆堆心温度 绝对地址
    A2_Stack_temperature_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[7]/div[2]/div[2]/div[2]/div[3]'
    #   A1电堆顶部温度 绝对地址
    A1_Stack_top_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[6]/div[1]/div[2]/div[7]/div[1]'
    #   A2电堆顶部温度 绝对地址
    A2_Stack_top_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[6]/div[1]/div[2]/div[7]/div[2]'

    #   B制氢机运行状态
    B_HG_Module_status_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div[4]/div[2]'
    # B制氢机氢气压力
    B_H2_Pressure_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div[1]/div[1]'
    #   B制氢机鼓风机温度
    B_Blower_temperature_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div[1]/div[3]'
    #   B制氢机提纯器温度
    B_Purifier_temperature_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div[3]/div[1]'
    #   B制氢机重整室温度
    B_Reformer_Temperature_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div[3]/div[2]'

    #  B电堆电压 绝对地址
    B_Stack_voltage_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[1]/div[1]'
    #   B电堆电流 绝对地址
    B_Stack_current_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[1]/div[2]'
    #   B电堆功率 绝对地址
    B_Stack_power_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[1]/div[3]'
    #   B电堆状态 绝对地址
    B_Stack_Module_status_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[4]/div[2]'
    #   B电堆堆心温度 绝对地址
    B_Stack_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]'
    #   B1电堆顶部温度 绝对地址
    B1_Stack_top_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[6]/div[1]/div[2]/div[7]/div[3]'
    #   B2电堆顶部温度 绝对地址
    B2_Stack_top_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[6]/div[2]/div[2]/div[6]/div[3]'
    #   电池1 SOC 绝对地址
    battery_1_Soc_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[6]/div[2]/div[2]/div[1]/div[3]'
    #   电池2 SOC 绝对地址
    battery_2_Soc_XPath = '/html/body/div/div/div[2]/div[2]/div/div/div/div/div[6]/div[2]/div[2]/div[2]/div[3]'

    try:

        #   日期时间
        time_localtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
        #   设备编号
        Serial_No = split_text_by_colon(wait, Serial_No_XPath, -1)
        #   设备名称
        Remark_temp = split_text_by_colon(wait, Remark_XPath, 1)
        Remark = Remark_temp.split(' ', 1)[0]
        # 等待元素出现
        #   A制氢机氢气压力
        A_H2_Pressure = split_text_by_colon(wait, H2_Pressure_XPath, -1)

        A_Purifier_temperature = split_text_by_colon(wait, Purifier_temperature_XPath, -1)

        A_Reformer_Temperature = split_text_by_colon(wait, Reformer_Temperature_XPath, -1)

        A_Blower_temperature = split_text_by_colon(wait, Blower_temperature_XPath, -1)

        A_Stack_voltage = split_text_by_colon(wait, Stack_voltage_XPath, -1)

        A1_Stack_temperature = split_text_by_colon(wait, A1_Stack_temperature_XPath, -1)

        A2_Stack_temperature = split_text_by_colon(wait, A2_Stack_temperature_XPath, -1)

        A1_Stack_top_temperature = split_text_by_colon(wait, A1_Stack_top_temperature_XPath, -1)

        A2_Stack_top_temperature = split_text_by_colon(wait, A2_Stack_top_temperature_XPath, -1)

        B1_Stack_top_temperature = split_text_by_colon(wait, B1_Stack_top_temperature_XPath, -1)

        B2_Stack_top_temperature = split_text_by_colon(wait, B2_Stack_top_temperature_XPath, -1)

        A_Stack_current = split_text_by_colon(wait, Stack_current_XPath, -1)

        A_Stack_power = split_text_by_colon(wait, Stack_power_XPath, -1)

        # print(f'Stack_power字符串长度：{len(Stack_power)}')
        A_HG_Module_status = split_text_by_colon(wait, HG_Module_status_XPath, -1)

        A_Stack_Module_status = split_text_by_colon(wait, Stack_Module_status_XPath, -1)

        System_status = split_text_by_colon(wait, System_status_XPath, -1)

        Current_Voltage = split_text_by_colon(wait, Current_Voltage_XPath, -1)

        Remaining_Fuel = split_text_by_colon(wait, Remaining_Fuel_XPath, -1)

        B_HG_Module_status = split_text_by_colon(wait, B_HG_Module_status_XPath, -1)

        B_H2_Pressure = split_text_by_colon(wait, B_H2_Pressure_XPath, -1)

        B_Blower_temperature = split_text_by_colon(wait, B_Blower_temperature_XPath, -1)

        B_Purifier_temperature = split_text_by_colon(wait, B_Purifier_temperature_XPath, -1)

        B_Reformer_Temperature = split_text_by_colon(wait, B_Reformer_Temperature_XPath, -1)

        B_Stack_voltage = split_text_by_colon(wait, B_Stack_voltage_XPath, -1)

        B_Stack_power = split_text_by_colon(wait, B_Stack_power_XPath, -1)

        B_Stack_Module_status = split_text_by_colon(wait, B_Stack_Module_status_XPath, -1)

        B_Stack_temperature = split_text_by_colon(wait, B_Stack_temperature_XPath, -1)

        B_Stack_current = split_text_by_colon(wait, B_Stack_current_XPath, -1)

        Out_Remaining_Fuel = split_text_by_colon(wait, Out_Remaining_Fuel_XPath, -1)

        In_Remaining_Fuel_mm = split_text_by_colon(wait, In_Remaining_Fuel_mm_XPath, -1)

        battery_1_Soc = split_text_by_colon(wait, battery_1_Soc_XPath, -1)

        battery_2_Soc = split_text_by_colon(wait, battery_2_Soc_XPath, -1)

        time_localtime_list.append(time_localtime)
        Serial_No_list.append(Serial_No)
        machine_name_list.append(Remark)
        A_H2_Pressure_list.append(round(float(A_H2_Pressure), 2))
        A_Purifier_temperature_list.append(round(float(A_Purifier_temperature), 2))
        A_Reformer_Temperature_list.append(round(float(A_Reformer_Temperature), 2))
        A_Blower_temperature_list.append(round(float(A_Blower_temperature), 2))
        A_Stack_voltage_list.append(round(float(A_Stack_voltage), 2))
        A1_Stack_temperature_list.append(round(float(A1_Stack_temperature), 2))
        A2_Stack_temperature_list.append(round(float(A2_Stack_temperature), 2))

        A1_Stack_top_temperature_list.append(round(float(A1_Stack_top_temperature), 2))
        A2_Stack_top_temperature_list.append(round(float(A2_Stack_top_temperature), 2))
        B1_Stack_top_temperature_list.append(round(float(B1_Stack_top_temperature), 2))
        B2_Stack_top_temperature_list.append(round(float(B2_Stack_top_temperature), 2))

        A_Stack_current_list.append(round(float(A_Stack_current), 2))
        A_Stack_power_list.append(round(float(A_Stack_power), 2))

        A_HG_Module_status_list.append(A_HG_Module_status)
        A_Stack_Module_status_list.append(A_Stack_Module_status)
        System_status_list.append(System_status)
        Current_Voltage_list.append(round(float(Current_Voltage), 2))
        Remaining_Fuel_list.append(round(float(Remaining_Fuel), 2))

        B_HG_Module_status_list.append(B_HG_Module_status)
        B_H2_Pressure_list.append(round(float(B_H2_Pressure), 2))
        B_Blower_temperature_list.append(round(float(B_Blower_temperature), 2))
        B_Purifier_temperature_list.append(round(float(B_Purifier_temperature), 2))
        B_Reformer_Temperature_list.append(round(float(B_Reformer_Temperature), 2))

        B_Stack_voltage_list.append(round(float(B_Stack_voltage), 2))
        B_Stack_power_list.append(round(float(B_Stack_power), 2))
        B_Stack_Module_status_list.append(B_Stack_Module_status)
        B_Stack_temperature_list.append(round(float(B_Stack_temperature), 2))
        B_Stack_current_list.append(round(float(B_Stack_current), 2))

        Out_Remaining_Fuel_list.append(round(float(Out_Remaining_Fuel), 2))
        In_Remaining_Fuel_mm_list.append(round(float(In_Remaining_Fuel_mm), 2))

        battery_1_Soc_list.append(round(float(battery_1_Soc), 2))
        battery_2_Soc_list.append(round(float(battery_2_Soc), 2))

        print("日期时间：", time_localtime_list[-1])

        print("设备编号：", Serial_No_list[-1])

        print("设备名称：", machine_name_list[-1])

        if network_state_list[-1] == 'On-line':
            print("设备网络状态：", network_state_list[-1])
        else:
            print("设备网络状态：", network_state_list[-1], '           设备网络离线      ！！！')
            remark.append(' 设备网络离线 ！')

        # 如果设备处于‘关机’状态，执行以下代码
        if System_status_list[-1] == 'Off':

            print("设备运行状态：", System_status_list[-1])

            if float(Current_Voltage_list[-1]) >= 49.8 or Serial_No_list[-1] in (
                    'CW-MFC6000-0001', 'CW-MFC6000-0002', 'CW-MFC6000-0008', 'CW-MFC6000-0010'):
                print("设备母线电压(V)：", Current_Voltage_list[-1])
            else:
                print("设备母线电压(V)：", Current_Voltage_list[-1], "      设备母线电压太低异常     ！！！")
                remark.append(f' 设备母线电压太低异常（ {Current_Voltage_list[-1]}(V) ） ！')

            # 针对管委会，川岛，四川 有两块电池情况
            if Serial_No_list[-1] in ('CW-10KW-0007', 'MFC6kD480022', 'MFC6kD480023'):

                if float(battery_1_Soc_list[-1]) >= 60:
                    print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}')
                else:

                    print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}         电池1_Soc太低异常     ！！！')
                    remark.append(f' 电池1_Soc太低异常（ {battery_1_Soc_list[-1]} ） ！')

                if float(battery_2_Soc_list[-1]) >= 60:
                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}')
                else:

                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}         电池2_Soc太低异常     ！！！')
                    remark.append(f' 电池2_Soc太低异常（ {battery_2_Soc_list[-1]} ） ！')

            # 针对白石，楼下 没有电池信息，不报故障
            elif Serial_No_list[-1] in ('CW-MFC6000-0001', 'CW-MFC6000-0002', 'CW-MFC6000-0008', 'CW-MFC6000-0010'):

                print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}')
                print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}')

            # 其它站点只有电池2的电池，SOC低于60报故障
            else:
                print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}')
                if float(battery_2_Soc_list[-1]) >= 60:
                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}')
                else:

                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}         电池2_Soc太低异常     ！！！')
                    remark.append(f' 电池2_Soc太低异常（ {battery_2_Soc_list[-1]} ） ！')

            if Serial_No_list[-1] in (
                    'MFC6kD480012', 'MFC6kD480013', 'MFC6kD480014', 'MFC6kD480019', 'MFC6kD480020', 'MFC6kD480022',
                    'MFC6kD480023'):
                if float(Out_Remaining_Fuel_list[-1]) >= 50:
                    print("外置燃料(L)：", Out_Remaining_Fuel_list[-1])
                else:
                    print("外置燃料(L)：", Out_Remaining_Fuel_list[-1], "      外置燃料太低异常     ！！！")
                    remark.append(f' 外置燃料太低异常( {Out_Remaining_Fuel_list[-1]}(L) ) ！')
            else:
                print("外置燃料(L)：", Out_Remaining_Fuel_list[-1])

            if float(Remaining_Fuel_list[-1]) >= 15 or Serial_No_list[-1] in (
                    'CW-MFC6000-0001', 'CW-MFC6000-0002', 'CW-MFC6000-0008', 'CW-MFC6000-0010'):
                print("内置燃料(L)：", Remaining_Fuel_list[-1])
            else:
                print("内置燃料(L)：", Remaining_Fuel_list[-1], "      内置燃料太低异常     ！！！")
                remark.append(f' 内置燃料太低异常( {Remaining_Fuel_list[-1]}(L) ) ！')

            print(f'内置燃料（mm）:{In_Remaining_Fuel_mm_list[-1]}\n')

            # 如果制氢机处于‘热待机’ 或者 ‘运行’
            if A_HG_Module_status_list[-1] == 'Hot Standby' or A_HG_Module_status_list[-1] == 'Run':

                print("A制氢机状态：", A_HG_Module_status_list[-1])

                if 15 <= float(A_H2_Pressure_list[-1]) <= 25:
                    print("A氢气压力(Psi)：", A_H2_Pressure_list[-1])
                elif 15 > float(A_H2_Pressure_list[-1]):
                    print("A氢气压力(Psi)：", A_H2_Pressure_list[-1], "        氢气压力太低异常      !!!")
                    remark.append(f' A_氢气压力太低异常( {A_H2_Pressure_list[-1]}(Psi) ) ！')
                else:
                    print("A氢气压力(Psi)：", A_H2_Pressure_list[-1], "        氢气压力太高异常      !!!")
                    remark.append(f' A_氢气压力太高异常( {A_H2_Pressure_list[-1]}(Psi) ) ！')

                if 350 <= float(A_Purifier_temperature_list[-1]) <= 403:
                    print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1])
                elif 350 > float(A_Purifier_temperature_list[-1]):
                    print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1], "        提纯器温度太低异常      !!!")
                    remark.append(f' A_提纯器温度太低异常( {A_Purifier_temperature_list[-1]}(℃) ) ！')
                else:
                    print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1], "        提纯器温度太高异常      !!!")
                    remark.append(f' A_提纯器温度太高异常( {A_Purifier_temperature_list[-1]}(℃) ) ！')

                if 350 <= float(A_Reformer_Temperature_list[-1]) <= 403:
                    print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1])
                elif 350 > float(A_Reformer_Temperature_list[-1]):
                    print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1], "        提纯器温度太低异常      !!!")
                    remark.append(f' A_重整室温度太低异常( {A_Reformer_Temperature_list[-1]}(℃) ) ！')
                else:
                    print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1], "         重整室温度太高异常      !!!")
                    remark.append(f' A_重整室温度太高异常( {A_Reformer_Temperature_list[-1]}(℃) ) ！')

                if 0 <= float(A_Blower_temperature_list[-1]) <= 60:
                    print("A鼓风机温度(℃)：", A_Blower_temperature_list[-1], end='\n\n')
                else:
                    print("A鼓风机温度(℃)：", A_Blower_temperature_list[-1], "        鼓风机温度太高异常      !!!",
                          end='\n\n')
                    remark.append(f' A_鼓风机温度太高异常( {A_Blower_temperature_list[-1]}(℃) ) ！')

            # 如果制氢机处于 ‘关机’
            elif A_HG_Module_status_list[-1] == 'System Off':

                print("A制氢机状态：", A_HG_Module_status_list[-1], '           A_制氢机关机  !!！')
                remark.append(f' A_制氢机关机 ！')

                #  制氢机处于 ‘关机’ ，提纯器温度超过温度250度就报警
                if 250 <= float(A_Purifier_temperature_list[-1]):

                    if 15 <= float(A_H2_Pressure_list[-1]):
                        print("A氢气压力(Psi)：", A_H2_Pressure_list[-1])
                    else:
                        print("A氢气压力(Psi)：", A_H2_Pressure_list[-1], "        氢气压力太低异常      !!!")
                        remark.append(f' A_氢气压力太低异常( {A_H2_Pressure_list[-1]}(Psi) ) ！')

                    print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1], "        提纯器温度太高异常      !!!")
                    remark.append(f' A_提纯器温度太高异常( {A_Purifier_temperature_list[-1]}(℃) ) ！')

                    if 250 <= float(A_Reformer_Temperature_list[-1]):
                        print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1],
                              "         重整室温度太高异常      !!!")
                        remark.append(f' A_重整室温度太高异常( {A_Reformer_Temperature_list[-1]}(℃) ) ！')
                    else:
                        print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1])

                else:
                    print("A氢气压力(Psi)：", A_H2_Pressure_list[-1])
                    print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1])
                    print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1])

                print("A鼓风机温度(℃)：", A_Blower_temperature_list[-1], end='\n\n')

            # 如果制氢机处于其它状态
            else:
                print("A制氢机状态：", A_HG_Module_status_list[-1])

                print("A氢气压力(Psi)：", A_H2_Pressure_list[-1])
                print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1])
                print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1])
                print("A鼓风机温度(℃)：", A_Blower_temperature_list[-1], end='\n\n')

            if Serial_No_list[-1] == 'CW-10KW-0007':

                # 如果制氢机处于‘热待机’ 或者 ‘运行’
                if B_HG_Module_status_list[-1] == 'Hot Standby' or B_HG_Module_status_list[-1] == 'Run':
                    print("B制氢机状态：", B_HG_Module_status_list[-1])
                    if 15 <= float(B_H2_Pressure_list[-1]) <= 25:
                        print("B氢气压力(Psi)：", B_H2_Pressure_list[-1])
                    elif 15 > float(B_H2_Pressure_list[-1]):
                        print("B氢气压力(Psi)：", B_H2_Pressure_list[-1], "        氢气压力太低异常      !!!")
                        remark.append(f' B_氢气压力太低异常( {B_H2_Pressure_list[-1]}(Psi) ) ！')
                    else:
                        print("B氢气压力(Psi)：", B_H2_Pressure_list[-1], "        氢气压力太高异常      !!!")
                        remark.append(f' B_氢气压力太高异常( {B_H2_Pressure_list[-1]}(Psi) ) ！')

                    if 350 <= float(B_Purifier_temperature_list[-1]) <= 403:
                        print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1])
                    elif 350 > float(A_Purifier_temperature_list[-1]):
                        print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1], "        提纯器温度太低异常      !!!")
                        remark.append(f' B_提纯器温度太低异常( {B_Purifier_temperature_list[-1]}(℃) ) ！')
                    else:
                        print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1], "        提纯器温度太高异常      !!!")
                        remark.append(f' B_提纯器温度太高异常( {B_Purifier_temperature_list[-1]}(℃) ) ！')

                    if 350 <= float(B_Reformer_Temperature_list[-1]) <= 403:
                        print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1])
                    elif 350 > float(B_Reformer_Temperature_list[-1]):
                        print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1], "        提纯器温度太低异常      !!!")
                        remark.append(f' B_重整室温度太低异常( {B_Reformer_Temperature_list[-1]}(℃) ) ！')
                    else:
                        print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1],
                              "         重整室温度太高异常      !!!")
                        remark.append(f' B_重整室温度太高异常( {B_Reformer_Temperature_list[-1]}(℃) ) ！')

                    if 0 <= float(B_Blower_temperature_list[-1]) <= 60:
                        print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], end='\n\n')
                    else:
                        print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], "        鼓风机温度太高异常      !!!",
                              end='\n\n')
                        remark.append(f' B_鼓风机温度太高异常( {B_Blower_temperature_list[-1]}(℃) ) ！')

                # 如果制氢机处于 ‘关机’
                elif B_HG_Module_status_list[-1] == 'System Off':
                    print("B制氢机状态：", B_HG_Module_status_list[-1], '          B_制氢机关机 !!')
                    remark.append(f' B_制氢机关机 ！')

                    #  制氢机处于 ‘关机’ ，提纯器温度超过温度250度就报警
                    if 250 <= float(B_Purifier_temperature_list[-1]):

                        if 15 <= float(B_H2_Pressure_list[-1]):
                            print("B氢气压力(Psi)：", B_H2_Pressure_list[-1])
                        else:
                            print("B氢气压力(Psi)：", B_H2_Pressure_list[-1], "        氢气压力太低异常      !!!")
                            remark.append(f' B_氢气压力太低异常( {B_H2_Pressure_list[-1]}(Psi) ) ！')

                        print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1], "        提纯器温度太高异常      !!!")
                        remark.append(f' B_提纯器温度太高异常( {B_Purifier_temperature_list[-1]}(℃) ) ！')

                        if 250 <= float(B_Reformer_Temperature_list[-1]):
                            print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1],
                                  "         重整室温度太高异常      !!!")
                            remark.append(f' B_重整室温度太高异常( {B_Reformer_Temperature_list[-1]}(℃) ) ！')
                        else:
                            print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1])

                    else:
                        print("B氢气压力(Psi)：", B_H2_Pressure_list[-1])
                        print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1])
                        print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1])

                    print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], end='\n\n')
                # 如果制氢机处于其它状态
                else:
                    print("B制氢机状态：", A_HG_Module_status_list[-1])

                    print("B氢气压力(Psi)：", B_H2_Pressure_list[-1])
                    print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1])
                    print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1])
                    print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], end='\n\n')

            else:
                print("B氢气压力(Psi)：", B_H2_Pressure_list[-1])
                print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1])
                print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1])
                print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], end='\n\n')

            print("A电堆状态：", A_Stack_Module_status_list[-1])

            if 0 <= float(A_Stack_voltage_list[-1]) <= 10:
                print("A电堆电压(V)：", A_Stack_voltage_list[-1])
            else:
                print("A电堆电压(V)：", A_Stack_voltage_list[-1], "        电堆电压异常      !!!")
                remark.append(f' A_电堆电压异常( {A_Stack_voltage_list[-1]}(V) ) ！')

            if 0 <= float(A_Stack_current_list[-1]) <= 3:
                print("A电堆电流(A)：", A_Stack_current_list[-1])
            else:
                print("A电堆电流(A)：", A_Stack_current_list[-1], "         电堆电流异常      !!!")
                remark.append(f' A_电堆电流异常( {A_Stack_current_list[-1]}(A) ) ！')

            if 0 <= float(A_Stack_power_list[-1]) <= 300:
                print("A电堆功率(W)：", A_Stack_power_list[-1])
            else:
                print("A电堆功率(W)：", A_Stack_power_list[-1], "         电堆功率异常      !!!")
                remark.append(f' A_电堆功率异常( {A_Stack_power_list[-1]}(W) ) ！')

            if float(A1_Stack_temperature_list[-1]) <= 50:
                print("A1电堆堆心温度(℃)：", A1_Stack_temperature_list[-1])
            else:
                print(f"A1电堆堆心温度(℃)：{A1_Stack_temperature_list[-1]}         A1电堆堆心温度太高异常      !!!")
                remark.append(F' A1_电堆堆心温度太高异常( {A1_Stack_temperature_list[-1]}(℃) ) ！')

            if float(A2_Stack_temperature_list[-1]) <= 50:
                print("A2电堆堆心温度(℃)：", A2_Stack_temperature_list[-1])
            else:
                print(f"A2电堆堆心温度(℃)：{A2_Stack_temperature_list[-1]}         A2电堆堆心温度太高异常      !!!")
                remark.append(f' A2_电堆堆心温度太高异常( {A2_Stack_temperature_list[-1]}(℃) ) ！')

            if float(A1_Stack_top_temperature_list[-1]) <= 50:
                print("A1电堆顶部温度(℃)：", A1_Stack_top_temperature_list[-1])
            else:
                print(f"A1电堆顶部温度(℃)：{A1_Stack_top_temperature_list[-1]}         A1电堆顶部温度太高异常      !!!")
                remark.append(f' A1_电堆顶部温度太高异常( {A1_Stack_top_temperature_list[-1]}(℃) ) ！')

            if float(A2_Stack_top_temperature_list[-1]) <= 50:
                print(f"A2电堆顶部温度(℃)：{A2_Stack_top_temperature_list[-1]}\n")
            else:
                print(f"A2电堆顶部温度(℃)：{A2_Stack_top_temperature_list[-1]}         A2电堆顶部温度太高异常      !!\n")
                remark.append(f' A2_电堆顶部温度太高异常( {A2_Stack_top_temperature_list[-1]}(℃) ) ！')

            print("B电堆状态：", B_Stack_Module_status_list[-1])

            if Serial_No_list[-1] == 'CW-10KW-0007':

                if 0 <= float(B_Stack_voltage_list[-1]) <= 10:
                    print("B电堆电压(V)：", B_Stack_voltage_list[-1])
                else:
                    print("B电堆电压(V)：", B_Stack_voltage_list[-1], "        电堆电压异常      !!!")
                    remark.append(f' B_电堆电压异常( {B_Stack_voltage_list[-1]}(V) ) ！')

                if 0 <= float(B_Stack_current_list[-1]) <= 3:
                    print("B电堆电流(A)：", B_Stack_current_list[-1])
                else:
                    print("B电堆电流(A)：", B_Stack_current_list[-1], "         电堆电流异常      !!!")
                    remark.append(f' B_电堆电流异常( {B_Stack_current_list[-1]}(A) ) ！')

                if 0 <= float(B_Stack_power_list[-1]) <= 300:
                    print("B电堆功率(W)：", B_Stack_power_list[-1])
                else:
                    print("B电堆功率(W)：", B_Stack_power_list[-1], "         电堆功率异常      !!!")
                    remark.append(f' B_电堆功率异常( {B_Stack_power_list[-1]}(W) ) ！')

                if float(B_Stack_temperature_list[-1]) <= 50:
                    print("B电堆堆心温度(℃)：", B_Stack_temperature_list[-1])
                else:
                    print(f"B电堆堆心温度(℃)：{B_Stack_temperature_list[-1]}         A1电堆堆心温度太高异常      !!!")
                    remark.append(f' B_电堆堆心温度太高异常( {B_Stack_temperature_list[-1]}(℃) ) ！')

                if float(B1_Stack_top_temperature_list[-1]) <= 50:
                    print("B1电堆顶部温度(℃)：", B1_Stack_top_temperature_list[-1])
                else:
                    print(
                        f"B1电堆顶部温度(℃)：{B1_Stack_top_temperature_list[-1]}         A1电堆顶部温度太高异常      !!!")
                    remark.append(f' B1_电堆顶部温度太高异常( {B1_Stack_top_temperature_list[-1]}(℃) ) ！')

                if float(B2_Stack_top_temperature_list[-1]) <= 50:
                    print(f"B2电堆顶部温度(℃)：{B2_Stack_top_temperature_list[-1]}")
                else:
                    print(
                        f"B2电堆顶部温度(℃)：{B2_Stack_top_temperature_list[-1]}         A2电堆顶部温度太高异常      !!!")
                    remark.append(f' B2_电堆顶部温度太高异常( {B2_Stack_top_temperature_list[-1]}(℃) ) ！')

            else:
                print("B电堆电压(V)：", B_Stack_voltage_list[-1])
                print("B电堆电流(A)：", B_Stack_current_list[-1])
                print("B电堆功率(W)：", B_Stack_power_list[-1])
                print("B电堆堆心温度(℃)：", B_Stack_temperature_list[-1])
                print("B1电堆顶部温度(℃)：", B1_Stack_top_temperature_list[-1])
                print(f"B2电堆顶部温度(℃)：{B2_Stack_top_temperature_list[-1]}")


        # 如果设备处于‘开机’状态，执行以下代码
        else:

            print("设备运行状态：", System_status_list[-1], '          设备发电中 ！！ ！')
            remark.append(f' 设备发电中 ！！ ！')

            if float(Current_Voltage_list[-1]) >= 49.8 or Serial_No_list[-1] in (
                    'CW-MFC6000-0001', 'CW-MFC6000-0002', 'CW-MFC6000-0008', 'CW-MFC6000-0010'):
                print("设备母线电压(V)：", Current_Voltage_list[-1])
            else:
                print("设备母线电压(V)：", Current_Voltage_list[-1], "      设备母线电压太低异常     ！！！")
                remark.append(f' 设备母线电压太低异常（ {Current_Voltage_list[-1]}(V) ） ！')

            # 针对管委会，川岛，四川 有两块电池情况
            if Serial_No_list[-1] in ('CW-10KW-0007', 'MFC6kD480022', 'MFC6kD480023'):

                if float(battery_1_Soc_list[-1]) >= 60:
                    print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}')
                else:

                    print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}         电池1_Soc太低异常     ！！！')
                    remark.append(f' 电池1_Soc太低异常（ {battery_1_Soc_list[-1]} ） ！')

                if float(battery_2_Soc_list[-1]) >= 60:
                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}')
                else:

                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}         电池2_Soc太低异常     ！！！')
                    remark.append(f' 电池2_Soc太低异常（ {battery_2_Soc_list[-1]} ） ！')

            # 针对白石，楼下 没有电池信息，不报故障
            elif Serial_No_list[-1] in ('CW-MFC6000-0001', 'CW-MFC6000-0002', 'CW-MFC6000-0008', 'CW-MFC6000-0010'):

                print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}')
                print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}')

            # 其它站点只有电池2的电池，SOC低于60报故障
            else:
                print(f'设备电池1_Soc: {battery_1_Soc_list[-1]}')
                if float(battery_2_Soc_list[-1]) >= 60:
                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}')
                else:

                    print(f'设备电池2_Soc: {battery_2_Soc_list[-1]}         电池2_Soc太低异常     ！！！')
                    remark.append(f' 电池2_Soc太低异常（ {battery_2_Soc_list[-1]} ） ！')

            if Serial_No_list[-1] in (
                    'MFC6kD480012', 'MFC6kD480013', 'MFC6kD480014', 'MFC6kD480019', 'MFC6kD480020', 'MFC6kD480022',
                    'MFC6kD480023'):
                if float(Out_Remaining_Fuel_list[-1]) >= 50:
                    print("外置燃料(L)：", Out_Remaining_Fuel_list[-1])
                else:
                    print("外置燃料(L)：", Out_Remaining_Fuel_list[-1], "      外置燃料太低异常     ！！！")
                    remark.append(f' 外置燃料太低异常( {Out_Remaining_Fuel_list[-1]}(L) ) ！')
            else:
                print("外置燃料(L)：", Out_Remaining_Fuel_list[-1])

            if float(Remaining_Fuel_list[-1]) >= 15 or Serial_No_list[-1] in (
                    'CW-MFC6000-0001', 'CW-MFC6000-0002', 'CW-MFC6000-0008', 'CW-MFC6000-0010'):
                print("内置燃料(L)：", Remaining_Fuel_list[-1])
            else:
                print("内置燃料(L)：", Remaining_Fuel_list[-1], "      内置燃料太低异常     ！！！")
                remark.append(f' 内置燃料太低异常( {Remaining_Fuel_list[-1]}(L) ) ！')

            print(f'内置燃料（mm）:{In_Remaining_Fuel_mm_list[-1]}\n')

            print("A制氢机状态：", A_HG_Module_status_list[-1])

            if 15 <= float(A_H2_Pressure_list[-1]) <= 25:
                print("A氢气压力(Psi)：", A_H2_Pressure_list[-1])
            elif 15 > float(A_H2_Pressure_list[-1]):
                print("A氢气压力(Psi)：", A_H2_Pressure_list[-1], "        氢气压力太低异常      !!!")
                remark.append(f' A_氢气压力太低异常( {A_H2_Pressure_list[-1]}(Psi) ) ！')
            else:
                print("A氢气压力(Psi)：", A_H2_Pressure_list[-1], "        氢气压力太高异常      !!!")
                remark.append(f' A_氢气压力太高异常( {A_H2_Pressure_list[-1]}(Psi) ) ！')

            if 350 <= float(A_Purifier_temperature_list[-1]) <= 403:
                print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1])
            elif 350 > float(A_Purifier_temperature_list[-1]):
                print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1], "        提纯器温度太低异常      !!!")
                remark.append(f' A_提纯器温度太低异常( {A_Purifier_temperature_list[-1]}(℃) ) ！')
            else:
                print("A提纯器温度(℃)：", A_Purifier_temperature_list[-1], "        提纯器温度太高异常      !!!")
                remark.append(f' A_提纯器温度太高异常( {A_Purifier_temperature_list[-1]}(℃) ) ！')

            if 350 <= float(A_Reformer_Temperature_list[-1]) <= 403:
                print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1])
            elif 350 > float(A_Reformer_Temperature_list[-1]):
                print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1], "        提纯器温度太低异常      !!!")
                remark.append(f' A_重整室温度太低异常( {A_Reformer_Temperature_list[-1]}(℃) ) ！')
            else:
                print("A重整室温度(℃)：", A_Reformer_Temperature_list[-1], "         重整室温度太高异常      !!!")
                remark.append(f' A_重整室温度太高异常( {A_Reformer_Temperature_list[-1]}(℃) ) ！')

            if 0 <= float(A_Blower_temperature_list[-1]) <= 60:
                print("A鼓风机温度(℃)：", A_Blower_temperature_list[-1], end='\n\n')
            else:
                print("A鼓风机温度(℃)：", A_Blower_temperature_list[-1], "        鼓风机温度太高异常      !!!",
                      end='\n\n')
                remark.append(f' A_鼓风机温度太高异常( {A_Blower_temperature_list[-1]}(℃) ) ！')

            print("B制氢机状态：", B_HG_Module_status_list[-1])

            if Serial_No_list[-1] == 'CW-10KW-0007':

                if 15 <= float(B_H2_Pressure_list[-1]) <= 25:
                    print("B氢气压力(Psi)：", B_H2_Pressure_list[-1])
                elif 15 > float(B_H2_Pressure_list[-1]):
                    print("B氢气压力(Psi)：", B_H2_Pressure_list[-1], "        氢气压力太低异常      !!!")
                    remark.append(f' B_氢气压力太低异常( {B_H2_Pressure_list[-1]}(Psi) ) ！')
                else:
                    print("B氢气压力(Psi)：", B_H2_Pressure_list[-1], "        氢气压力太高异常      !!!")
                    remark.append(f' B_氢气压力太高异常( {B_H2_Pressure_list[-1]}(Psi) ) ！')

                if 350 <= float(B_Purifier_temperature_list[-1]) <= 403:
                    print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1])
                elif 350 > float(A_Purifier_temperature_list[-1]):
                    print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1], "        提纯器温度太低异常      !!!")
                    remark.append(f' B_提纯器温度太低异常( {B_Purifier_temperature_list[-1]}(℃) ) ！')
                else:
                    print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1], "        提纯器温度太高异常      !!!")
                    remark.append(f' B_提纯器温度太高异常( {B_Purifier_temperature_list[-1]}(℃) ) ！')

                if 350 <= float(B_Reformer_Temperature_list[-1]) <= 403:
                    print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1])
                elif 350 > float(B_Reformer_Temperature_list[-1]):
                    print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1], "        提纯器温度太低异常      !!!")
                    remark.append(f' B_重整室温度太低异常( {B_Reformer_Temperature_list[-1]}(℃) ) ！')
                else:
                    print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1], "         重整室温度太高异常      !!!")
                    remark.append(f' B_重整室温度太高异常( {B_Reformer_Temperature_list[-1]}(℃) ) ！')

                if 0 <= float(B_Blower_temperature_list[-1]) <= 60:
                    print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], end='\n\n')
                else:
                    print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], "        鼓风机温度太高异常      !!!",
                          end='\n\n')
                    remark.append(f' B_鼓风机温度太高异常( {B_Blower_temperature_list[-1]}(℃) ) ！')
            else:
                print("B氢气压力(Psi)：", B_H2_Pressure_list[-1])
                print("B提纯器温度(℃)：", B_Purifier_temperature_list[-1])
                print("B重整室温度(℃)：", B_Reformer_Temperature_list[-1])
                print("B鼓风机温度(℃)：", B_Blower_temperature_list[-1], end='\n\n')

            print("A电堆状态：", A_Stack_Module_status_list[-1])
            print("A电堆电压(V)：", A_Stack_voltage_list[-1])
            remark.append(f' A_电堆电压：{A_Stack_voltage_list[-1]}(V)')

            print("A电堆电流(A)：", A_Stack_current_list[-1])
            print("A电堆功率(W)：", A_Stack_power_list[-1])
            remark.append(f' A_电堆功率：{A_Stack_power_list[-1]}(W) ')

            print("A1电堆堆心温度(℃)：", A1_Stack_temperature_list[-1])
            print("A2电堆堆心温度(℃)：", A2_Stack_temperature_list[-1])
            print("A1电堆顶部温度(℃)：", A1_Stack_top_temperature_list[-1])
            print(f"A2电堆顶部温度(℃)：{A2_Stack_top_temperature_list[-1]}\n")

            print("B电堆状态：", B_Stack_Module_status_list[-1])
            print("B电堆电压(V)：", B_Stack_voltage_list[-1])
            print("B电堆电流(A)：", B_Stack_current_list[-1])
            print("B电堆功率(W)：", B_Stack_power_list[-1])
            print("B电堆堆心温度(℃)：", B_Stack_temperature_list[-1])
            print("B1电堆顶部温度(℃)：", B1_Stack_top_temperature_list[-1])
            print(f"B2电堆顶部温度(℃)：{B2_Stack_top_temperature_list[-1]}\n")

        if len(remark) > 0:
            #  将所有故障加入remark_set，最后生成excel表格的时候提取出来
            #   ' , '.join(remark)会取出remark列表中的每个元素，并将它们用一个逗号连接起来，形成一个单独的字符串。
            remark_set.append(f'"{machine_name_list[-1]}"设备状态:' + ' , '.join(remark))
            print(f'\n{remark_set[-1]}')

        else:
            print('\n设备没有出现异常故障     ... ... ...')
            remark_set.append(''.join(remark))

        remark.clear()

        print(f'\n=================        =================\n')

    except Exception as e:
        # print(f"发生错误：{e}")
        print(f'\n=================        =================\n')


def page1_data_processing(driver, wait, row_key):
    """
    第一页设备列表，数据采集
    :param driver:
    :param wait:
    :param row_key:
    :return:
    """
    # 跳转到指定设备页面
    JMP(driver, wait, row_key)
    # 读取设备页面指定数据
    data_processing(driver, wait)


def page2_data_processing(driver, wait, row_key):
    """
    第二页设备列表，数据采集
    :param driver:
    :param wait:
    :param row_key:
    :return:
    """
    # 跳转到第二页去
    jum_page_2(driver, wait)
    # 使用 JavaScript 滚动到页面顶部
    # driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(2)
    click_find_target_Details(driver, row_key)
    # 读取设备页面指定数据
    data_processing(driver, wait)


def print_array_length():
    print(f'日期时间 长度: {len(time_localtime_list)}')
    print(f'设备编号 长度: {len(Serial_No_list)}')
    print(f'设备名称 长度: {len(machine_name_list)}')
    print(f'设备网络状态 长度: {len(network_state_list)}')
    print(f'设备运行状态 长度: {len(System_status_list)}')
    print(f'设备母线电压(V) 长度: {len(Current_Voltage_list)}')
    print(f'电池1_Soc 长度: {len(battery_1_Soc_list)}')
    print(f'电池2_Soc 长度: {len(battery_2_Soc_list)}')
    print(f'外置燃料(L) 长度: {len(Out_Remaining_Fuel_list)}')
    print(f'内置燃料(L) 长度: {len(Remaining_Fuel_list)}')
    print(f'内置燃料(mm) 长度: {len(In_Remaining_Fuel_mm_list)}')
    print(f'A_制氢机状态 长度: {len(A_HG_Module_status_list)}')
    print(f'A_氢气压力(Psi) 长度: {len(A_H2_Pressure_list)}')
    print(f'A_鼓风机温度(℃) 长度: {len(A_Blower_temperature_list)}')
    print(f'A_提纯器温度(℃) 长度: {len(A_Purifier_temperature_list)}')
    print(f'A_重整室温度(℃) 长度: {len(A_Reformer_Temperature_list)}')
    print(f'B_制氢机状态 长度: {len(B_HG_Module_status_list)}')
    print(f'B_氢气压力(Psi) 长度: {len(B_H2_Pressure_list)}')
    print(f'B_鼓风机温度(℃) 长度: {len(B_Blower_temperature_list)}')
    print(f'B_提纯器温度(℃) 长度: {len(B_Purifier_temperature_list)}')
    print(f'B_重整室温度(℃) 长度: {len(B_Reformer_Temperature_list)}')
    print(f'A_电堆状态 长度: {len(A_Stack_Module_status_list)}')
    print(f'A_电堆电压(V) 长度: {len(A_Stack_voltage_list)}')
    print(f'A_电堆电流(A) 长度: {len(A_Stack_current_list)}')
    print(f'A_电堆功率(W) 长度: {len(A_Stack_power_list)}')
    print(f'A1_电堆堆心温度(℃) 长度: {len(A1_Stack_temperature_list)}')
    print(f'A2_电堆堆心温度(℃) 长度: {len(A2_Stack_temperature_list)}')
    print(f'A1_电堆顶部温度(℃) 长度: {len(A1_Stack_top_temperature_list)}')
    print(f'A2_电堆顶部温度(℃) 长度: {len(A2_Stack_top_temperature_list)}')
    print(f'B_电堆状态 长度: {len(B_Stack_Module_status_list)}')
    print(f'B_电堆电压(V) 长度: {len(B_Stack_voltage_list)}')
    print(f'B_电堆电流(A) 长度: {len(B_Stack_current_list)}')
    print(f'B_电堆功率(W) 长度: {len(B_Stack_power_list)}')
    print(f'B_电堆堆心温度(℃) 长度: {len(B_Stack_temperature_list)}')
    print(f'B1_电堆顶部温度(℃) 长度: {len(B1_Stack_top_temperature_list)}')
    print(f'B2_电堆顶部温度(℃) 长度: {len(B2_Stack_top_temperature_list)}')
    print(f'备注 长度: {len(remark_set)}\n\n')


def excelfile_save(file_path):
    new_pd = pd.DataFrame({

        "日期时间": time_localtime_list,
        "设备编号": Serial_No_list,
        "设备名称": machine_name_list,
        "设备网络状态": network_state_list,
        "设备运行状态": System_status_list,
        "设备母线电压(V)": Current_Voltage_list,
        '电池1_Soc': battery_1_Soc_list,
        '电池2_Soc': battery_2_Soc_list,
        "外置燃料(L)": Out_Remaining_Fuel_list,
        "内置燃料(L)": Remaining_Fuel_list,
        "内置燃料(mm)": In_Remaining_Fuel_mm_list,

        "A_制氢机状态": A_HG_Module_status_list,
        'A_氢气压力(Psi)': A_H2_Pressure_list,
        'A_鼓风机温度(℃)': A_Blower_temperature_list,
        'A_提纯器温度(℃)': A_Purifier_temperature_list,
        'A_重整室温度(℃)': A_Reformer_Temperature_list,

        "B_制氢机状态": B_HG_Module_status_list,
        'B_氢气压力(Psi)': B_H2_Pressure_list,
        'B_鼓风机温度(℃)': B_Blower_temperature_list,
        'B_提纯器温度(℃)': B_Purifier_temperature_list,
        'B_重整室温度(℃)': B_Reformer_Temperature_list,

        'A_电堆状态': A_Stack_Module_status_list,
        'A_电堆电压(V)': A_Stack_voltage_list,
        'A_电堆电流(A)': A_Stack_current_list,
        'A_电堆功率(W)': A_Stack_power_list,
        'A1_电堆堆心温度(℃)': A1_Stack_temperature_list,
        'A2_电堆堆心温度(℃)': A2_Stack_temperature_list,
        'A1_电堆顶部温度(℃)': A1_Stack_top_temperature_list,
        'A2_电堆顶部温度(℃)': A2_Stack_top_temperature_list,

        'B_电堆状态': B_Stack_Module_status_list,
        'B_电堆电压(V)': B_Stack_voltage_list,
        'B_电堆电流(A)': B_Stack_current_list,
        'B_电堆功率(W)': B_Stack_power_list,
        'B_电堆堆心温度(℃)': B_Stack_temperature_list,
        'B1_电堆顶部温度(℃)': B1_Stack_top_temperature_list,
        'B2_电堆顶部温度(℃)': B2_Stack_top_temperature_list,

        '备注': remark_set

    })
    # 判断是否有相同的文件存在，有的话改名
    if os.path.exists(file_path):
        # 文件存在，生成新的文件名
        base_name, extension = os.path.splitext(file_path)
        counter = 1
        new_file_path = f"{base_name}_{counter}{extension}"
        while os.path.exists(new_file_path):
            counter += 1
            new_file_path = f"{base_name}_{counter}{extension}"
        file_path = new_file_path  # 更新文件路径

    new_pd.to_excel(file_path, index=False)
    # 检查文件是否已存在
    # 保存DataFrame到Excel
    # 打开现有的Excel文件
    workbook = openpyxl.load_workbook(file_path)
    # 选择第一个工作表
    sheet = workbook.active
    # 设置第一行的行高
    sheet.row_dimensions[1].height = 50
    # 设置第一列和第二列的宽度为 25
    sheet.column_dimensions['A'].width = 21  # 第一列
    sheet.column_dimensions['B'].width = 21  # 第二列
    sheet.column_dimensions['C'].width = 25  # 第三列
    # 设置其余列的宽度为 10
    for col in sheet.columns:
        if col[0].column_letter not in ['A', 'B', 'C']:
            sheet.column_dimensions[col[0].column_letter].width = 15
    # 遍历第一行的所有单元格，并为每个单元格对象同时设置自动换行、水平居中和垂直居中。
    for cell in sheet[1]:
        cell_obj = cell
        cell_obj.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')

    workbook.save(file_path)
    print(f'文件保存成功  ！ 保存路径：{file_path}')


# 主函数入口 ，调用这个函数程序从这里开始执行
def main():
    # 目标网页URL
    url = 'http://47.113.86.137:880/#/device/detail?serialNo=CW-0002'
    # 设置WebDriver路径
    driver_path = 'C:/Users/11016/AppData/Local/Google/Chrome/Application/chromedriver.exe'  # 例如 ChromeDriver 的路径
    # 设备文件保存路径和保存文件名称
    file_path = 'D:/爬虫数据/网页采集数据.xlsx'
    # 网页登录账号
    loginName = 'admin'
    # 网页登录密码
    passWord = 'GJM456789'
    # 循环体里面程序暂停的时间
    sleeptime = 0  # 程序暂停运行,时间单位：min

    # 初始化函数 ，打开谷歌浏览器，输入账号密码，登录网页首页
    driver, wait = Program_Init(driver_path, url, loginName, passWord)

    print('\n~~~~~~~~~    开始爬虫     ~~~~~~~~~~\n')
    for i in range(0, 1):
        page1_data_processing(driver, wait, 3)  # 管委会                   row_key=3
        page1_data_processing(driver, wait, 14)  # 楼下机房1号机             row_key=14
        page1_data_processing(driver, wait, 15)  # 楼下机房2号机             row_key=15

        page2_data_processing(driver, wait, 0)  # 白石机房1号机              row_key=0
        # page2_data_processing(driver, wait, 1)  # 龙榜机房                  row_key=1
        page2_data_processing(driver, wait, 2)  # 白石机房2号机              row_key=2
        page2_data_processing(driver, wait, 8)  # 洋美                     row_key=8
        page2_data_processing(driver, wait, 9)  # 红关                     row_key=9
        page2_data_processing(driver, wait, 10)  # 墩寨                     row_key=10
        page2_data_processing(driver, wait, 11)  # 谭溪                     row_key=11
        page2_data_processing(driver, wait, 12)  # 华安                     row_key=12
        page2_data_processing(driver, wait, 13)  # 新美                     row_key=13
        page2_data_processing(driver, wait, 14)  # 升平                     row_key=14
        page2_data_processing(driver, wait, 15)  # 平石                     row_key=15
        page2_data_processing(driver, wait, 16)  # 三联                     row_key=16
        page2_data_processing(driver, wait, 17)  # 三堡                     row_key=17
        page2_data_processing(driver, wait, 18)  # 上川岛长堤                row_key=18
        page2_data_processing(driver, wait, 19)  # 四川江油太平唐僧           row_key=19

        excelfile_save(file_path)

        print(f'第 {i} 次系统进入休眠 ， 休眠时长：{sleeptime} min')
        time.sleep(60 * sleeptime)

    print('\n~~~~~~~~~    结束爬虫     ~~~~~~~~~~')
    driver.quit()


if __name__ == "__main__":
    main()
