from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ================================================= #
# 版本更新：2024_6_18   更新时间2024.6.18
# 网页爬虫 COWIN 数据，目前测试抓取’5G汇聚机房1‘ 的数据 ’记录时间，设备编号（Serial_No），设备名称（Remark），氢气压力（H2_Pressure）
# 提纯器温度（Purifier_temperature），重整室温度（Reformer_Temperature），鼓风机温度（Blower_temperature），电堆电压（Stack_voltage）
# 电堆电流（Stack_current），电堆功率（Stack_power），电堆温度（Stack_temperature）‘
#


# ================================================= #





time_localtime_list = []
Serial_No_list = []
Remark_list = []
H2_Pressure_list = []
Purifier_temperature_list = []
Reformer_Temperature_list = []
Blower_temperature_list = []
Stack_voltage_list = []
Stack_temperature_list = []
Stack_current_list = []
Stack_power_list = []


def Program_Init():
    # 设置WebDriver路径
    driver_path = 'C:/Users/FCK/AppData/Local/Google/Chrome/Application/chromedriver.exe'  # 例如 ChromeDriver 的路径

    # 创建Service对象，指定ChromeDriver路径
    service = Service(executable_path=driver_path)

    # 使用Service对象作为服务启动Chrome
    driver = webdriver.Chrome(service=service)

    # 目标网页URL
    url = 'http://47.113.86.137:880/#/device/detail?serialNo=CW-0002'

    # 使用Selenium打开网页
    driver.get(url)

    # 等待登录页面加载完成
    wait = WebDriverWait(driver, 10)

    # 定位账号和密码输入框，以及登录按钮，并输入账号密码
    username_input = wait.until(EC.presence_of_element_located((By.ID, 'loginName')))
    password_input = wait.until(EC.presence_of_element_located((By.ID, 'password')))
    # 使用CSS选择器定位登录按钮
    login_button = driver.find_element(By.CSS_SELECTOR, 'button.login-button.ant-btn.ant-btn-primary.ant-btn-lg')

    username_input.send_keys('admin')
    password_input.send_keys('GJM456789')

    # 提交登录信息
    login_button.click()
    return driver, wait


def JMP(driver, wait):
    # 等待下拉菜单标题加载完成
    # wait = WebDriverWait(driver, 10)
    # 使用更具体的CSS选择器，确保选中的是可点击的元素
    submenu_title = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-menu-submenu-title')))

    # 点击下拉菜单标题以展开菜单
    submenu_title.click()

    # 使用XPath等待“Equipment List”菜单项变得可见
    # 这里假设下拉菜单展开后，包含Equipment List文本的<a>标签会直接成为可见元素
    wait.until(EC.visibility_of_element_located((By.XPATH, '//a/span[text()="Equipment List"]')))

    # 定位并点击“Equipment List”列表项
    # 如果菜单项是一个<a>标签包裹<span>，确保XPath正确地定位到这个<a>标签
    equipment_list_item = wait.until(EC.element_to_be_clickable((By.XPATH, '//a/span[text()="Equipment List"]')))
    # 使用 JavaScript 执行点击操作

    driver.execute_script("arguments[0].click();", equipment_list_item)

    # 等待表格体元素加载完成
    wait = WebDriverWait(driver, 10)
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

    # 目标表格行的 data-row-key
    row_key = '14'

    # 等待表格行元素加载完成

    row_css_selector = f".ant-table-row[data-row-key='{row_key}']"
    row_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, row_css_selector)))

    # 使用 JavaScript 滚动到表格行，并使其位于视窗底部
    # 这里我们传递了 true 作为第二个参数给 scrollIntoView 方法
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'end'});", row_element)

    # 定位到 <tbody> 元素
    tbody_selector = ".ant-table-tbody"
    tbody_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, tbody_selector)))

    # 在 <tbody> 中找到具有特定 data-row-key 的 <tr> 元素
    target_row_css_selector = f".ant-table-row[data-row-key='{row_key}']"
    # 正确使用 find_element 在 <tbody> 的上下文中查找 <tr>
    target_row_element = tbody_element.find_element(By.CSS_SELECTOR, target_row_css_selector)

    # 在找到的 <tr> 元素中定位到 "Details" 链接
    # 假设 "Details" 链接具有 data-v-11b2bf7e 属性
    details_link_selector = "a[data-v-11b2bf7e]"
    details_link = target_row_element.find_element(By.CSS_SELECTOR, details_link_selector)

    # 点击 "Details"
    details_link.click()

    # 使用 JavaScript 滚动到页面顶部
    driver.execute_script("window.scrollTo(0, 0);")

    # 这里添加10秒的等待时间
    time.sleep(5)


def data_processing(driver, wait):
    #   设备编号
    Serial_No_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/span[1]/span'
    #   设备名称
    Remark_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div[2]/div[1]'
    #   重整室温度
    Reformer_Temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[3]/div[2]'
    #   鼓风机温度
    Blower_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[1]/div[3]'
    # 提纯器温度。使用XPath定位元素
    Purifier_temperature_XPath = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[3]/div[1]"
    # 缓冲罐氢气压力。使用提供的XPath定位元素
    H2_Pressure_XPath = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[1]/div[1]"
    #   电堆电压
    Stack_voltage_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]'
    #   电堆温度
    Stack_temperature_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[2]/div[1]'
    #   电堆电流
    Stack_current_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[2]'
    #   电堆功率
    Stack_power_XPath = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[3]'

    try:
        time_localtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

        Serial_No = wait.until(EC.presence_of_element_located((By.XPATH, Serial_No_XPath))).text.split(':', 1)[
            -1].strip()
        Remark_temp = wait.until(EC.presence_of_element_located((By.XPATH, Remark_XPath))).text.split(':', 1)[1].strip()
        Remark = Remark_temp.split(' ', 1)[0]
        # 等待元素出现
        H2_Pressure = wait.until(EC.presence_of_element_located((By.XPATH, H2_Pressure_XPath))).text.split(':', 1)[
            -1].strip()
        Purifier_temperature = \
            wait.until(EC.presence_of_element_located((By.XPATH, Purifier_temperature_XPath))).text.split(':', 1)[
                -1].strip()
        Reformer_Temperature = \
            wait.until(EC.presence_of_element_located((By.XPATH, Reformer_Temperature_XPath))).text.split(':', 1)[
                -1].strip()
        Blower_temperature = \
            wait.until(EC.presence_of_element_located((By.XPATH, Blower_temperature_XPath))).text.split(':', 1)[
                -1].strip()
        Stack_voltage = wait.until(EC.presence_of_element_located((By.XPATH, Stack_voltage_XPath))).text.split(':', 1)[
            -1].strip()
        Stack_temperature = \
            wait.until(EC.presence_of_element_located((By.XPATH, Stack_temperature_XPath))).text.split(':', 1)[
                -1].strip()
        Stack_current = wait.until(EC.presence_of_element_located((By.XPATH, Stack_current_XPath))).text.split(':', 1)[
            -1].strip()
        Stack_power = wait.until(EC.presence_of_element_located((By.XPATH, Stack_power_XPath))).text.split(':', 1)[
            -1].strip()

        time_localtime_list.append(time_localtime)
        Serial_No_list.append(Serial_No)
        Remark_list.append(Remark)
        H2_Pressure_list.append(round(float(H2_Pressure), 2))
        Purifier_temperature_list.append(round(float(Purifier_temperature), 2))
        Reformer_Temperature_list.append(round(float(Reformer_Temperature), 2))
        Blower_temperature_list.append(round(float(Blower_temperature), 2))
        Stack_voltage_list.append(round(float(Stack_voltage), 2))
        Stack_temperature_list.append(round(float(Stack_temperature), 2))
        Stack_current_list.append(round(float(Stack_current), 2))
        Stack_power_list.append(round(float(Stack_power), 2))

        print("记录时间：", time_localtime_list[-1])

        print("设备编号：", Serial_No_list[-1])

        print("设备名称：", Remark_list[-1])

        if 15 <= float(H2_Pressure_list[-1]) <= 25:
            print("氢气压力：", H2_Pressure_list[-1])
        else:
            print("氢气压力：", H2_Pressure_list[-1], "        氢气压力异常      !!!")

        if 350 <= float(Purifier_temperature_list[-1]) <= 403:
            print("提纯器温度：", Purifier_temperature_list[-1])
        else:
            print("提纯器温度：", Purifier_temperature_list[-1], "        提纯器温度异常      !!!")

        if 350 <= float(Reformer_Temperature_list[-1]) <= 403:
            print("重整室温度：", Reformer_Temperature_list[-1])
        else:
            print("重整室温度：", Reformer_Temperature_list[-1], "         重整室温度异常      !!!")

        if 0 <= float(Blower_temperature_list[-1]) <= 60:
            print("鼓风机温度：", Blower_temperature_list[-1])
        else:
            print("鼓风机温度：", Blower_temperature_list[-1], "        鼓风机温度异常      !!!")

        if 0 <= float(Stack_voltage_list[-1]) <= 10:
            print("电堆电压：", Stack_voltage_list[-1])
        else:
            print("电堆电压：", Stack_voltage_list[-1], "        电堆电压异常      !!!")

        if 0 <= float(Stack_current_list[-1]) <= 3:
            print("电堆电流：", Stack_current_list[-1])
        else:
            print("电堆电流：", Stack_current_list[-1], "         电堆电流异常      !!!")

        if 0 <= float(Stack_power_list[-1]) <= 300:
            print("电堆功率：", Stack_power_list[-1])
        else:
            print("电堆功率：", Stack_power_list[-1], "         电堆功率异常      !!!")

        if float(Stack_temperature_list[-1]) <= 50:
            print("电堆温度：", Stack_temperature_list[-1])
            print(f'\n=================        =================\n')
        else:
            print(f"电堆温度：{Stack_temperature_list[-1]}         电堆温度异常      !!!")
            print(f'\n=================        =================\n')

    except Exception as e:
        print(f"发生错误：{e}")
        print(f'\n=================        =================\n')


# 主函数入口
def main():
    driver, wait = Program_Init()
    JMP(driver, wait)

    count = 3
    while count:
        data_processing(driver, wait)
        time.sleep(10)
        count -= 1

    # 等待页面加载完成，可能需要根据实际情况调整等待条件
    try:
        WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.ID, "某个元素的ID"))  # 根据页面元素调整
        )
        # 或者等待页面的某个特定元素加载完成
        # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "某个元素的类名")))
    except TimeoutException:
        print("页面加载超时")

    # 获取页面源代码
    html_content = driver.page_source

    # 打印或处理html_content
    print(html_content)

    # 完成后关闭浏览器
    driver.quit()


if __name__ == "__main__":
    main()
