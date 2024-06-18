from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# 设置WebDriver路径
driver_path = 'C:/Users/11016/AppData/Local/Google/Chrome/Application/chromedriver.exe'  # 例如 ChromeDriver 的路径

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
# 使用提供的XPath定位元素
xpath_expression_1 = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[1]/div[1]"
# 使用XPath定位元素
xpath_expression_2 = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div[3]/div[1]"

xpath_expression_3 = '/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div[1]/div[1]'

try:

    # 等待元素出现
    element_1 = wait.until(EC.presence_of_element_located((By.XPATH, xpath_expression_1)))
    element_2 = wait.until(EC.presence_of_element_located((By.XPATH, xpath_expression_2)))
    element_3 = wait.until(EC.presence_of_element_located((By.XPATH, xpath_expression_3)))

    H2_Pressure = element_1.text.split(':', 1)[-1].strip()
    print("找到的文本内容：", H2_Pressure)
    # 获取并打印元素的文本内容
    print("找到的文本内容：", element_1.text)
    # 获取并打印元素的文本内容
    print("找到的文本内容：", element_2.text)
    print("找到的文本内容：", element_3.text)


except Exception as e:
    print("发生错误：", e)

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
