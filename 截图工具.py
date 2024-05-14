import time
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import ImageGrab
import os
import threading
from datetime import datetime

# 全局变量，用于控制截图线程的运行和停止
stop_event = threading.Event()
# 全局变量，用于跟踪已截图的数量
current_screenshot_count = 0

# 定义截图的函数
def take_screenshot(path):
    # 使用global关键字来声明我们将会修改全局变量
    global current_screenshot_count
    # 获取当前时间并格式化为字符串
    current_time = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    # 创建截图的文件名
    filename = f"截图_{current_time}.jpg"
    # 完整的文件保存路径
    fullpath = os.path.join(path, filename)
    # 截取屏幕并保存图像
    img = ImageGrab.grab()
    img.save(fullpath)
    # 增加截图计数
    current_screenshot_count += 1  
    # 更新GUI上的截图数量显示
    update_screenshot_count()
    # 返回保存截图的完整路径
    return fullpath

# 更新截图数量显示的函数
def update_screenshot_count():
    # 更新截图数量标签的文本
    screenshot_count_label.config(text=f"已截图: {current_screenshot_count}")

# 定义截图线程的工作函数
def screenshot_thread(interval, total, path):
    # 使用global关键字来声明我们将会修改全局变量
    global stop_event
    # 循环指定次数进行截图
    for i in range(total):
        # 检查是否接收到停止信号
        if stop_event.is_set():
            break  # 如果是，则退出线程
        # 执行截图
        take_screenshot(path)
        # 等待一段时间，即设置的间隔时间
        time.sleep(interval)

    # 所有截图完成后，弹出消息告知用户
    messagebox.showinfo("截图完成", f"所有截图已保存至指定文件夹：\n{path}")

# 定义开始截图的函数
def start_screenshot_timer():
    # 使用global关键字来声明我们将会修改全局变量
    global stop_event
    try:
        # 从输入框获取间隔时间和总截图数量
        interval = int(interval_entry.get())
        total = int(total_entry.get())
        # 获取用户设置的保存路径
        path = path_entry.get()
        # 检查设置的保存路径是否存在，如果不存在，则创建
        if not os.path.exists(path):
            os.makedirs(path)

        # 重置停止信号，以便线程可以开始
        stop_event.clear()
        # 创建一个线程用于执行截图
        thread = threading.Thread(target=screenshot_thread, args=(interval, total, path))
        # 启动线程
        thread.start()
    except ValueError:
        # 如果输入的不是有效的整数，弹出错误消息
        messagebox.showerror("错误", "请输入有效的数字。")

# 定义停止截图的函数
def stop_screenshot_timer():
    # 使用global关键字来声明我们将会修改全局变量
    global stop_event
    # 设置停止信号，这将导致截图线程在下次循环时停止
    stop_event.set()  
    # 弹出消息告知用户截图已停止
    messagebox.showinfo("截图已停止", "截图已停止。")

# 创建主窗口
root = tk.Tk()
root.title("老王出品--自动截图工具--V1.0")

# 创建保存路径的输入框和按钮
path_label = tk.Label(root, text="保存路径: ")
path_entry = tk.Entry(root, width=30)
browse_button = tk.Button(root, text="浏览", command=lambda: path_entry.insert(0, filedialog.askdirectory()))

# 创建间隔时间和总截图数量的输入框
interval_label = tk.Label(root, text="间隔时间(秒): ")
interval_entry = tk.Entry(root, width=30)
total_label = tk.Label(root, text="总截图数: ")
total_entry = tk.Entry(root, width=30)

# 创建开始截图的按钮
start_button = tk.Button(root, text="开始截图", command=start_screenshot_timer)
# 创建停止截图的按钮
stop_button = tk.Button(root, text="停止截图", command=stop_screenshot_timer)

# 创建显示当前截图数量的标签
screenshot_count_label = tk.Label(root, text=f"已截图: {current_screenshot_count}")

# 使用grid方法布局
path_label.grid(row=0, column=0, sticky="e")  # 设置标签靠东对齐
path_entry.grid(row=0, column=1)  # 输入框占据第二列
browse_button.grid(row=0, column=2)  # 按钮占据第三列
interval_label.grid(row=1, column=0, sticky="e")  # 设置标签靠东对齐
interval_entry.grid(row=1, column=1)  # 输入框占据第二列
total_label.grid(row=2, column=0, sticky="e")  # 设置标签靠东对齐
total_entry.grid(row=2, column=1)  # 输入框占据第二列
start_button.grid(row=3, column=0, columnspan=2)  # 按钮占据第一列和第二列
stop_button.grid(row=3, column=2)  # 按钮占据第三列
screenshot_count_label.grid(row=4, column=1, columnspan=3, pady=10)  # 标签占据第四行

# 运行主事件循环
root.mainloop()
