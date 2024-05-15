import time  # 导入时间模块，用于时间相关的操作
import tkinter as tk  # 导入Tkinter库，用于创建图形用户界面
from tkinter import messagebox, filedialog  # 导入Tkinter的消息框和文件对话框模块
from PIL import ImageGrab  # 导入PIL库中的ImageGrab模块，用于截图
import os  # 导入操作系统模块，用于文件路径和操作
import threading  # 导入线程模块，用于多线程操作
from datetime import datetime  # 导入日期时间模块，用于获取当前时间

# 作者：老王出品   时间：2024.5.15  版本：V1.2


count = []  # 初始化一个用于计数的列表

# 定义截图函数，保存截图到指定路径，并更新截图数量
def take_screenshot(path, count):
    current_time = datetime.now().strftime("%Y-%m-%d %H-%M-%S")  # 获取当前时间格式化为字符串
    filename = f"截图_{current_time}.jpg"  # 构造文件名
    fullpath = os.path.join(path, filename)  # 获取完整的文件保存路径
    img = ImageGrab.grab()  # 截图整个屏幕
    img.save(fullpath)  # 保存截图到文件
    count[0] += 1  # 截图计数加1
    update_screenshot_count(count)  # 更新截图数量显示

# 定义截图线程函数，用于在指定的时间间隔和总次数内进行截图
def screenshot_thread(stop_event, interval, total, path, count):
    for i in range(total):  # 循环总次数
        if stop_event.is_set():  # 如果停止事件被设置，则退出循环
            break
        take_screenshot(path, count)  # 执行截图

        # 等待间隔时间，并检查停止事件状态
        start_time = time.time()  # 记录当前时间
        while time.time() - start_time < interval:  # 等待直到达到指定的时间间隔
            if stop_event.is_set():  # 如果停止事件被设置，则退出等待
                break
            time.sleep(0.5)  # 每隔0.5秒检查一次停止事件

    # 截图完成后显示消息框
    messagebox.showinfo("截图完成", f"所有截图已保存至指定文件夹：\n{path}")

# 定义开始截图计时器的函数
def start_screenshot_timer(stop_event, interval_entry, total_entry, path_entry):
    global count  # 全局变量声明
    try:
        interval = int(interval_entry.get())  # 获取输入的时间间隔
        total = int(total_entry.get())  # 获取输入的总截图数
        path = path_entry.get()  # 获取输入的保存路径
        if not os.path.exists(path):  # 如果路径不存在，则创建
            os.makedirs(path)
        stop_event.clear()  # 清除停止事件
        count = [0]  # 初始化截图计数
        update_screenshot_count(count)  # 更新截图计数显示
        # 创建并启动截图线程
        thread = threading.Thread(target=screenshot_thread, args=(stop_event, interval, total, path, [0]))
        count_label["text"] = f"已截图: 0"  # 更新截图计数显示
        thread.start()  # 启动线程

    except ValueError:  # 如果输入不是有效的数字，则显示错误消息
        messagebox.showerror("错误", "请输入有效的数字。")

# 定义停止截图计时器的函数
def stop_screenshot_timer(stop_event):
    stop_event.set()  # 设置停止事件
    messagebox.showinfo("截图已停止", "截图已停止。")  # 显示消息框

# 定义更新截图计数的函数
def update_screenshot_count(count):
    count_label.config(text=f"已截图: {count[0]}")  # 更新标签显示的截图数量

# 创建主窗口
root = tk.Tk()
root.title("老王出品--自动截图工具--V1.2")  # 设置窗口标题

# 输入框和按钮
path_entry = tk.Entry(root,width=30)
interval_entry = tk.Entry(root,width=30)
total_entry = tk.Entry(root,width=30)

browse_button = tk.Button(root, text="浏览", command=lambda: path_entry.insert(0, filedialog.askdirectory()))
start_button = tk.Button(root, text="开始截图",
                         command=lambda: start_screenshot_timer(stop_event, interval_entry, total_entry, path_entry))
stop_button = tk.Button(root, text="停止截图", command=lambda: stop_screenshot_timer(stop_event))
count_label = tk.Label(root, text=f"已截图: 0")

# 布局
path_label = tk.Label(root, text="保存路径: ")
interval_label = tk.Label(root, text="间隔时间(秒): ")
total_label = tk.Label(root, text="总截图数: ")

path_label.grid(row=0, column=0, sticky="e")
path_entry.grid(row=0, column=1)
browse_button.grid(row=0, column=2)
interval_label.grid(row=1, column=0, sticky="e")
interval_entry.grid(row=1, column=1)
total_label.grid(row=2, column=0, sticky="e")
total_entry.grid(row=2, column=1)
# 将开始截图按钮和停止截图按钮以及计数标签放在同一行
start_button.grid(row=3, column=0)  # 调整开始截图按钮的列跨度为1
stop_button.grid(row=3, column=1)  # 调整停止截图按钮的列位置和跨度
count_label.grid(row=3, column=2, pady=10,padx=(0,20) ) # 调整计数标签的列位置

# 运行主事件循环
stop_event = threading.Event()
root.mainloop()
