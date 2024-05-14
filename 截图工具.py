import time
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import ImageGrab
import os
import threading
from datetime import datetime

# 全局变量，用于控制截图线程
stop_event = threading.Event()
# 全局变量，用于跟踪已截图张数
current_screenshot_count = 0


# 定义截图的函数
def take_screenshot(path):
    global current_screenshot_count
    current_time = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    filename = f"截图_{current_time}.jpg"
    fullpath = os.path.join(path, filename)
    img = ImageGrab.grab()
    img.save(fullpath)
    current_screenshot_count += 1  # 更新截图数量
    update_screenshot_count()  # 更新GUI上的截图数量显示
    return fullpath


# 更新截图数量显示的函数
def update_screenshot_count():
    screenshot_count_label.config(text=f"已截图: {current_screenshot_count}")


# 定义截图线程的工作函数
def screenshot_thread(interval, total, path):
    global stop_event
    for i in range(total):
        if stop_event.is_set():
            break  # 如果接收到停止信号，退出线程
        take_screenshot(path)
        time.sleep(interval)

    # 所有截图完成后显示消息
    messagebox.showinfo("截图完成", f"所有截图已保存至指定文件夹：\n{path}")


# 定义开始截图的函数
def start_screenshot_timer():
    global stop_event
    try:
        interval = int(interval_entry.get())
        total = int(total_entry.get())
        path = path_entry.get()
        if not os.path.exists(path):
            os.makedirs(path)

        stop_event.clear()  # 重置停止信号
        thread = threading.Thread(target=screenshot_thread, args=(interval, total, path))
        thread.start()  # 开始截图线程
    except ValueError:
        messagebox.showerror("错误", "请输入有效的数字。")


# 定义停止截图的函数
def stop_screenshot_timer():
    global stop_event
    stop_event.set()  # 设置停止信号
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
path_label.grid(row=0, column=0, sticky="e")
path_entry.grid(row=0, column=1)
browse_button.grid(row=0, column=2)
interval_label.grid(row=1, column=0, sticky="e")
interval_entry.grid(row=1, column=1)
total_label.grid(row=2, column=0, sticky="e")
total_entry.grid(row=2, column=1)
start_button.grid(row=3, column=0, columnspan=2)
stop_button.grid(row=3, column=2)
screenshot_count_label.grid(row=4, column=1, columnspan=3, pady=1)
# 运行主事件循环
root.mainloop()
