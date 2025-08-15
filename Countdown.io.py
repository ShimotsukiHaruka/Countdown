import tkinter as tk
import time
import datetime
import threading
import sys
import winreg
import os
import win32com.client
import ctypes
import keyboard

class CountdownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("倒计时软件")
        
        # 去除窗口边框
        self.root.overrideredirect(True)
        
        # 设置窗口大小（宽8cm，长7cm，转换为像素）
        width_px = 8 * 37.8  # 8cm 转换为像素
        height_px = 7 * 37.8  # 7cm 转换为像素
        self.root.geometry(f"{int(width_px)}x{int(height_px)}")  # 设置窗口大小
        
        # 设置窗口背景为黑色
        self.root.config(bg="black")
        
        # 设置窗口透明度（模拟透明效果）
        self.root.attributes("-alpha", 0.7)  # 0.0 完全透明, 1.0 完全不透明
        
        # 设置窗口始终在最上层
        self.root.attributes("-topmost", True)

        self.activities_frame = tk.Frame(self.root, bg="black")  # 设置活动区域背景为黑色
        self.activities_frame.pack(pady=10)

        self.activities = []  # 存储所有活动的目标时间和名称

        # 实现窗口拖动
        self._dragging = False
        self._drag_data = {"x": 0, "y": 0}

        self.root.bind("<Button-1>", self.on_drag_start)
        self.root.bind("<B1-Motion>", self.on_drag_motion)
        
        # 注册快捷键 Ctrl+Q 关闭程序
        keyboard.add_hotkey('ctrl+q', self.exit_app)

    def on_drag_start(self, event):
        # 记录鼠标点击位置
        self._dragging = True
        self._drag_data["x"] = event.x
        self._drag_data["y"] = event.y

    def on_drag_motion(self, event):
        # 移动窗口
        if self._dragging:
            x = self.root.winfo_x() - self._drag_data["x"] + event.x
            y = self.root.winfo_y() - self._drag_data["y"] + event.y
            self.root.geometry(f"+{x}+{y}")

    def add_activity(self, activity_name, target_time_str):
        if not activity_name or not target_time_str:
            return  # 如果没有输入活动名称或时间，直接返回，不做任何提示
        
        try:
            # 解析用户输入的目标日期时间
            target_time = datetime.datetime.strptime(target_time_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return  # 如果日期时间格式不正确，直接返回，不做任何提示
        
        # 创建倒计时标签
        countdown_label = tk.Label(self.activities_frame, text=f"{activity_name}\n{target_time_str}", font=("Arial", 10), bg="black", fg="white", anchor='w', justify='left')
        countdown_label.pack(fill='x', pady=5)  # fill='x' 保证标签宽度充满窗口，pady=5 给标签之间留出间隔
        
        # 启动倒计时线程
        threading.Thread(target=self.countdown, args=(activity_name, target_time, countdown_label), daemon=True).start()
        
        # 将活动添加到活动列表
        self.activities.append({"name": activity_name, "target_time": target_time, "label": countdown_label})

    def countdown(self, activity_name, target_time, countdown_label):
        while True:
            # 获取当前时间
            current_time = datetime.datetime.now()
            time_remaining = target_time - current_time
            
            if time_remaining.total_seconds() <= 0:
                countdown_label.config(text=f"{activity_name}\n已结束!")
                break
            
            # 计算剩余时间
            days = time_remaining.days
            hours, remainder = divmod(time_remaining.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            
            # 更新显示剩余时间
            countdown_label.config(text=f"{activity_name}\n剩余时间 {days}天 {hours}小时 {minutes}分钟 {seconds}秒")
            time.sleep(1)
    
    def exit_app(self):
        """关闭应用程序"""
        self.root.destroy()
        os._exit(0)

# 隐藏控制台窗口
def hide_console():
    if sys.platform == "win32":
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# 设置开机自启动
def add_to_startup():
    if sys.platform == "win32":
        path = sys.argv[0]  # 获取当前脚本路径
        reg_key = winreg.HKEY_CURRENT_USER
        reg_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        reg_value = "CountdownApp"
        
        try:
            reg = winreg.OpenKey(reg_key, reg_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(reg, reg_value, 0, winreg.REG_SZ, path)
            winreg.CloseKey(reg)
        except Exception as e:
            print(f"无法设置开机自启动: {e}")

# 创建当前目录下的快捷方式
def create_shortcut():
    if sys.platform == "win32":
        current_dir = os.path.dirname(sys.argv[0])  # 获取当前脚本所在的目录
        shortcut_path = os.path.join(current_dir, "CountdownApp.lnk")  # 设置快捷方式路径为当前目录下
        
        target = sys.argv[0]  # 获取当前脚本路径
        wsh = win32com.client.Dispatch("WScript.Shell")
        shortcut = wsh.CreateShortCut(shortcut_path)
        shortcut.TargetPath = target  # 设置快捷方式目标路径为当前脚本
        shortcut.WorkingDirectory = current_dir  # 设置工作目录为当前脚本所在的目录
        shortcut.IconLocation = target  # 设置快捷方式图标为当前脚本
        shortcut.save()  # 保存快捷方式

# 主函数
if __name__ == "__main__":
    # 隐藏控制台窗口
    hide_console()
    
    root = tk.Tk()
    app = CountdownApp(root)
    
    # 设置开机自启动
    add_to_startup()
    
    # 创建当前目录的快捷方式
    create_shortcut()
    
    # 添加活动（你可以在这里直接调用 add_activity 方法添加预设的活动）
    app.add_activity("日文N2报名", "2025-8-26 00:00:00")
    app.add_activity("日文N2检定", "2025-12-1 00:00:00")
    app.add_activity("日文N1检定", "2026-12-1 00:00:00")
    app.add_activity("托福考试", "2026-7-7 00:00:00")
    app.add_activity("EJU", "2027-6-1 00:00:00")
    
    root.mainloop()