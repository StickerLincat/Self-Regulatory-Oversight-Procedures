# -*- coding: utf-8 -*-
import os
import sys
import json
import time
import random
import winreg
import ctypes
import threading
import psutil
import subprocess
import hashlib
from datetime import datetime, time as dt_time, timedelta
from tkinter import *
from tkinter import ttk, messagebox, simpledialog
import pystray
from PIL import Image
import winsound

# 配置路径
CONFIG_DIR = os.path.join(os.getenv('APPDATA'), 'SupervisorApp')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'supervisor_config.json')
ADMIN_CHECK = hasattr(ctypes, 'windll')
INSTANCE_LOCK = hashlib.md5(os.path.abspath(__file__).encode()).hexdigest()[:8] + ".lock"
ALERT_DURATION = 10  # 弹窗自动关闭时间（秒）
GRACE_PERIOD = 30    # 提前检测时间（分钟）

class SupervisorApp:
    def __init__(self, is_guardian=False):
        # 创建配置目录
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)

        # 单实例检测
        if self.is_already_running() and not is_guardian:
            messagebox.showwarning("警告", "程序已经在运行中")
            os._exit(1)

        self.root = Tk()
        self.root.withdraw()
        self.tray_icon = None
        self.supervision_items = []
        self.process_blacklist = []
        self.tomato_duration = 1500  # 默认25分钟
        self.is_working = False
        self.is_guardian = is_guardian

        self.load_config()
        self.create_tray_icon()

        if not self.is_guardian:
            self.launch_guardian()

        threading.Thread(target=self.time_monitor, daemon=True).start()
        threading.Thread(target=self.process_monitor, daemon=True).start()

    def is_already_running(self):
        try:
            if os.path.exists(INSTANCE_LOCK):
                with open(INSTANCE_LOCK, 'r') as f:
                    pid = int(f.read().strip())
                    if psutil.pid_exists(pid):
                        return True
            with open(INSTANCE_LOCK, 'w') as f:
                f.write(str(os.getpid()))
            return False
        except:
            return False

    def create_tray_icon(self):
        menu_items = [
            pystray.MenuItem('打开控制面板', self.show_control_panel),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('番茄工作法', self.show_tomato_panel),
            pystray.MenuItem('开机自启动', self.toggle_autorun, checked=lambda x: self.is_autorun_enabled()),
            pystray.MenuItem('退出', self.quit_app)
        ]
        image = Image.new('RGB', (64, 64), 'white')
        self.tray_icon = pystray.Icon("supervisor", image, "自律监督程序", pystray.Menu(*menu_items))
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def load_config(self):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                self.supervision_items = data.get('items', [])
                self.process_blacklist = [p.lower() for p in data.get('blacklist', [])]
                self.tomato_duration = data.get('tomato_duration', 1500)
        except FileNotFoundError:
            self.save_config()

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump({
                'items': self.supervision_items,
                'blacklist': self.process_blacklist,
                'tomato_duration': self.tomato_duration
            }, f)

    # 开机自启动功能
    def is_autorun_enabled(self):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_READ
            )
            value, _ = winreg.QueryValueEx(key, "SupervisorApp")
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            return False
        except WindowsError:
            return False

    def toggle_autorun(self):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            if self.is_autorun_enabled():
                winreg.DeleteValue(key, "SupervisorApp")
            else:
                exe_path = os.path.abspath(sys.executable)
                script_path = os.path.abspath(__file__)
                winreg.SetValueEx(
                    key, "SupervisorApp", 0, winreg.REG_SZ,
                    f'"{exe_path}" "{script_path}"'
                )
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror("错误", f"注册表操作失败: {str(e)}")

    # 时间监控功能
    def time_monitor(self):
        while True:
            now = datetime.now().time()
            for item in self.supervision_items:
                if not item['active']:
                    continue
                start_time = datetime.strptime(item['start'], "%H:%M").time()
                end_time = datetime.strptime(item['end'], "%H:%M").time()
                if start_time <= now <= end_time:
                    self.execute_supervision(item)
            time.sleep(30)

    def execute_supervision(self, item):
        if item['action'] == '关机':
            subprocess.run(["shutdown", "/s", "/t", "60"])
        elif item['action'] == '锁定':
            ctypes.windll.user32.LockWorkStation()
        self.show_force_alert(item)

    def show_force_alert(self, item):
        alert = Toplevel()
        alert.attributes("-topmost", True)
        alert.protocol("WM_DELETE_WINDOW", lambda: None)
        alert.geometry("400x200+500+300")
        alert.after(ALERT_DURATION * 1000, alert.destroy)
        Frame(alert, bg="#f0f0f0").pack(fill=BOTH, expand=True)
        Label(alert, text="\n自律监督提醒\n", font=("微软雅黑", 14), bg="#f0f0f0").pack()
        Label(alert, 
             text=f"当前处于监督时段：{item['name']}\n{item['start']} - {item['end']}",
             bg="#f0f0f0").pack()
        Label(alert, 
             text=random.choice([
                 "坚持就是胜利！", "未来属于自律的人！",
                 "成功需要持之以恒！", "今日的付出是明日的收获！"
             ]), bg="#f0f0f0").pack(pady=10)
        Button(alert, text="我知道了", command=alert.destroy).pack(pady=5)

    # 进程监控功能
    def process_monitor(self):
        while True:
            for proc in psutil.process_iter(['name']):
                proc_name = proc.info['name'].lower()
                if proc_name in [p.lower() for p in self.process_blacklist]:
                    try:
                        proc.kill()
                        self.show_alert("已阻止分心程序", f"已终止进程: {proc_name}")
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
            time.sleep(5)

    def show_alert(self, title, message):
        alert = Toplevel()
        alert.title(title)
        Label(alert, text=message).pack(padx=20, pady=10)
        Button(alert, text="确定", command=alert.destroy).pack(pady=5)

    # 控制面板功能
    def show_control_panel(self):
        panel = Toplevel()
        panel.title("控制面板")
        panel.geometry("600x400")

        # 监督时段管理
        ttk.Label(panel, text="监督时段管理", font=("微软雅黑", 12)).pack(pady=5)
        self.supervision_frame = ttk.Frame(panel)
        self.supervision_frame.pack(fill=BOTH, expand=True, padx=10)
        ttk.Button(panel, text="添加监督时段", command=self.add_supervision_item).pack(pady=5)
        self.refresh_supervision_list()

        # 进程黑名单
        ttk.Label(panel, text="进程黑名单", font=("微软雅黑", 12)).pack(pady=5)
        self.process_frame = ttk.Frame(panel)
        self.process_frame.pack(fill=BOTH, expand=True, padx=10)
        ttk.Button(panel, text="添加黑名单进程", command=self.add_blacklist_process).pack(pady=5)
        self.refresh_blacklist()

    def add_supervision_item(self):
        add_win = Toplevel()
        add_win.title("添加监督时段")
        
        # 输入控件
        ttk.Label(add_win, text="事项名称:").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(add_win)
        name_entry.grid(row=0, column=1)

        ttk.Label(add_win, text="开始时间 (HH:MM):").grid(row=1, column=0)
        start_entry = ttk.Entry(add_win)
        start_entry.grid(row=1, column=1)

        ttk.Label(add_win, text="结束时间 (HH:MM):").grid(row=2, column=0)
        end_entry = ttk.Entry(add_win)
        end_entry.grid(row=2, column=1)

        ttk.Label(add_win, text="执行操作:").grid(row=3, column=0)
        action_var = StringVar()
        ttk.Combobox(add_win, textvariable=action_var, 
                     values=["关机", "锁定", "提醒"]).grid(row=3, column=1)

        def save():
            if not all([name_entry.get(), start_entry.get(), end_entry.get(), action_var.get()]):
                messagebox.showerror("错误", "请填写所有字段")
                return

            try:
                start = datetime.strptime(start_entry.get(), "%H:%M").time()
                end = datetime.strptime(end_entry.get(), "%H:%M").time()
                if start >= end:
                    messagebox.showerror("错误", "结束时间必须晚于开始时间")
                    return
            except ValueError:
                messagebox.showerror("错误", "时间格式应为HH:MM")
                return

            new_item = {
                "name": name_entry.get(),
                "start": start_entry.get(),
                "end": end_entry.get(),
                "action": action_var.get(),
                "active": True
            }
            self.supervision_items.append(new_item)
            self.save_config()
            add_win.destroy()
            self.refresh_supervision_list()

        ttk.Button(add_win, text="保存", command=save).grid(row=4, columnspan=2, pady=10)

    def refresh_supervision_list(self):
        for widget in self.supervision_frame.winfo_children():
            widget.destroy()
        
        for idx, item in enumerate(self.supervision_items):
            frame = ttk.Frame(self.supervision_frame)
            frame.pack(fill=X, pady=2)

            chk_var = BooleanVar(value=item['active'])
            ttk.Checkbutton(
                frame,
                variable=chk_var,
                command=lambda i=idx, var=chk_var: self.toggle_item(i, var)
            ).pack(side=LEFT)
            ttk.Label(frame, text=f"{item['name']} {item['start']}-{item['end']}").pack(side=LEFT)
            ttk.Button(frame, text="编辑", command=lambda i=idx: self.edit_item(i)).pack(side=LEFT, padx=5)
            ttk.Button(frame, text="删除", command=lambda i=idx: self.delete_item(i)).pack(side=RIGHT)

    def toggle_item(self, index, check_var):
        if self.check_restricted_operation("修改状态"):
            check_var.set(not check_var.get())
            return
        self.supervision_items[index]['active'] = check_var.get()
        self.save_config()

    def edit_item(self, index):
        item = self.supervision_items[index]
        if self.check_restricted_operation("编辑"):
            return

        edit_win = Toplevel()
        edit_win.title("编辑监督时段")
        
        # 预填充数据
        ttk.Label(edit_win, text="事项名称:").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(edit_win)
        name_entry.insert(0, item['name'])
        name_entry.grid(row=0, column=1)

        ttk.Label(edit_win, text="开始时间 (HH:MM):").grid(row=1, column=0)
        start_entry = ttk.Entry(edit_win)
        start_entry.insert(0, item['start'])
        start_entry.grid(row=1, column=1)

        ttk.Label(edit_win, text="结束时间 (HH:MM):").grid(row=2, column=0)
        end_entry = ttk.Entry(edit_win)
        end_entry.insert(0, item['end'])
        end_entry.grid(row=2, column=1)

        action_var = StringVar(value=item['action'])
        ttk.Combobox(edit_win, textvariable=action_var, values=["关机", "锁定", "提醒"]).grid(row=3, column=1)

        def save():
            # 验证和保存逻辑
            if not all([name_entry.get(), start_entry.get(), end_entry.get(), action_var.get()]):
                messagebox.showerror("错误", "请填写所有字段")
                return

            try:
                start = datetime.strptime(start_entry.get(), "%H:%M").time()
                end = datetime.strptime(end_entry.get(), "%H:%M").time()
                if start >= end:
                    messagebox.showerror("错误", "结束时间必须晚于开始时间")
                    return
            except ValueError:
                messagebox.showerror("错误", "时间格式应为HH:MM")
                return

            self.supervision_items[index] = {
                "name": name_entry.get(),
                "start": start_entry.get(),
                "end": end_entry.get(),
                "action": action_var.get(),
                "active": item['active']
            }
            self.save_config()
            edit_win.destroy()
            self.refresh_supervision_list()

        ttk.Button(edit_win, text="保存", command=save).grid(row=4, columnspan=2, pady=10)

    def delete_item(self, index):
        if self.check_restricted_operation("删除"):
            return
        if messagebox.askyesno("确认删除", "确定要删除这个监督时段吗？"):
            del self.supervision_items[index]
            self.save_config()
            self.refresh_supervision_list()

    def add_blacklist_process(self):
        proc = simpledialog.askstring("添加进程", "输入要阻止的进程名称（如chrome.exe）：")
        if proc and proc not in self.process_blacklist:
            self.process_blacklist.append(proc)
            self.save_config()
            self.refresh_blacklist()

    def refresh_blacklist(self):
        for widget in self.process_frame.winfo_children():
            widget.destroy()
        
        for idx, proc in enumerate(self.process_blacklist):
            frame = ttk.Frame(self.process_frame)
            frame.pack(fill=X, pady=2)
            ttk.Label(frame, text=proc).pack(side=LEFT)
            ttk.Button(frame, text="删除", 
                      command=lambda i=idx: self.delete_process(i)).pack(side=RIGHT)

    def delete_process(self, index):
        del self.process_blacklist[index]
        self.save_config()
        self.refresh_blacklist()

    def is_in_restricted_period(self, check_time=None):
        now = datetime.now().time() if check_time is None else check_time
        for item in self.supervision_items:
            if not item['active']:
                continue
            
            start = datetime.strptime(item['start'], "%H:%M").time()
            end = datetime.strptime(item['end'], "%H:%M").time()
            early_start = (datetime.combine(datetime.today(), start) - 
                          timedelta(minutes=GRACE_PERIOD)).time()
            
            if (early_start <= now <= end) or (start <= now <= end):
                return True
        return False

    def check_restricted_operation(self, operation_type):
        if self.is_in_restricted_period():
            msg = f"当前处于或即将进入监督时段，无法{operation_type}！"
            messagebox.showwarning("操作受限", msg)
            return True
        return False

    # 番茄工作法功能
    def show_tomato_panel(self):
        tomato_win = Toplevel()
        tomato_win.title("番茄钟")
        
        self.tomato_remaining = self.tomato_duration
        self.is_working = False
        
        self.time_label = ttk.Label(tomato_win, text=self.format_time(), font=("Arial", 24))
        self.time_label.pack(pady=20)
        
        btn_frame = ttk.Frame(tomato_win)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="开始专注", command=self.start_tomato).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="重置", command=self.reset_tomato).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="设置", command=self.set_tomato_time).grid(row=0, column=2, padx=5)

    def format_time(self):
        mins, secs = divmod(self.tomato_remaining, 60)
        return f"{mins:02d}:{secs:02d}"

    def set_tomato_time(self):
        mins = simpledialog.askinteger("设置时间", "请输入专注时长（分钟）：", initialvalue=self.tomato_duration//60)
        if mins and 1 <= mins <= 120:
            self.tomato_duration = mins * 60
            self.save_config()
            self.reset_tomato()

    def start_tomato(self):
        if not self.is_working:
            self.is_working = True
            self.tomato_countdown()

    def reset_tomato(self):
        self.tomato_remaining = self.tomato_duration
        self.is_working = False
        self.time_label.config(text=self.format_time())

    def tomato_countdown(self):
        def update():
            if self.tomato_remaining <= 0 or not self.is_working:
                if self.tomato_remaining <= 0:
                    winsound.Beep(1000, 800)
                    messagebox.showinfo("完成", "专注时段结束！")
                    self.reset_tomato()
                return
            self.time_label.config(text=self.format_time())
            self.tomato_remaining -= 1
            self.root.after(1000, update)
        
        update()

    def quit_app(self):
        if self.check_restricted_operation("退出"):
            return
        try:
            os.remove(INSTANCE_LOCK)
        except:
            pass
        self.tray_icon.stop()
        self.root.destroy()
        os._exit(0)

    # 守护进程功能
    def launch_guardian(self):
        if not self.is_guardian and not self.is_process_running("--guardian"):
            try:
                subprocess.Popen(
                    [sys.executable, __file__, "--guardian"],
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            except Exception as e:
                print(f"启动守护进程失败: {str(e)}")

    def is_process_running(self, name):
        current_pid = os.getpid()
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if proc.info['pid'] == current_pid:
                    continue
                cmdline = proc.info.get('cmdline', []) or []
                if name in cmdline:
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False

if __name__ == "__main__":
    is_guardian = "--guardian" in sys.argv
    if not ADMIN_CHECK and not is_guardian:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    else:
        app = SupervisorApp(is_guardian=is_guardian)
        app.root.mainloop()
