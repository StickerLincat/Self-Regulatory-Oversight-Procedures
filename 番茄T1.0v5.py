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
import pywintypes
from datetime import datetime, time as dt_time, timedelta
from tkinter import *
from tkinter import ttk, messagebox, simpledialog
import pystray
from PIL import Image
import winsound
import win32service
import win32event
import servicemanager
import win32serviceutil

CONFIG_DIR = os.path.join(os.getenv('APPDATA'), 'SupervisorApp')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'supervisor_config.json')
ADMIN_CHECK = hasattr(ctypes, 'windll')
INSTANCE_LOCK = hashlib.md5(os.path.abspath(__file__).encode()).hexdigest()[:8] + ".lock"
ALERT_DURATION = 10
GRACE_PERIOD = 30
WINDOW_POSITIONS = {}

class SupervisorService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'SupervisorService'
    _svc_display_name_ = 'Supervisor Service'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def SvcDoRun(self):
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        self.main()

    def main(self):
        while self.is_alive:
            if not SupervisorApp().is_process_running(__file__):
                subprocess.Popen([sys.executable, __file__], creationflags=subprocess.CREATE_NO_WINDOW)
            time.sleep(60)

class SupervisorApp:
    def __init__(self, is_guardian=False):
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)

        if self.is_already_running() and not is_guardian:
            messagebox.showwarning("警告", "程序已经在运行中")
            os._exit(1)

        self.root = Tk()
        self.root.withdraw()
        self.tray_icon = None
        self.supervision_items = []
        self.global_blacklist = []
        self.tomato_duration = 1500
        self.is_working = False
        self.is_guardian = is_guardian

        self.load_config()
        self.create_tray_icon()
        self.load_window_positions()

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
            pystray.MenuItem('开启服务', self.toggle_service),
            pystray.MenuItem('退出', self.quit_app)
        ]
        image = Image.new('RGB', (64, 64), 'white')
        self.tray_icon = pystray.Icon("supervisor", image, "自律监督程序", pystray.Menu(*menu_items))
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def generate_monitor_script(self):
        script_content = """
import os
import time
import psutil
import subprocess

main_program_path = r'{}'

def is_program_running():
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            cmdline = proc.info.get('cmdline', []) or []
            if main_program_path in cmdline:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return False

def start_program():
    subprocess.Popen([main_program_path], creationflags=subprocess.CREATE_NO_WINDOW)

def main():
    while True:
        if not is_program_running():
            start_program()
        time.sleep(30)

if __name__ == "__main__":
    main()
""".format(os.path.abspath(__file__))

        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_script.py')
        try:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            print(f"监控脚本已生成: {script_path}")
            return script_path
        except Exception as e:
            print(f"生成监控脚本失败: {e}")
            return None

    def install_pyinstaller(self):
        try:
            import pyinstaller
        except ImportError:
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], check=True)

    def pack_to_exe(self, script_path):
        try:
            dist_dir = os.path.join(os.path.dirname(script_path), 'dist')
            exe_path = os.path.join(dist_dir, 'monitor_script.exe')
            
            if os.path.exists(exe_path):
                print(f"可执行文件已存在: {exe_path}")
                return exe_path
            
            try:
                import PyInstaller
            except ImportError:
                print("正在安装PyInstaller...")
                subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], check=True)
            
            print(f"正在打包: {script_path}")
            subprocess.run([
                sys.executable, '-m', 'PyInstaller',
                '--onefile', '--noconsole',
                '--distpath', dist_dir,
                script_path
            ], check=True)
            
            if os.path.exists(exe_path):
                print(f"可执行文件生成成功: {exe_path}")
                return exe_path
            else:
                print(f"可执行文件未生成: {exe_path}")
                return None
        except subprocess.CalledProcessError as e:
            print(f"打包失败: {e.stderr}")
            messagebox.showerror("错误", f"打包失败: {str(e)}")
            return None

    def toggle_service(self):
        try:
            if self.is_service_running():
                win32serviceutil.StopService('SupervisorService')
            else:
                if not self.is_service_installed():
                    self.install_pyinstaller()
                    
                    script_path = self.generate_monitor_script()
                    if not script_path:
                        messagebox.showerror("错误", "监控脚本生成失败")
                        return
                    
                    exe_path = self.pack_to_exe(script_path)
                    if not exe_path or not os.path.exists(exe_path):
                        messagebox.showerror("错误", f"生成的可执行文件不存在: {exe_path}")
                        return
                    
                    print(f"可执行文件路径: {exe_path}")
                    try:
                        result = subprocess.run([
                            'sc', 'create', 'SupervisorService',
                            'binPath=', exe_path,
                            'start=', 'auto',
                            'DisplayName=', 'Supervisor Service'
                        ], capture_output=True, text=True, check=True)
                        print(f"服务创建输出: {result.stdout}")
                    except subprocess.CalledProcessError as e:
                        print(f"服务创建失败: {e.stderr}")
                        raise
                    
                try:
                    win32serviceutil.StartService('SupervisorService')
                    print("服务启动成功")
                except pywintypes.error as e:
                    print(f"服务启动失败: {e}")
                    messagebox.showerror("错误", f"服务启动失败: {e}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("错误", f"服务创建失败: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"服务操作失败: {str(e)}")

    def is_service_installed(self):
        try:
            result = subprocess.run(
                ['sc', 'query', 'SupervisorService'],
                capture_output=True, text=True
            )
            return "SERVICE_NAME: SupervisorService" in result.stdout
        except subprocess.CalledProcessError:
            return False

    def is_service_running(self):
        try:
            status = win32serviceutil.QueryServiceStatus('SupervisorService')[1]
            return status == win32service.SERVICE_RUNNING
        except:
            return False

    def save_window_positions(self):
        for window in self.root.winfo_children():
            if isinstance(window, Toplevel):
                geometry = window.geometry()
                WINDOW_POSITIONS[window.title()] = geometry
        with open(os.path.join(CONFIG_DIR, 'window_positions.json'), 'w') as f:
            json.dump(WINDOW_POSITIONS, f)

    def load_window_positions(self):
        try:
            with open(os.path.join(CONFIG_DIR, 'window_positions.json'), 'r') as f:
                global WINDOW_POSITIONS
                WINDOW_POSITIONS = json.load(f)
                for title, geometry in WINDOW_POSITIONS.items():
                    for window in self.root.winfo_children():
                        if isinstance(window, Toplevel) and window.title() == title:
                            window.geometry(geometry)
                            window.bind("<Configure>", lambda e: self.save_window_positions())
        except:
            pass

    def show_tomato_panel(self):
        tomato_win = Toplevel()
        tomato_win.title("番茄钟")
        tomato_win.bind("<Configure>", lambda e: self.save_window_positions())
        
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
        mins = simpledialog.askinteger("设置时间", "请输入专注时长（分钟）：", initialvalue=self.tomato_duration // 60)
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

    def load_config(self):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                self.supervision_items = data.get('items', [])
                self.global_blacklist = data.get('global_blacklist', [])
                self.tomato_duration = data.get('tomato_duration', 1500)
        except FileNotFoundError:
            self.save_config()

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump({
                'items': self.supervision_items,
                'global_blacklist': self.global_blacklist,
                'tomato_duration': self.tomato_duration
            }, f)

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
        except:
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
        elif item['action'] == '仅启用黑名单（不弹窗）':
            pass
        else:
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

    def process_monitor(self):
        while True:
            if self.is_in_restricted_period():
                for item in self.supervision_items:
                    if not item['active']:
                        continue
                    start_time = datetime.strptime(item['start'], "%H:%M").time()
                    end_time = datetime.strptime(item['end'], "%H:%M").time()
                    now = datetime.now().time()
                    if start_time <= now <= end_time:
                        if item.get('enable_blacklist', False):
                            self.kill_blacklist_processes(item.get('blacklist', []))
                        else:
                            self.kill_blacklist_processes(self.global_blacklist)
            time.sleep(5)

    def kill_blacklist_processes(self, blacklist):
        for proc in psutil.process_iter(['name']):
            proc_name = proc.info['name'].lower()
            for blacklist_item in blacklist:
                if blacklist_item['active'] and proc_name == blacklist_item['name'].lower():
                    try:
                        proc.kill()
                        self.show_alert("已阻止分心程序", f"已终止进程: {proc_name}")
                    except:
                        pass

    def show_alert(self, title, message):
        alert = Toplevel()
        alert.title(title)
        Label(alert, text=message).pack(padx=20, pady=10)
        Button(alert, text="确定", command=alert.destroy).pack(pady=5)

    def show_control_panel(self):
        panel = Toplevel()
        panel.title("控制面板")
        panel.geometry("600x400")
        panel.bind("<Configure>", lambda e: self.save_window_positions())

        ttk.Label(panel, text="监督时段管理", font=("微软雅黑", 12)).pack(pady=5)
        self.supervision_frame = ttk.Frame(panel)
        self.supervision_frame.pack(fill=BOTH, expand=True, padx=10)
        ttk.Button(panel, text="添加监督时段", command=self.add_supervision_item).pack(pady=5)
        self.refresh_supervision_list()

        ttk.Label(panel, text="全局黑名单", font=("微软雅黑", 12)).pack(pady=5)
        self.global_blacklist_frame = ttk.Frame(panel)
        self.global_blacklist_frame.pack(fill=BOTH, expand=True, padx=10)
        ttk.Button(panel, text="添加全局黑名单进程", command=self.add_global_blacklist_process).pack(pady=5)
        self.refresh_global_blacklist()

    def check_new_item_conflict(self, new_item):
        try:
            new_start = datetime.strptime(new_item['start'], "%H:%M").time()
            new_end = datetime.strptime(new_item['end'], "%H:%M").time()
            now = datetime.now().time()
            return new_start <= now <= new_end
        except:
            return False

    def is_item_restricted(self, item):
        if not item['active']:
            return False
            
        try:
            now = datetime.now()
            today = now.date()
            start_time = datetime.strptime(item['start'], "%H:%M").time()
            end_time = datetime.strptime(item['end'], "%H:%M").time()
            early_start = (datetime.combine(today, start_time) - timedelta(minutes=GRACE_PERIOD)).time()
            return (early_start <= now.time() <= end_time)
        except:
            return False

    def add_supervision_item(self):
        add_win = Toplevel()
        add_win.title("添加监督时段")
        
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
        ttk.Combobox(add_win, textvariable=action_var, values=["关机", "锁定", "提醒", "仅启用黑名单（不弹窗）"]).grid(row=3, column=1)

        enable_blacklist_var = BooleanVar(value=False)
        ttk.Checkbutton(add_win, text="启用独立黑名单", variable=enable_blacklist_var).grid(row=4, columnspan=2, pady=5)

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
                "enable_blacklist": enable_blacklist_var.get(),
                "blacklist": [],
                "active": True
            }
            
            if self.check_new_item_conflict(new_item):
                if not messagebox.askyesno("时间冲突", "当前正处于新设定的监督时段内！\n确定要立即启用该规则吗？"):
                    return

            self.supervision_items.append(new_item)
            self.save_config()
            add_win.destroy()
            self.refresh_supervision_list()

        ttk.Button(add_win, text="保存", command=save).grid(row=5, columnspan=2, pady=10)

    def refresh_supervision_list(self):
        for widget in self.supervision_frame.winfo_children():
            widget.destroy()
        
        for idx, item in enumerate(self.supervision_items):
            frame = ttk.Frame(self.supervision_frame)
            frame.pack(fill=X, pady=2)

            chk_var = BooleanVar(value=item['active'])
            is_locked = self.is_item_restricted(item)
            
            chk_btn = ttk.Checkbutton(
                frame,
                variable=chk_var,
                command=lambda i=idx, var=chk_var: self.toggle_item(i, var)
            )
            chk_btn.pack(side=LEFT)
            if is_locked:
                chk_btn.state(['disabled'])
            
            ttk.Label(frame, text=f"{item['name']} {item['start']}-{item['end']}").pack(side=LEFT)
            
            edit_btn = ttk.Button(frame, text="编辑", command=lambda i=idx: self.edit_item(i))
            edit_btn.pack(side=LEFT, padx=5)
            if is_locked:
                edit_btn.state(['disabled'])
            
            del_btn = ttk.Button(frame, text="删除", command=lambda i=idx: self.delete_item(i))
            del_btn.pack(side=RIGHT)
            if is_locked:
                del_btn.state(['disabled'])

    def toggle_item(self, index, check_var):
        item = self.supervision_items[index]
        if self.is_item_restricted(item):
            check_var.set(not check_var.get())
            messagebox.showwarning("操作受限", "该监督时段处于执行期或准备期（前30分钟）\n无法修改状态！")
            return
            
        self.supervision_items[index]['active'] = check_var.get()
        self.save_config()

    def edit_item(self, index):
        item = self.supervision_items[index]
        if self.is_item_restricted(item):
            messagebox.showwarning("操作受限", "该监督时段处于执行期或准备期（前30分钟）\n无法编辑！")
            return

        edit_win = Toplevel()
        edit_win.title("编辑监督时段")
        
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
        ttk.Combobox(edit_win, textvariable=action_var, values=["关机", "锁定", "提醒", "仅启用黑名单（不弹窗）"]).grid(row=3, column=1)

        enable_blacklist_var = BooleanVar(value=item.get('enable_blacklist', False))
        ttk.Checkbutton(edit_win, text="启用独立黑名单", variable=enable_blacklist_var).grid(row=4, columnspan=2, pady=5)

        ttk.Button(edit_win, text="管理独立黑名单", command=lambda: self.manage_blacklist(item, edit_win)).grid(row=5, columnspan=2, pady=5)

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
                "enable_blacklist": enable_blacklist_var.get(),
                "blacklist": item.get('blacklist', []),
                "active": item['active']
            }

            if self.check_new_item_conflict(new_item):
                if not messagebox.askyesno("时间冲突", "当前正处于新设定的监督时段内！\n确定要立即启用该规则吗？"):
                    return

            self.supervision_items[index] = new_item
            self.save_config()
            edit_win.destroy()
            self.refresh_supervision_list()

        ttk.Button(edit_win, text="保存", command=save).grid(row=6, columnspan=2, pady=10)

    def manage_blacklist(self, item, parent_window):
        blacklist_win = Toplevel(parent_window)
        blacklist_win.title("管理独立黑名单")
        blacklist_win.geometry("400x300")

        def refresh_blacklist():
            for widget in blacklist_frame.winfo_children():
                widget.destroy()
            
            for idx, proc in enumerate(item.get('blacklist', [])):
                frame = ttk.Frame(blacklist_frame)
                frame.pack(fill=X, pady=2)

                chk_var = BooleanVar(value=proc['active'])
                chk_btn = ttk.Checkbutton(
                    frame,
                    variable=chk_var,
                    command=lambda i=idx, var=chk_var: self.toggle_blacklist_item(item, i, var)
                )
                chk_btn.pack(side=LEFT)
                
                ttk.Label(frame, text=proc['name']).pack(side=LEFT)
                
                btn = ttk.Button(frame, text="删除", command=lambda i=idx: self.delete_blacklist_item(item, i, refresh_blacklist))
                btn.pack(side=RIGHT)

        ttk.Label(blacklist_win, text="独立黑名单进程", font=("微软雅黑", 12)).pack(pady=5)
        blacklist_frame = ttk.Frame(blacklist_win)
        blacklist_frame.pack(fill=BOTH, expand=True, padx=10)

        refresh_blacklist()

        ttk.Button(blacklist_win, text="添加进程", command=lambda: self.add_blacklist_item(item, refresh_blacklist)).pack(pady=10)

    def add_blacklist_item(self, item, refresh_callback):
        proc = simpledialog.askstring("添加进程", "输入要阻止的进程名称（如chrome.exe）：")
        if proc and proc not in [item['name'] for item in item.get('blacklist', [])]:
            item['blacklist'].append({"name": proc, "active": True})
            self.save_config()
            refresh_callback()

    def toggle_blacklist_item(self, item, index, check_var):
        item['blacklist'][index]['active'] = check_var.get()
        self.save_config()

    def delete_blacklist_item(self, item, index, refresh_callback):
        if messagebox.askyesno("确认删除", "确定要删除这个进程吗？"):
            del item['blacklist'][index]
            self.save_config()
            refresh_callback()

    def delete_item(self, index):
        item = self.supervision_items[index]
        if self.is_item_restricted(item):
            messagebox.showwarning("操作受限", "该监督时段处于执行期或准备期（前30分钟）\n无法删除！")
            return
            
        if messagebox.askyesno("确认删除", "确定要删除这个监督时段吗？"):
            del self.supervision_items[index]
            self.save_config()
            self.refresh_supervision_list()

    def add_global_blacklist_process(self):
        proc = simpledialog.askstring("添加进程", "输入要阻止的进程名称（如chrome.exe）：")
        if proc and proc not in [item['name'] for item in self.global_blacklist]:
            self.global_blacklist.append({"name": proc, "active": True})
            self.save_config()
            self.refresh_global_blacklist()

    def refresh_global_blacklist(self):
        for widget in self.global_blacklist_frame.winfo_children():
            widget.destroy()
        
        for idx, proc in enumerate(self.global_blacklist):
            frame = ttk.Frame(self.global_blacklist_frame)
            frame.pack(fill=X, pady=2)

            chk_var = BooleanVar(value=proc['active'])
            is_locked = self.is_in_restricted_period()
            
            chk_btn = ttk.Checkbutton(
                frame,
                variable=chk_var,
                command=lambda i=idx, var=chk_var: self.toggle_global_blacklist(i, var)
            )
            chk_btn.pack(side=LEFT)
            if is_locked:
                chk_btn.state(['disabled'])
            
            ttk.Label(frame, text=proc['name']).pack(side=LEFT)
            
            btn = ttk.Button(frame, text="删除", command=lambda i=idx: self.delete_global_blacklist_process(i))
            if is_locked:
                btn.state(['disabled'])
            btn.pack(side=RIGHT)

    def toggle_global_blacklist(self, index, check_var):
        if self.is_in_restricted_period():
            check_var.set(not check_var.get())
            messagebox.showwarning("操作受限", "当前处于或即将进入监督时段，无法修改黑名单状态！")
            return
            
        self.global_blacklist[index]['active'] = check_var.get()
        self.save_config()

    def delete_global_blacklist_process(self, index):
        if self.is_in_restricted_period():
            messagebox.showwarning("操作受限", "当前处于或即将进入监督时段，无法删除黑名单进程！")
            return
        if messagebox.askyesno("确认删除", "确定要删除这个进程吗？"):
            del self.global_blacklist[index]
            self.save_config()
            self.refresh_global_blacklist()

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
            except:
                continue
        return False

if __name__ == "__main__":
    is_guardian = "--guardian" in sys.argv
    if not ADMIN_CHECK and not is_guardian:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    else:
        app = SupervisorApp(is_guardian=is_guardian)
        app.root.mainloop()
