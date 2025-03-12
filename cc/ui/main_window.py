import os
import sys
import threading
import logging
import webbrowser
import ctypes
import pywinstyles

import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from cc.utils.cleanup import clean_old_backups
from cc.utils.path import resource_path
from cc.ps_manager import save_psd_as, get_ps_info, thread_save_psd_as
from cc.ui.tray_icon import start_tray_icon
from cc.config import ConfigManager

class MainWindow:
    def __init__(self):
        # fix dpi
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        # 获取屏幕的缩放比例
        self.scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100.0
        
        # 初始窗口可见状态
        self.visible = True
        self.title = "Creative Cache"
        self.icon = resource_path("resources\icon.ico")
        self.tray = None
        start_tray_icon(self)

        self.root = tk.Tk()
        # self.root.overrideredirect(True)
        self.apply_theme_to_titlebar()

        self.cleanup_thread = None
        self.cleanup_job = None
        self.cleanup_info = None
        self.auto_save_job = None
        self.auto_save_thread = None

        # PS 是否为第一次启动
        self.first_psd = True
        # PS 信息获取尝试次数
        self.ps_check_count = 0
        # 创建变量
        self.folder_path_var = tk.StringVar()
        self.auto_save_interval_var = tk.IntVar()
        self.auto_save_var = tk.BooleanVar()
        self.backup_clean_var = tk.BooleanVar()
        self.backup_clean_interval_var = tk.IntVar()
        
        self.config_manager = ConfigManager()
        self.initialize_variables()
        self.setup_ui()
        self.update_ps_info_periodically()
        self.handle_auto_save()

        self.root.mainloop()

    def setup_ui(self):
        self.root.title(" " + self.title)
        self.root.iconbitmap(self.icon)
        
        self.root.geometry(f"{int(600*self.scale)}x{int(320*self.scale)}")

        # 窗口可以拖动
        self.root.bind("<ButtonPress-1>", self.start_move)
        self.root.bind("<ButtonRelease-1>", self.stop_move)
        self.root.bind("<B1-Motion>", self.on_move)

        self.style = ttk.Style("doughnut")

        self.create_widgets()

    def create_widgets(self):
        # 创建并布局所有的界面元素

        # 使用网格布局
        padx = pady = int(10*self.scale)

        # Frame 0: Photoshop info
        frame0 = ttk.Frame(self.root)
        frame0.grid(row=0, column=0, sticky='ew', padx=padx, pady=pady)

        # 创建两个Label，分别用于显示版本和文档信息
        self.ps_version_label = ttk.Label(frame0, text="", foreground='black')
        self.ps_version_label.grid(row=0, column=0, padx=padx, sticky='w')

        self.active_doc_label = ttk.Label(frame0, text="", foreground='black')
        self.active_doc_label.grid(row=1, column=0, padx=padx, sticky='w')

        self.clean_info_label = ttk.Label(frame0, text="")
        self.clean_info_label.grid(row=2, column=0, padx=padx, sticky='w')

        # Frame 1 和 Frame 2 之间的填充 Frame
        filler_frame = ttk.Frame(self.root)
        filler_frame.grid(row=1, column=0, sticky='ns', padx=padx, pady=pady)

        # 配置 filler_frame 的行配置，使其在垂直方向上扩展
        filler_frame.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Frame 2: Auto-save settings and backup retention
        combined_frame = ttk.Frame(self.root)
        combined_frame.grid(row=2, column=0, sticky='ew', padx=padx, pady=pady)
        combined_frame.columnconfigure(1, weight=1)  # 备份间隔的 Spinbox 可伸缩
        combined_frame.columnconfigure(5, weight=1)  # 保留时间
        combined_frame.columnconfigure(0, weight=0)
        combined_frame.columnconfigure(2, weight=0)
        combined_frame.columnconfigure(3, weight=0)
        combined_frame.columnconfigure(4, weight=0)
        combined_frame.columnconfigure(6, weight=0)

        # 自动备份设置
        auto_save_label = ttk.Label(combined_frame, text="备份间隔 \\ 分钟: ")
        auto_save_label.grid(row=0, column=0, padx=padx)
        self.auto_save_entry = ttk.Spinbox(combined_frame, from_=1, to=999, textvariable=self.auto_save_interval_var, width=padx)
        self.auto_save_entry.grid(row=0, column=1, sticky='ew')
        self.auto_save_button = ttk.Checkbutton(combined_frame, text="自动备份", bootstyle="round-toggle")
        self.auto_save_button.grid(row=0, column=2, padx=padx)
        self.auto_save_button.bind("<Button-1>", self.confirm_backup)
        self.auto_save_button.state(['selected' if self.auto_save_var.get() else '!selected'])

        # 分隔符
        separator = ttk.Separator(combined_frame, orient='vertical')
        separator.grid(row=0, column=3, sticky='ns')

        # 备份保留设置
        backup_clean_interval_label = ttk.Label(combined_frame, text="保留时间 \\ 天: ")
        backup_clean_interval_label.grid(row=0, column=4, padx=padx)
        self.backup_clean_interval_entry = ttk.Spinbox(combined_frame, from_=1, to=999, textvariable=self.backup_clean_interval_var, width=padx)
        self.backup_clean_interval_entry.grid(row=0, column=5, sticky='ew')

        self.backup_clean_button = ttk.Checkbutton(combined_frame, text="自动清理", bootstyle="round-toggle")
        self.backup_clean_button.grid(row=0, column=6, padx=padx)
        self.backup_clean_button.bind("<Button-1>", self.confirm_cleanup)
        self.backup_clean_button.state(['selected' if self.backup_clean_var.get() else '!selected'])

        # Frame 1: Folder path selection
        frame1 = ttk.Frame(self.root)
        frame1.grid(row=3, column=0, sticky='ew', padx=padx, pady=pady)
        frame1.columnconfigure(0, weight=1)  # 第一列可伸缩
        frame1.columnconfigure(1, weight=0)  # 不伸缩
        frame1.columnconfigure(2, weight=0)  # 不伸缩

        self.folder_path_entry = ttk.Entry(frame1, textvariable=self.folder_path_var)
        self.folder_path_entry.grid(row=0, column=0, sticky='ew', padx=padx)

        browse_button = ttk.Button(frame1, text="选择目录", command=self.browse_folder)
        browse_button.grid(row=0, column=1, padx=padx)

        open_button = ttk.Button(frame1, text="打开目录", command=self.open_backup_folder)
        open_button.grid(row=0, column=2, padx=padx)

        # 配置 root 的列 0 来使 frame1 可以扩展填满整个窗口宽度
        self.root.columnconfigure(0, weight=1)

        # Frame 3: Control buttons
        frame3 = ttk.Frame(self.root)
        frame3.grid(row=4, column=0, sticky='ew', padx=padx, pady=pady)

        # Set the weight of the second column to a large number
        frame3.columnconfigure(1, weight=1)

        link = ttk.Label(frame3, text="v0.2.2", cursor="hand2", foreground='#dc888e')
        link.grid(row=0, column=0, padx=padx, sticky='ew')
        link.bind("<Button-1>", self.open_link)

        save_button = ttk.Button(frame3, text="立即备份", command=self.start_save, bootstyle="secondary")
        save_button.grid(row=0, column=2, padx=padx)
        self.clean_button = ttk.Button(frame3, text="清理备份", command=self.start_cleanup, bootstyle="warning")
        self.clean_button.grid(row=0, column=3, padx=padx)

        exit_button = ttk.Button(frame3, text="退出程序", command=self.exit_button_clicked, bootstyle="info")
        exit_button.grid(row=0, column=4, padx=padx)

        # Frame 5: Blank
        frame5 = ttk.Frame(self.root)
        frame5.grid(row=5, column=0, sticky='ew', padx=padx, pady=pady/2)

        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

    def initialize_variables(self):
        """Initializes the Tkinter variables with values from the configuration."""
        # Load values from the configuration manager
        folder_path = self.config_manager.get_folder_path()
        auto_save, auto_save_interval = self.config_manager.get_auto_save_settings()
        backup_clean, backup_clean_interval = self.config_manager.get_backup_clean_settings()

        # Assign values to Tkinter variables
        self.folder_path_var = tk.StringVar(value=folder_path)
        self.auto_save_var = tk.BooleanVar(value=auto_save)
        self.auto_save_interval_var = tk.IntVar(value=auto_save_interval)
        self.backup_clean_var = tk.BooleanVar(value=backup_clean)
        self.backup_clean_interval_var = tk.IntVar(value=backup_clean_interval)

        # Adding traces
        self.folder_path_var.trace_add("write", self.update_folder_path)
        self.auto_save_interval_var.trace_add("write", self.validate_auto_save_interval)
        self.auto_save_var.trace_add("write", self.update_auto_save)
        self.backup_clean_var.trace_add("write", self.update_backup_clean)
        self.backup_clean_interval_var.trace_add("write", self.validate_backup_clean_interval)

    def start_move(self, event):
        self.root.x = event.x
        self.root.y = event.y

    def stop_move(self, event):
        self.root.x = None
        self.root.y = None

    def on_move(self, event):
        if event.widget != self.root:
            return
        deltax = event.x - self.root.x
        deltay = event.y - self.root.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def apply_theme_to_titlebar(self):
        version = sys.getwindowsversion()

        if version.major == 10 and version.build >= 22000:
            # Set the title bar color to the background color on Windows 11 for better appearance
            pywinstyles.change_header_color(self.root, "#1c1c1c")
        elif version.major == 10:
            pywinstyles.apply_style(self.root, "dark")

            # A hacky way to update the title bar's color on Windows 10 (it doesn't update instantly like on Windows 11)
            self.root.wm_attributes("-alpha", 0.99)
            self.root.wm_attributes("-alpha", 1)

    def update_folder_path(self, *args):
        self.config_manager.set_folder_path(self.folder_path_var.get())

    def validate_auto_save_interval(self, *args):
        try:
            interval = self.auto_save_interval_var.get()
        except tk.TclError:
            if self.auto_save_entry.get().strip() != "":
                messagebox.showerror("Invalid Input", "Auto save interval must be a positive integer.")
                self.auto_save_interval_var.set(self.config_manager.get_auto_save_settings()[1])
            return
        
        if interval > 0 and interval < 99999:
            self.update_auto_save_interval()
        else:
            messagebox.showerror("Invalid Input", "Auto save interval must be between 1 and 99999.")
            # Reset to last valid value
            self.auto_save_interval_var.set(self.config_manager.get_auto_save_settings()[1])

    def update_auto_save_interval(self):
        self.config_manager.set_auto_save_settings(self.auto_save_var.get(), self.auto_save_interval_var.get())

    def update_auto_save(self, *args):
        self.config_manager.set_auto_save_settings(self.auto_save_var.get(), self.auto_save_interval_var.get())
        self.handle_auto_save()

    def validate_backup_clean_interval(self, *args):
        try:
            interval = self.backup_clean_interval_var.get()
        except tk.TclError:
            if self.backup_clean_interval_entry.get().strip() != "":
                messagebox.showerror("Invalid Input", "Backup clean interval must be a positive integer.")
                self.backup_clean_interval_var.set(self.config_manager.get_backup_clean_settings()[1])
            return
        
        if interval > 0 and interval < 99999:
            self.update_backup_clean_interval()
        else:
            messagebox.showerror("Invalid Input", "Backup clean interval must be between 1 and 99999.")
            self.backup_clean_interval_var.set(self.config_manager.get_backup_clean_settings()[1])

    def update_backup_clean(self, *args):
        self.config_manager.set_backup_clean_settings(self.backup_clean_var.get(), self.backup_clean_interval_var.get())
        self.handle_cleanup()

    def update_backup_clean_interval(self):
        self.config_manager.set_backup_clean_settings(self.backup_clean_var.get(), self.backup_clean_interval_var.get())

    def update_ps_info(self):
        ps_version, active_doc = get_ps_info()
        ps_unavailable = ps_version == "Unavailable" and active_doc == "No documents open"

        if self.first_psd:
            if not ps_unavailable:
                # 第一次检测到有文档时触发保存
                self.first_psd = False
                self.start_save(True)
        else:
            if ps_unavailable:
                # 连续检测3次，防止窗口拖动期间 COM 对象无法访问
                self.ps_check_count += 1
                if self.ps_check_count >= 3:
                    self.ps_check_count = 0
                    self.first_psd = True
            else:
                self.ps_check_count = 0
        
        if self.ps_version_label.winfo_exists() and self.active_doc_label.winfo_exists():
            self.ps_version_label.config(text=f"PS Version: {ps_version}")
            self.active_doc_label.config(text=f"Active Doc: {active_doc}")

    def update_ps_info_periodically(self):
        self.update_ps_info()  # 更新 Photoshop 信息
        self.root.after(2000, self.update_ps_info_periodically)  # 每隔2秒调用一次

    def finish_cleanup(self, cleaned_files_count):
        self.clean_info_label.config(text=f"清理完成，共删除 {cleaned_files_count} 个文件。")
        # 清理完成，启用按钮并更新提示信息
        self.clean_button.config(state=tk.NORMAL)
        if self.cleanup_info is not None:
            self.root.after_cancel(self.cleanup_info)
        self.cleanup_info = self.root.after(8000, lambda: self.clean_info_label.config(text=""))


    def confirm_cleanup(self, event):
        backup_path = self.folder_path_var.get()
        if not os.path.isdir(backup_path):
            messagebox.showerror("错误", "无效的备份目录")
            return "break"
        current_value = self.backup_clean_var.get()
        if current_value:
            self.backup_clean_var.set(False)
            self.backup_clean_button.state(['!selected'])
            return "break"

        # 弹出对话框询问用户
        response = messagebox.askyesno("确认", f"您确定要开启自动清理吗？本操作会清理{backup_path}目录中所有以_psbackup结尾，并且时间戳大于{str(self.backup_clean_interval_var.get())}天的psd文件！为保险起见，请尽量不要在备份目录存放其他文件！！！")
        if response:
            # 切换变量的值
            self.backup_clean_var.set(True)
            # 更新Checkbutton的显示状态
            self.backup_clean_button.state(['selected'])

            return "break"


    def start_cleanup(self, silent=False):
        backup_path = self.folder_path_var.get()
        if not os.path.isdir(backup_path):
            logging.error("Invalid backup directory.")
            if not silent:
                messagebox.showerror("错误", "无效的备份目录")
            return
        if not silent:
            if not messagebox.askyesno("确认", f"您确定要清理吗？本操作会清理{backup_path}目录中所有以_psbackup结尾，并且时间戳大于{str(self.backup_clean_interval_var.get())}天的psd文件！为保险起见，请尽量不要在备份目录存放其他文件！！！"):
                return

        if self.cleanup_thread is not None and self.cleanup_thread.is_alive():
            logging.info("Cleanup already in progress.")
            return

        self.clean_button.config(state=tk.DISABLED)
        self.clean_info_label.config(text="清理中...")

        self.cleanup_thread = threading.Thread(target=lambda: clean_old_backups(backup_path, self.backup_clean_interval_var.get(), self.finish_cleanup))
        self.cleanup_thread.daemon = True
        self.cleanup_thread.start()

    def handle_cleanup(self):
        if self.backup_clean_var.get():
            try:
                # TODO 自动清理间隔暂时设置为比自动保存稍长一点
                interval = int(self.auto_save_interval_var.get())
                interval_ms = interval * 70000  # 将分钟转换为毫秒
                if self.cleanup_job is not None:
                    self.root.after_cancel(self.cleanup_job)
                self.cleanup_schedule(interval_ms)
                logging.info(f"Auto cleanup scheduled every {interval} minutes, deleting backups older than {self.backup_clean_interval_var.get()} days.")
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字")
        else:
            if self.cleanup_job is not None:
                self.root.after_cancel(self.cleanup_job)
                self.cleanup_job = None
            logging.info("Auto cleanup disabled.")

    def cleanup_schedule(self, interval_ms):
        self.start_cleanup(silent=True)
        # 确保即使在异常发生时也能重新调度
        try:
            self.cleanup_job = self.root.after(interval_ms, self.cleanup_schedule, interval_ms)
        except Exception as e:
            logging.error(f"Error in scheduling auto save: {str(e)}")
        # logging.info("Auto-save executed and rescheduled.")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.config_manager.set_folder_path(folder_selected)
            self.folder_path_var.set(folder_selected)

    def open_link(self, event):
        webbrowser.open_new(r"https://github.com/aimkiray/cc")

    def confirm_backup(self, event):
        backup_path = self.folder_path_var.get()
        if not os.path.isdir(backup_path):
            messagebox.showerror("错误", "无效的备份目录")
            return "break"
        current_value = self.auto_save_var.get()
        if current_value:
            self.auto_save_var.set(False)
            self.auto_save_button.state(['!selected'])
            return "break"

        self.auto_save_var.set(True)
        self.auto_save_button.state(['selected'])
        return "break"


    def start_save(self, silent=False):
        if self.auto_save_thread and self.auto_save_thread.is_alive():
            logging.info("An automatic backup operation is already in progress, please try again later or adjust the automatic backup interval.")
            if not silent:
                messagebox.showinfo("信息", "已有一个备份操作正在进行，请稍后再试。")
            return
        
        backup_path = self.folder_path_var.get()
        if not os.path.isdir(backup_path):
            logging.error("Invalid backup directory.")
            if not silent:
                messagebox.showerror("错误", "无效的备份目录")
            return
        
        if silent:
            # 后台自动备份，不需要提示
            self.run_auto_save(backup_path)
        else:
            backup_result = save_psd_as(backup_path)

            if backup_result:
                messagebox.showinfo("成功", f"文件已成功备份到：{backup_result}")
            else:
                messagebox.showerror("错误", f"备份未成功：Photoshop 未准备就绪。请查阅日志文件以了解详细信息。")

    def handle_auto_save(self):
        if self.auto_save_var.get():
            try:
                interval = int(self.auto_save_interval_var.get())
                interval_ms = interval * 60000  # 将分钟转换为毫秒
                if self.auto_save_job is not None:
                    self.root.after_cancel(self.auto_save_job)
                self.auto_save_schedule(interval_ms)
                logging.info(f"Auto save scheduled every {interval} minutes.")
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字")
        else:
            if self.auto_save_job is not None:
                self.root.after_cancel(self.auto_save_job)
                self.auto_save_job = None
            logging.info("Auto save disabled.")

    def auto_save_schedule(self, interval_ms):
        self.start_save(silent=True)
        # 确保即使在异常发生时也能重新调度
        try:
            self.auto_save_job = self.root.after(interval_ms, self.auto_save_schedule, interval_ms)
        except Exception as e:
            logging.error(f"Error in scheduling auto save: {str(e)}")
        # logging.info("Auto-save executed and rescheduled.")

    def run_auto_save(self, folder_path):
        self.auto_save_thread = threading.Thread(target=thread_save_psd_as, args=(folder_path,))
        self.auto_save_thread.daemon = True
        self.auto_save_thread.start()

    def hide_window(self):
        self.root.overrideredirect(True)
        self.root.withdraw()
        self.visible = False
        if self.tray:
            self.tray.update_menu()

    def show_window(self):
        self.root.deiconify()
        self.visible = True
        if self.tray:
            self.tray.update_menu()
        self.root.overrideredirect(False)

    def open_backup_folder(self):
      backup_folder = self.folder_path_var.get()
      if os.path.isdir(backup_folder):
          os.startfile(backup_folder)
      else:
          messagebox.showinfo("错误", "无效的备份目录")

    def exit_button_clicked(self):
      if messagebox.askyesno("确认", "您确定要退出程序吗？"):
          self.exit_app()

    def exit_app(self):
        self.root.destroy()
        # self.root.quit()
