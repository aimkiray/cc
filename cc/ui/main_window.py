import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import logging
import webbrowser

from ..utils.cleanup import clean_old_backups
from ..utils.path import resource_path
from ..ps_manager import save_psd_as, get_ps_info, thread_save_psd_as
from .tray_icon import start_tray_icon
from ..config import ConfigManager

class MainWindow:
    def __init__(self, root, scale):
        self.root = root
        self.scale = scale
        self.tray = start_tray_icon(resource_path("icon.ico"), "CreativeCache", self.show_window, self.hide_window, self.exit_app)
        self.cleanup_thread = None
        self.auto_save_job = None
        self.auto_save_thread = None
        # 创建变量
        self.folder_path_var = tk.StringVar()
        self.auto_save_interval_var = tk.IntVar()
        self.auto_save_var = tk.IntVar()
        self.backup_retention_days_var = tk.IntVar()
        
        self.config_manager = ConfigManager()
        self.initialize_variables()
        self.setup_ui()
        self.update_ps_info_periodically()
        self.handle_auto_save()

    def setup_ui(self):
        self.root.title("CreativeCache")
        self.root.iconbitmap(resource_path("icon.ico"))
        self.root.configure(background="#bbded6")
        
        self.root.geometry(f"{int(600*self.scale)}x{int(320*self.scale)}")

        self.style = ttk.Style(self.root)
        self.style.theme_use('default')
        self.style.configure('TButton', background="#fae3d9")
        self.style.configure('TEntry', fieldbackground="#fae3d9")
        self.style.configure('TCheckbutton', background="#bbded6", indicatorcolor="#fae3d9", indicatordiameter=int(10*self.scale))
        self.style.configure('TSpinbox', arrowsize=int(10*self.scale), arrowcolor="#61c0bf", fieldbackground="#fae3d9", background="#bbded6")
        self.style.configure('TFrame', background="#bbded6")
        self.style.configure('TLabel', background="#bbded6")

        self.create_widgets()

    def create_widgets(self):
        # 创建并布局所有的界面元素

        # 使用网格布局
        padx = pady = int(10*self.scale)

        # Frame 0: Photoshop info
        frame0 = ttk.Frame(self.root)
        frame0.grid(row=0, column=0, sticky='ew', padx=padx, pady=pady)

        # 创建两个Label，分别用于显示版本和文档信息
        self.ps_version_label = ttk.Label(frame0, text="", foreground='blue')
        self.ps_version_label.grid(row=0, column=0, padx=padx, sticky='w')

        self.active_doc_label = ttk.Label(frame0, text="", foreground='blue')
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
        auto_save_check = ttk.Checkbutton(combined_frame, text="自动备份", variable=self.auto_save_var)
        auto_save_check.grid(row=0, column=2, padx=padx)

        # 分隔符
        separator = ttk.Separator(combined_frame, orient='vertical')
        separator.grid(row=0, column=3, sticky='ns')

        # 备份保留设置
        backup_retention_days_label = ttk.Label(combined_frame, text="保留时间 \\ 天: ")
        backup_retention_days_label.grid(row=0, column=4, padx=padx)
        self.backup_retention_days_entry = ttk.Spinbox(combined_frame, from_=1, to=999, textvariable=self.backup_retention_days_var, width=padx)
        self.backup_retention_days_entry.grid(row=0, column=5, sticky='ew')

        self.clean_button = ttk.Button(combined_frame, text="清理", command=self.start_cleanup)
        self.clean_button.grid(row=0, column=6, padx=padx)

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

        # Set the weight of the first column to a large number
        frame3.columnconfigure(0, weight=1)

        link = ttk.Label(frame3, text="v0.1.0", cursor="hand2")
        link.grid(row=0, column=0, padx=padx, sticky='ew')
        link.bind("<Button-1>", self.open_link)

        save_button = ttk.Button(frame3, text="立即备份", command=self.start_save)
        save_button.grid(row=0, column=1, padx=padx)
        minimize_button = ttk.Button(frame3, text="最小化", command=self.hide_window)
        minimize_button.grid(row=0, column=2, padx=padx)
        exit_button = ttk.Button(frame3, text="退出", command=self.exit_button_clicked)
        exit_button.grid(row=0, column=3, padx=padx)

        # Frame 5: Blank
        frame5 = ttk.Frame(self.root)
        frame5.grid(row=5, column=0, sticky='ew', padx=padx, pady=pady/2)

        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

    def initialize_variables(self):
        """Initializes the Tkinter variables with values from the configuration."""
        # Load values from the configuration manager
        folder_path = self.config_manager.get_folder_path()
        auto_save_enabled, auto_save_interval = self.config_manager.get_auto_save_settings()
        backup_retention_days = self.config_manager.get_backup_retention_days()

        # Assign values to Tkinter variables
        self.folder_path_var = tk.StringVar(value=folder_path)
        self.auto_save_interval_var = tk.IntVar(value=auto_save_interval)
        self.auto_save_var = tk.IntVar(value=int(auto_save_enabled))
        self.backup_retention_days_var = tk.IntVar(value=backup_retention_days)

        # Adding traces
        self.folder_path_var.trace_add("write", self.update_folder_path)
        self.auto_save_interval_var.trace_add("write", self.validate_auto_save_interval)
        self.auto_save_var.trace_add("write", self.update_auto_save_enabled)
        self.backup_retention_days_var.trace_add("write", self.validate_backup_retention_days)

    def update_folder_path(self, *args):
        self.config_manager.set_folder_path(self.folder_path_var.get())

    def validate_auto_save_interval(self, *args):
        interval = self.auto_save_interval_var.get()
        if isinstance(interval, int) and interval > 0:
            self.update_auto_save_interval()
        else:
            messagebox.showerror("Invalid Input", "Auto save interval must be a positive integer.")
            self.auto_save_interval_var.set(self.config_manager.get_auto_save_settings()[1])  # Reset to last valid value

    def update_auto_save_interval(self):
        self.config_manager.set_auto_save_settings(self.auto_save_var.get() == 1, self.auto_save_interval_var.get())

    def update_auto_save_enabled(self, *args):
        self.config_manager.set_auto_save_settings(self.auto_save_var.get() == 1, self.auto_save_interval_var.get())
        self.handle_auto_save()

    def validate_backup_retention_days(self, *args):
        days = self.backup_retention_days_var.get()
        if isinstance(days, int) and days >= 0:
            self.update_backup_retention_days()
        else:
            messagebox.showerror("Invalid Input", "Backup retention days must be a non-negative integer.")
            self.backup_retention_days_var.set(self.config_manager.get_backup_retention_days())

    def update_backup_retention_days(self):
        self.config_manager.set_backup_retention_days(self.backup_retention_days_var.get())

    def update_ps_info(self):
        ps_version, active_doc = get_ps_info()
        if self.ps_version_label.winfo_exists() and self.active_doc_label.winfo_exists():
            self.ps_version_label.config(text=f"PS Version: {ps_version}")
            self.active_doc_label.config(text=f"Active PSD: {active_doc}")

    def update_ps_info_periodically(self):
        self.update_ps_info()  # 更新 Photoshop 信息
        self.root.after(2000, self.update_ps_info_periodically)  # 每隔2秒调用一次

    def finish_cleanup(self, cleaned_files_count):
        self.clean_info_label.config(text=f"清理完成，共删除 {cleaned_files_count} 个文件。")
        # 清理完成，启用按钮并更新提示信息
        self.clean_button.config(state=tk.NORMAL)
        self.root.after(8000, lambda: self.clean_info_label.config(text=""))

    def start_cleanup(self):
        if not messagebox.askyesno("确认", f"您确定要清理吗？本操作会清理备份文件夹中所有以_psbackup结尾，并且时间戳大于{str(self.backup_retention_days_var.get())}天的psd文件！为保险起见，请尽量不要在备份目录存放其他文件！！！"):
            return
        if self.cleanup_thread is not None and self.cleanup_thread.is_alive():
            logging.info("Cleanup already in progress.")
            return

        self.clean_button.config(state=tk.DISABLED)
        self.clean_info_label.config(text="清理中...")

        self.cleanup_thread = threading.Thread(target=lambda: clean_old_backups(self.folder_path_var.get(), self.backup_retention_days_var.get(), self.finish_cleanup))
        self.cleanup_thread.daemon = True
        self.cleanup_thread.start()

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.config_manager.set_folder_path(folder_selected)

    def open_link(self, event):
        webbrowser.open_new(r"https://github.com/aimkiray/psb")

    def start_save(self, silent=False):
        if self.auto_save_thread and self.auto_save_thread.is_alive():
            logging.info("An automatic backup operation is already in progress, please try again later or adjust the automatic backup interval.")
            if not silent:
                messagebox.showinfo("信息", "已有一个备份操作正在进行，请稍后再试。")
            return
        
        folder_path = self.folder_path_var.get()
        if not folder_path:
            logging.error("No backup directory specified.")
            if not silent:
                messagebox.showerror("错误", "请先选择文件保存目录。")
            return
        
        if silent:
            # 后台自动备份，不需要提示
            self.run_auto_save(folder_path)
        else:
            backup_result = save_psd_as(folder_path)

            if backup_result:
                messagebox.showinfo("成功", f"文件已成功备份到：{backup_result}")
            else:
                messagebox.showerror("错误", f"保存文件时出错，请检查日志文件")

    def handle_auto_save(self):
        if self.auto_save_var.get() == 1:
            try:
                interval = int(self.auto_save_interval_var.get())
                interval_ms = interval * 60 * 1000  # 将分钟转换为毫秒
                if self.auto_save_job is not None:
                    self.root.after_cancel(self.auto_save_job)
                self.auto_save_schedule(interval_ms)
                logging.info(f"Auto-save scheduled every {interval} minutes.")
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字")
        else:
            if self.auto_save_job is not None:
                self.root.after_cancel(self.auto_save_job)
                self.auto_save_job = None
            logging.info("Auto-save disabled.")

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
        self.root.withdraw()

    def show_window(self):
        self.root.deiconify()

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
      

def main():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
