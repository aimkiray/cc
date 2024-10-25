import os
import tkinter as tk
from tempfile import gettempdir
import tkinter as tk
from tkinter import messagebox
from .ui.main_window import MainWindow
from .utils.logger import setup_logging
from .utils.process import is_process_running

def main():
    setup_logging()

    # 检查是否已经有实例运行
    pid_file = os.path.join(gettempdir(), 'psbackup.lock')
    try:
        with open(pid_file, 'r') as f:
            pid = int(f.read())
        if is_process_running(pid):
            root = tk.Tk()
            root.withdraw()  # 隐藏主窗口
            messagebox.showwarning("Warning", "Another instance of the application is already running.")
            root.destroy()
            return
    except FileNotFoundError:
        pass

    # 写入当前进程 PID
    with open(pid_file, 'w') as f:
        f.write(str(os.getpid()))

    try:
        # 创建并运行 GUI
        app = MainWindow() 
    finally:
        # 清理 PID 文件
        os.remove(pid_file)

if __name__ == '__main__':
    main()