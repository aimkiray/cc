import psutil
import os

def get_current_process_name():
    """获取当前进程的名称"""
    current_pid = os.getpid()
    process = psutil.Process(current_pid)
    return process.name()  

def is_process_running(pid):
    """检查给定的 PID 是否在运行."""
    try:
        p = psutil.Process(pid)
        return p.name() == get_current_process_name()
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        return False

