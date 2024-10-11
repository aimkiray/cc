import psutil

def is_process_running(pid):
    """检查给定的 PID 是否在运行."""
    try:
        p = psutil.Process(pid)
        return p.is_running()
    except psutil.NoSuchProcess:
        return False

