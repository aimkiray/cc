import os
import sys

def resource_path(relative_path):
    """ 获取资源的绝对路径。用于 PyInstaller 打包后资源文件的访问 """
    try:
        # PyInstaller 创建的临时文件夹
        base_path = sys._MEIPASS
    except Exception:
        # 正常执行时的路径
        base_path = os.path.abspath(".")

    relative_path = os.path.normpath(relative_path)

    return os.path.join(base_path, relative_path)