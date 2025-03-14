import os
from datetime import datetime
import win32com.client
import pythoncom
import logging


def save_psd_as(folder_path):
    if not folder_path:
        logging.error("No backup directory specified.")
        return False

    folder_path = os.path.normpath(folder_path)

    try:
        # 尝试获取 Photoshop 应用程序的引用，如果未打开则不会创建新实例
        psApp = win32com.client.GetActiveObject("Photoshop.Application")
        if psApp.Documents.Count == 0:
            logging.error("No open documents in Photoshop.")
            return False

        doc = psApp.ActiveDocument
        original_name = doc.Name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_filename = f"{original_name.split('.')[0]}_{timestamp}_psbackup.psd"
        full_path = os.path.join(folder_path, new_filename)

        psdOptions = win32com.client.Dispatch("Photoshop.PhotoshopSaveOptions")
        psdOptions.layers = True

        doc.SaveAs(full_path, psdOptions, True)
        logging.info(
            f"The file has been successfully backed up to: {full_path}")
        return full_path
    except Exception as e:
        logging.error(f"No Photoshop process found. {e}")
        return False

def thread_save_psd_as(folder_path):
    pythoncom.CoInitialize()  # 初始化 COM
    try:
        # 执行 COM 操作
        save_psd_as(folder_path)
    finally:
        pythoncom.CoUninitialize()  # 清理 COM

def get_ps_info():
    try:
        # 尝试连接到正在运行的 Photoshop 实例
        ps_app = win32com.client.GetActiveObject("Photoshop.Application")
        ps_version = ps_app.Version
        active_doc = ps_app.ActiveDocument.Name if ps_app.Documents.Count > 0 else "No open documents"
    except Exception:
        ps_version = "Unavailable"
        active_doc = "No open documents"
    return ps_version, active_doc
