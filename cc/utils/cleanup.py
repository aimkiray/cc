import os
import time
import logging

def clean_old_backups(folder_path, backup_retention_days, on_cleanup_complete):
    now = time.time()
    cleaned_files_count = 0  # 文件计数
    try:
        for filename in os.listdir(folder_path):
            if not filename.endswith('_psbackup.psd'):
                continue
            full_path = os.path.join(folder_path, filename)
            if os.path.getmtime(full_path) < now - backup_retention_days * 86400:
                os.remove(full_path)
                cleaned_files_count += 1
    except PermissionError:
        logging.error(f"Permission denied: cannot read or remove files in {folder_path}")
    except Exception as e:
        logging.error(f"Error during old backups cleaning: {str(e)}")

    # 清理完成后调用回调函数
    if on_cleanup_complete:
        time.sleep(1)
        on_cleanup_complete(cleaned_files_count)
