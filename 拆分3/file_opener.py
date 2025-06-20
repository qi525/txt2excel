import os
import sys
import time
import subprocess
from pathlib import Path
from typing import List
from file_system_utils import normalize_drive_letter


def open_output_files_automatically(file_paths: List[Path], logger_obj):
    """
    根据用户设置自动打开生成的输出文件（Excel和Log文件）。
    Args:
        file_paths (List[Path]): 包含要打开的文件路径的列表。
        logger_obj (logger): Loguru logger 实例。
    """
    if os.getenv("DISABLE_AUTO_OPEN", "0") == "1":
        logger_obj.info("已禁用自动打开文件功能。")
        return

    # 定义延迟时间 (秒)
    OPEN_FILE_DELAY_SECONDS = 2 # 你建议的2秒延迟

    for file_path in file_paths:
        # --- 主要修改点 START ---
        # 在尝试打开文件前增加延迟
        logger_obj.debug(f"尝试打开文件 '{normalize_drive_letter(str(file_path))}' 前，等待 {OPEN_FILE_DELAY_SECONDS} 秒。")
        time.sleep(OPEN_FILE_DELAY_SECONDS)
        # --- 主要修改点 END ---

        actual_path_to_open = file_path
        # 特殊处理 Loguru 压缩后的日志文件
        if file_path.suffix == '.txt' and not file_path.exists():
            zip_path = file_path.with_suffix('.zip')
            if zip_path.exists():
                actual_path_to_open = zip_path
                logger_obj.info(f"日志文件 '{normalize_drive_letter(str(file_path))}' 不存在，尝试打开压缩文件: {normalize_drive_letter(str(zip_path))}")

        if not actual_path_to_open.exists():
            logger_obj.warning(f"警告: 无法自动打开文件 '{normalize_drive_letter(str(actual_path_to_open))}'，因为文件不存在。")
            continue

        try:
            normalized_path = normalize_drive_letter(str(actual_path_to_open))
            logger_obj.info(f"自动打开: {normalized_path}")
            if sys.platform == "win32":
                os.startfile(normalized_path)
            elif sys.platform == "darwin": # macOS
                subprocess.run(['open', normalized_path], check=True)
            else: # Linux
                subprocess.run(['xdg-open', normalized_path], check=True)
        except FileNotFoundError:
            logger_obj.error(f"错误: 无法找到打开文件 '{normalized_path}' 的应用程序。请手动打开。")
        except Exception as e:
            logger_obj.error(f"错误: 自动打开文件 '{normalized_path}' 时发生意外错误: {e}")