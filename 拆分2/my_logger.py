# my_logger.py

import os
import datetime
from pathlib import Path
from typing import Optional

# 移除这行： from utils import normalize_drive_letter # 确保 utils.py 中有此函数

# 将 normalize_drive_letter 函数定义直接放在这里
def normalize_drive_letter(path_str: str) -> str:
    """
    标准化驱动器盘符大小写，确保路径在不同系统上的一致性。
    例如，将 'c:\\path' 转换为 'C:\\path'。
    """
    if os.name == 'nt' and len(path_str) > 1 and path_str[1] == ':':
        return path_str[0].upper() + path_str[1:]
    return path_str

class LogManager:
    """
    负责程序的日志记录。
    """
    def __init__(self, log_directory: Path, log_file_name: str = None, 
                 error_log_manager: Optional['LogManager'] = None,
                 is_error_log_manager: bool = False):
        self.log_directory = log_directory
        self.log_file_path = None
        self.file_handle = None
        self.error_log_manager = error_log_manager
        self.is_error_log_manager = is_error_log_manager
        self._is_initialized = False

        # 尝试创建日志目录
        try:
            if not self.log_directory.exists():
                os.makedirs(self.log_directory)
                if not self.is_error_log_manager:
                    print(f"已创建日志文件夹: {normalize_drive_letter(str(self.log_directory))}")
                    self.write_log(f"已创建日志文件夹: {normalize_drive_letter(str(self.log_directory))}", level="INFO", to_error_log=False)
        except Exception as e:
            print(f"关键错误: 无法创建日志文件夹 {normalize_drive_letter(str(self.log_directory))}. 日志将仅打印到控制台. 错误: {e}")
            if not self.is_error_log_manager:
                self.write_log(f"关键错误: 无法创建日志文件夹 {normalize_drive_letter(str(self.log_directory))}. 日志将仅打印到控制台. 错误: {e}", level="CRITICAL", to_error_log=False)
            self.log_directory = None

        if self.log_directory:
            if log_file_name is None:
                self.log_file_path = self.log_directory / f"main_program_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            else:
                self.log_file_path = self.log_directory / log_file_name
        else:
            self.log_file_path = None

    def _ensure_log_file_open(self):
        """
        确保日志文件已打开。如果未打开，则尝试打开。
        对于错误日志管理器，只有当它被写入WARNING/ERROR/CRITICAL级别日志时才真正打开。
        对于其他日志管理器，只要调用此方法就尝试打开。
        """
        if self.file_handle is None and self.log_file_path:
            try:
                self.file_handle = open(self.log_file_path, 'a', encoding='utf-8')
                if not self._is_initialized:
                    if not self.is_error_log_manager:
                        print(f"日志文件已打开: {normalize_drive_letter(str(self.log_file_path))}")
                        self.file_handle.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [INFO] 日志文件已打开: {normalize_drive_letter(str(self.log_file_path))}\n")
                        self.file_handle.flush()
                    self._is_initialized = True
            except Exception as e:
                print(f"关键错误: 无法打开日志文件 {normalize_drive_letter(str(self.log_file_path))}. 所有后续日志将仅打印到控制台. 错误: {e}")
                self.file_handle = None


    def write_log(self, message: str, level: str = "INFO", to_file_only: bool = False, to_error_log: bool = True):
        """
        写入日志信息到文件，如果文件句柄无效则打印到控制台。
        Args:
            message (str): 日志消息。
            level (str): 日志级别 (INFO, WARNING, ERROR, CRITICAL).
            to_file_only (bool): 如果为True，则只写入文件，不打印到控制台。
            to_error_log (bool): 如果为True且存在error_log_manager，则将WARNING/ERROR/CRITICAL日志写入错误日志。
        """
        timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        log_message = f"{timestamp} [{level}] {message}"
        
        if to_error_log and self.error_log_manager and self.error_log_manager is not self and level in ["WARNING", "ERROR", "CRITICAL"]:
            self.error_log_manager.write_log(message, level=level, to_file_only=True, to_error_log=False)

        should_write_to_this_file = False
        if not self.is_error_log_manager:
            should_write_to_this_file = True
        elif self.is_error_log_manager and level in ["WARNING", "ERROR", "CRITICAL"]:
            should_write_to_this_file = True

        if should_write_to_this_file:
            self._ensure_log_file_open()
            if self.file_handle:
                try:
                    self.file_handle.write(log_message + "\n")
                    self.file_handle.flush()
                    if not to_file_only:
                        print(log_message)
                except Exception as e:
                    print(f"关键错误: 写入日志文件 {normalize_drive_letter(str(self.log_file_path))} 失败. 消息: {message}. 错误: {e}")
                    if self.file_handle:
                        self.file_handle.close()
                    self.file_handle = None
                    if not to_file_only:
                        print(f"日志消息重定向到控制台: {log_message}")
            else:
                if not to_file_only:
                    print(log_message)
        else:
            if not to_file_only:
                print(log_message)

    def close(self):
        """
        关闭日志文件句柄。
        """
        if self.file_handle:
            try:
                self.file_handle.close()
                self.file_handle = None
                self._is_initialized = False
                if not self.is_error_log_manager:
                    print(f"日志文件已关闭: {normalize_drive_letter(str(self.log_file_path))}")
            except Exception as e:
                print(f"关闭日志文件 {normalize_drive_letter(str(self.log_file_path))} 失败. 错误: {e}")

    def __del__(self):
        """
        析构函数，确保在对象被销毁时关闭文件句柄。
        """
        self.close()