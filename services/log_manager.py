# services/log_manager.py
import datetime
from pathlib import Path
import os
import sys

class LogManager:
    """
    负责程序的日志记录。
    """
    def __init__(self, log_directory: Path, log_file_name: str = None):
        self.log_directory = log_directory
        # 如果没有指定日志文件名，则生成一个默认的（主日志文件）
        if log_file_name is None:
            # 这里的日志文件名应该基于当前时间，确保唯一性，避免重名冲突
            self.log_file_path = self.log_directory / f"main_scan_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        else:
            self.log_file_path = self.log_directory / log_file_name
        
        self.file_handle = None # 初始化文件句柄为None
        self._open_log_file() # 尝试打开日志文件

    def _open_log_file(self):
        """
        尝试打开日志文件，如果失败则打印到控制台。
        """
        try:
            # 使用 'a' 模式（append），如果文件不存在则创建
            self.file_handle = open(self.log_file_path, 'a', encoding='utf-8')
        except Exception as e:
            print(f"Critical Error: Failed to open log file at {self.log_file_path}. All subsequent logs will be printed to console only. Error: {e}")
            self.file_handle = None

    def write_log(self, message: str):
        """
        写入日志信息到文件，如果文件句柄无效则打印到控制台。
        """
        timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        log_message = f"{timestamp} {message}\n"
        
        if self.file_handle:
            try:
                self.file_handle.write(log_message)
                self.file_handle.flush() # 立即将缓冲区内容写入文件
            except Exception as e:
                print(f"Critical Error: Failed to write log to {self.log_file_path}. Message: {message}. Error: {e}")
                if self.file_handle:
                    self.file_handle.close()
                self.file_handle = None
                print(f"Log message redirected to console: {log_message.strip()}")
        else:
            print(f"No log file handle. Printing to console: {log_message.strip()}")

    def close(self):
        """
        关闭日志文件句柄。
        """
        if self.file_handle:
            try:
                self.file_handle.close()
                self.file_handle = None
            except Exception as e:
                print(f"Error closing log file {self.log_file_path}. Error: {e}")

    def __del__(self):
        """
        析构函数，确保在对象被销毁时关闭文件句柄。
        """
        self.close()