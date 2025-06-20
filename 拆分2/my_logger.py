# my_logger.py (Loguru 重构版本 - 确保这是你的实际内容)

from loguru import logger
from pathlib import Path
import os
import datetime # 确保这个导入也在，用于文件名


# --- 新增功能点：用于存储错误日志文件路径的全局变量 ---
_error_log_file_path: Path = Path("N/A") # 初始化一个默认值，防止未设置时访问


logger.remove() # 移除 Loguru 默认处理器


def setup_logger(log_directory: Path):
    # ... Loguru 的配置代码，保持不变 ...
    # 你的 Loguru my_logger.py 代码应该从这里开始

    global _error_log_file_path # 声明要修改全局变量

    try:
        if not log_directory.exists():
            os.makedirs(log_directory)
            logger.info(f"已创建日志文件夹: {log_directory}")
    except Exception as e:
        print(f"关键错误: 无法创建日志文件夹 {log_directory}. 日志将仅打印到控制台. 错误: {e}")
        return

    # 1. 配置主程序日志 (INFO 及以上，到文件和控制台)
    main_log_file_name = f"main_program_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    main_log_path = log_directory / main_log_file_name

    logger.add(
        sink=main_log_path,
        level="INFO",
        format="{time:YYYY-MM-DD HH:mm:ss} [{level}] {message}",
        rotation="10 MB",  # 例如，每当文件大小达到10MB时进行轮转
        retention="7 days", # 例如，只保留7天的日志文件
        compression="zip",  # 轮转后的旧日志文件会被压缩成zip
        enqueue=True,       # 启用多进程安全
        backtrace=True,     # 在异常时显示完整的堆栈跟踪
        diagnose=True       # 在异常时显示变量值
    )

    # 2. 配置错误和警告日志 (WARNING 及以上，到独立文件)
    error_warning_log_file_name = f"error_warning_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    error_warning_log_path = log_directory / error_warning_log_file_name

    logger.add(
        sink=error_warning_log_path,
        level="WARNING", # 记录 WARNING 级别及以上的日志
        format="{time:YYYY-MM-DD HH:mm:ss} [{level}] {message}",
        rotation="5 MB",
        retention="7 days",
        compression="zip",
        enqueue=True,
        backtrace=True,
        diagnose=True
    )

    # --- 新增功能点：保存错误日志文件路径到全局变量 ---
    _error_log_file_path = error_warning_log_path
    logger.info("Loguru logger initialized and configured.")


# --- 新增功能点：提供获取错误日志文件路径的函数 ---
def get_error_log_file_path() -> Path:
    """
    获取错误和警告日志文件的路径。
    """
    return _error_log_file_path