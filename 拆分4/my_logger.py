# my_logger.py (Loguru 重构版本 - 优化后)
from loguru import logger
from pathlib import Path
import os
import datetime
import sys

# --- 新增功能点：用于存储错误日志文件路径的全局变量 ---
_error_log_file_path: Path = Path("N/A") # 初始化一个默认值，防止未设置时访问

def setup_logger(log_directory: Path) -> Path: # 修改函数签名，明确返回 Path 类型
    """
    配置 Loguru 日志器，设置多个日志输出目标。
    返回错误和警告日志文件的路径。
    """
    # !!! 新增：在函数开始时移除所有现有处理器，确保每次调用都重新配置 !!!
    logger.remove()

    global _error_log_file_path # 声明要修改全局变量

    # 确保日志目录存在
    try:
        if not log_directory.exists():
            os.makedirs(log_directory)
            logger.info(f"已创建日志文件夹: {log_directory}")
    except Exception as e:
        logger.critical(f"关键错误: 无法创建日志文件夹 {log_directory}. 日志将仅打印到控制台. 错误: {e}")
        # 如果创建日志文件夹失败，返回一个无效路径或None，让调用者处理
        return Path("invalid_log_path") # 或者直接 raise e，取决于你希望如何处理这种核心错误

    # 1. 配置主程序日志 (INFO 及以上，到独立文件)
    main_log_file_name = f"main_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    main_log_path = log_directory / main_log_file_name

    logger.add(
        sink=main_log_path,
        level="INFO",
        format="{time:YYYY-MM-DD HH:mm:ss} [{level}] {message}",
        rotation="10 MB",  # 每10MB文件大小轮转
        retention="7 days", # 保留7天的日志文件
        compression="zip",  # 压缩旧日志文件
        enqueue=True,       # 启用多进程/多线程安全
        backtrace=True,     # 在异常时显示详细回溯
        diagnose=True       # 在异常时显示变量值
    )

    # 2. 配置错误和警告日志 (WARNING 及以上，到独立文件)
    error_warning_log_file_name = f"error_warning_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    error_warning_log_path = log_directory / error_warning_log_file_name

    logger.add(
        sink=error_warning_log_path,
        level="WARNING", # 捕获 WARNING 及以上级别的日志
        format="{time:YYYY-MM-DD HH:mm:ss} [{level}] {message}",
        rotation="10 MB",
        retention="7 days",
        compression="zip",
        enqueue=True,
        backtrace=True,
        diagnose=True,
        filter=lambda record: record["level"].name in ["WARNING", "ERROR", "CRITICAL"] # 只包含这几个级别
    )

    # 3. 配置控制台输出（仅 INFO 及以上，不写入文件）
    # 确保控制台输出只有在文件日志设置完成后才添加，避免重复
    logger.add(
        sys.stderr, # 输出到标准错误，通常是控制台
        level="INFO",
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> <level>[{level}]</level> <level>{message}</level>",
        colorize=True,
        filter=lambda record: record["level"].name in ["INFO", "SUCCESS"] # 只显示 INFO 和 SUCCESS 级别的消息
    )
    # 更新全局错误日志文件路径变量
    _error_log_file_path = error_warning_log_path

    return error_warning_log_path

def get_error_log_file_path() -> Path:
    """
    获取当前会话的错误日志文件的路径。
    """
    return _error_log_file_path