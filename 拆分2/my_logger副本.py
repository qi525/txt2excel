# my_logger.py (使用 Loguru 重构)

from loguru import logger
from pathlib import Path
import os
import sys

# Loguru 默认会向 stderr 输出，通常也带有颜色。
# 我们需要移除默认处理器，以便完全控制日志输出。
logger.remove()

def setup_logger(log_directory: Path):
    """
    配置 Loguru 日志。
    - 主日志：INFO 级别及以上，写入 'main_program_log_YYYYMMDD_HHMMSS.txt'
      并自动轮转（例如每天或达到大小），保留旧日志。
    - 错误/警告日志：WARNING 级别及以上，写入 'error_warning_log_YYYYMMDD_HHMMSS.txt'。
      这个文件只有在有警告或错误时才会创建。
    """
    # 确保日志目录存在
    try:
        if not log_directory.exists():
            os.makedirs(log_directory)
            logger.info(f"已创建日志文件夹: {log_directory}") # Loguru 打印，可以带有颜色
    except Exception as e:
        # 目录创建失败是关键错误，直接打印到控制台，不依赖日志文件
        print(f"关键错误: 无法创建日志文件夹 {log_directory}. 日志将仅打印到控制台. 错误: {e}")
        return # 无法创建目录，后续日志配置也可能失败

    # 1. 配置主程序日志 (INFO 及以上，到文件和控制台)
    # 文件名格式：main_program_log_YYYYMMDD_HHMMSS.txt
    main_log_file_name = f"main_program_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    main_log_path = log_directory / main_log_file_name

    logger.add(
        sink=main_log_path,
        level="INFO",
        format="{time:YYYY-MM-DD HH:mm:ss} [{level}] {message}",
        rotation="10 MB", # 例如，文件达到10MB时轮转
        retention="7 days", # 保留7天的日志文件
        compression="zip", # 轮转后压缩旧日志文件
        enqueue=True, # 启用多进程/线程安全
        backtrace=True, # 打印详细回溯
        diagnose=True,  # 打印局部变量值（用于调试，生产环境可能关闭）
        # filter=lambda record: record["level"].no >= logger.level("INFO").no # 明确过滤级别
    )
    logger.info(f"主日志文件已配置: {main_log_path}")

    # 2. 配置错误和警告日志 (WARNING 及以上，到单独的文件)
    # 这个 sink 只有当接收到 WARNING/ERROR/CRITICAL 级别日志时才会写入
    error_warning_log_file_name = f"error_warning_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    error_warning_log_path = log_directory / error_warning_log_file_name

    # Loguru 的 filter 参数可以实现你“按需生成错误日志”的需求
    # 当没有 WARNING/ERROR/CRITICAL 级别的日志时，这个文件不会被创建
    logger.add(
        sink=error_warning_log_path,
        level="WARNING", # 只记录 WARNING, ERROR, CRITICAL 级别
        format="{time:YYYY-MM-DD HH:mm:ss} [{level}] {message}",
        rotation="5 MB", # 错误日志也可以轮转
        retention="30 days", # 错误日志保留时间可以更长
        compression="zip",
        enqueue=True,
        backtrace=True,
        diagnose=True,
        # filter=lambda record: record["level"].no >= logger.level("WARNING").no
    )
    # 这里不需要像你之前那样打印“错误和警告日志文件已创建”信息，
    # 因为 Loguru 会延迟创建，并且主日志已经记录了相关信息。

    # 配置控制台输出（可选，如果不想让主日志文件也输出到控制台）
    # logger.add(sys.stderr, level="INFO", format="{time:YYYY-MM-DD HH:mm:ss} [{level}] {message}", colorize=True)

    # 假设你的 normalize_drive_letter 仍然在 my_logger.py 中
    # 或者你可以选择在 main.py 直接使用 str(Path_object) 传递给 logger
    # Loguru 内部通常会处理路径显示，所以 normalize_drive_letter 可能不再那么必要
    # 但如果为了统一风格，可以保留。

# 注意：Loguru 的 logger 对象是全局的，你可以直接从任何文件导入 `logger` 并使用它。
# 所以不再需要 `LogManager` 类。
# 调用方式：
# from my_logger import logger, setup_logger
# setup_logger(Path("logs")) # 在程序启动时调用一次
# logger.info("程序启动")
# logger.warning("这是一个警告")
# logger.error("这是一个错误")
# try:
#    1 / 0
# except Exception:
#    logger.exception("发生了一个除零错误") # 自动捕获并打印详细堆栈