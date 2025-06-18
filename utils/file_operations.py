# utils/file_operations.py
import os
import shutil
from pathlib import Path
from typing import Optional # Import Optional for type hinting

# 从 services.log_manager 导入 LogManager，以便在需要时进行类型提示
# 由于LogManager是在services文件夹中，这里需要相对导入或绝对导入
# 假设utils和services在同一层级，可以使用绝对导入
try:
    from services.log_manager import LogManager
except ImportError:
    # Fallback for environments where direct import might be tricky
    # This scenario is less likely in our current structure but good practice
    LogManager = None 


def validate_directory(path: Path, log_manager: Optional[LogManager]) -> bool:
    """
    验证给定的路径是否是一个存在的目录。
    """
    if not path.is_dir():
        if log_manager:
            log_manager.write_log(f"Validation failed: Directory does not exist or is not a directory: {path}")
        return False
    return True

def create_directory_if_not_exists(directory_path: Path, log_manager: Optional[LogManager]) -> bool:
    """
    如果指定目录不存在，则创建它。
    Args:
        directory_path (Path): 要创建的目录路径。
        log_manager (Optional[LogManager]): 日志管理器实例，可选。
    Returns:
        bool: 如果目录存在或成功创建，则返回True；否则返回False。
    """
    if not directory_path.exists():
        try:
            os.makedirs(directory_path)
            if log_manager:
                log_manager.write_log(f"Created directory: {directory_path}")
            print(f"已创建文件夹: {directory_path}")
            return True
        except OSError as e:
            if log_manager:
                log_manager.write_log(f"Error creating directory {directory_path}: {e}")
            print(f"错误: 无法创建文件夹 {directory_path}。错误: {e}")
            return False
    return True

def copy_file(source_path: Path, destination_path: Path, log_manager: Optional[LogManager]) -> bool:
    """
    复制文件从源路径到目标路径。
    """
    try:
        shutil.copy2(str(source_path), str(destination_path)) # shutil.copy2 复制文件和元数据
        if log_manager:
            log_manager.write_log(f"Copied '{source_path}' to '{destination_path}'")
        return True
    except Exception as e:
        if log_manager:
            log_manager.write_log(f"Error copying file from '{source_path}' to '{destination_path}': {e}")
        print(f"错误: 无法复制文件从 '{source_path}' 到 '{destination_path}'。错误: {e}")
        return False

def get_file_details(file_path: Path) -> tuple[str, str]:
    """
    获取文件的名称（不含扩展名）和扩展名。
    """
    return file_path.stem, file_path.suffix