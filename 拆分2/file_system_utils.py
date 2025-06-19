import os
import shutil
from pathlib import Path
from typing import Tuple
from loguru import logger # 如果这些函数内部使用了 logger 对象

# 可能还需要从 my_logger 导入 normalize_drive_letter
from my_logger import normalize_drive_letter # 因为 copy_file, validate_directory, create_directory_if_not_exists 内部使用了它
# --- File Operations ---
def validate_directory(path: Path, logger_obj: logger) -> bool:
    """
    验证给定的路径是否是一个存在的目录。
    """
    if not path.is_dir():
        if logger_obj:
            logger_obj.warning(f"验证失败: 目录不存在或不是一个目录: {normalize_drive_letter(str(path))}")#warning
        return False
    return True

# def create_directory_if_not_exists(directory_path: Path, logger_obj: Optional[logger]) -> bool:
def create_directory_if_not_exists(directory_path: Path, logger_obj) -> bool: # 更改参数名为 logger_obj，并移除 logger 类型提示
    """
    如果指定目录不存在，则创建它。
    Args:
        directory_path (Path): 要创建的目录路径。
        logger_obj (Optional[logger]): 日志管理器实例，可选。
    Returns:
        bool: 如果目录存在或成功创建，则返回True；否则返回False。
    """
    if not directory_path.exists():
        try:
            os.makedirs(directory_path)
            if logger_obj:
                #logger_obj.info(f"已创建目录: {normalize_drive_letter(str(directory_path))}"
                logger_obj.info(f"已创建目录: {normalize_drive_letter(str(directory_path))}") # 替换为 Loguru 的 info 方法
            return True
        except OSError as e:
            if logger_obj:
                #logger_obj.info(f"错误: 无法创建目录 {normalize_drive_letter(str(directory_path))}: {e}")#error
                logger_obj.error(f"创建目录失败 {normalize_drive_letter(str(directory_path))}: {e}") # 替换为 Loguru 的 error 方法
            print(f"错误: 无法创建文件夹 {directory_path}。错误: {e}")
            return False
    return True

def copy_file(source_path: Path, destination_path: Path, logger_obj:logger) -> bool:
    """
    复制文件从源路径到目标路径。
    增加对权限错误的捕获和提示。
    """
    if not source_path.exists():
        if logger_obj:
            logger_obj.error(f"错误: 源文件不存在，无法复制: {normalize_drive_letter(str(source_path))}")#error
        print(f"错误: 源文件不存在，无法复制: {source_path}")
        return False

    try:
        shutil.copy2(str(source_path), str(destination_path)) 
        if logger_obj:
            logger_obj.info(f"已复制 '{normalize_drive_letter(str(source_path))}' 到 '{normalize_drive_letter(str(destination_path))}'")
        return True
    except PermissionError as e:
        if logger_obj:
            logger_obj.critical(
                f"权限错误: 复制文件从 '{normalize_drive_letter(str(source_path))}' 到 '{normalize_drive_letter(str(destination_path))}' 失败: {e}. 请确保目标文件未被其他程序（如Excel）占用。")
        print(f"错误: 权限拒绝！无法复制文件到 '{destination_path}'。请确保该文件未被其他程序（如Excel）打开。错误: {e}")
        return False
    except Exception as e:
        if logger_obj:
            logger_obj.error(f"错误: 复制文件从 '{normalize_drive_letter(str(source_path))}' 到 '{normalize_drive_letter(str(destination_path))}' 失败: {e}")#error
        print(f"错误: 无法复制文件从 '{source_path}' 到 '{destination_path}'。错误: {e}")
        return False

def get_file_details(file_path: Path) -> Tuple[str, str]:
    """
    获取文件的名称（不含扩展名）和扩展名。
    """
    return file_path.stem, file_path.suffix
