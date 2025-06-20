import os
import sys
import shutil
from pathlib import Path
from typing import Tuple,List
from loguru import logger # 导入 Loguru logger

# file_namer.py
import re # 导入re模块用于正则表达式
import hashlib # 导入hashlib用于生成文件夹名的哈希值

# --- NEW FUNCTION: Generate a safe and identifiable folder prefix for filenames ---
def generate_folder_prefix(folder_path: Path) -> str:
    """
    根据文件夹路径生成一个安全且可识别的前缀，用于文件名。
    原理：
        1. 获取文件夹的basename（即文件夹本身的名称）。
        2. 如果 basename 包含中文或特殊字符，为了确保文件名在各种文件系统中的兼容性，
           我们使用该 basename 的MD5哈希值的前8位作为唯一标识。
        3. 如果 basename 只包含ASCII字符（数字、字母、下划线、短横线），
           则直接使用 basename。
        4. 最终前缀会限制长度，避免文件名过长。
    Args:
        folder_path (Path): 文件夹的Path对象。
    Returns:
        str: 一个安全且短小的字符串，用于作为文件名前缀。
    """
    folder_name = folder_path.name
    # 检查是否包含非ASCII字符（例如中文），或者其他不适合作为文件名的字符
    if not re.fullmatch(r'[\w.-]+', folder_name): # 允许字母、数字、下划线、点、短横线
        # 如果包含特殊字符或中文，则使用哈希值
        return hashlib.md5(folder_name.encode('utf-8')).hexdigest()[:8]
    else:
        # 否则，使用文件夹名，并限制长度，防止文件名过长
        return folder_name[:30] # 限制为30个字符，避免过长

# --- File Operations ---
def validate_directory(path: Path, logger_obj) -> bool:
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

def copy_file(source_path: Path, destination_path: Path, logger_obj) -> bool:
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

# --- Utility Function to Normalize Drive Letter ---
def normalize_drive_letter(path_str: str) -> str:
    """
    如果路径以驱动器号开头，将其转换为大写。
    例如: c:\\test -> C:\\test
    """
    if sys.platform.startswith('win') and len(path_str) >= 2 and path_str[1] == ':':
        return path_str[0].upper() + path_str[1:]
    return path_str



def read_batch_paths(batch_file_path: Path, logger_obj) -> List[Path]:
    """
    从 batchPath.txt 文件中读取需要扫描的文件夹路径列表。
    Args:
        batch_file_path (Path): batchPath.txt 文件的路径。
        logger_obj (logger): 日志管理器实例。
    Returns:
        List[Path]: 文件夹路径的列表。
    """
    folders = []
    if not batch_file_path.exists():
        logger_obj.error(f"错误: 批量路径文件 '{normalize_drive_letter(str(batch_file_path))}' 不存在。")#error
        print(f"错误: 批量路径文件 '{batch_file_path}' 不存在。")
        return folders
    try:
        with open(batch_file_path, 'r', encoding='utf-8') as f:
            for line in f:
                path_str = line.strip()
                if path_str and not path_str.startswith('#'): # 忽略空行和注释行
                    folder_path = Path(path_str)
                    if validate_directory(folder_path, logger_obj):
                        folders.append(folder_path)
                    else:
                        logger_obj.warning(f"警告: 批量路径文件中的路径无效或不存在，已跳过: {normalize_drive_letter(str(folder_path))}")#warning
        if not folders:
            logger_obj.warning(f"警告: 批量路径文件 '{normalize_drive_letter(str(batch_file_path))}' 中没有找到有效的文件夹路径。")#warning
    except Exception as e:
        logger_obj.critical(f"错误: 读取批量路径文件 '{normalize_drive_letter(str(batch_file_path))}' 失败: {e}")#critical
        print(f"错误: 读取批量路径文件 '{batch_file_path}' 失败。错误: {e}")
    return folders

