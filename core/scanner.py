# core/scanner.py
import os
import sys
from collections import defaultdict
from pathlib import Path
from typing import Tuple, Dict
from openpyxl.worksheet.worksheet import Worksheet

from core.data_processor import detect_types, clean_tags
from services.log_manager import LogManager
from utils.file_operations import get_file_details

def scan_files_and_extract_data(
    base_folder_path: Path,
    ws_matched: Worksheet,
    ws_no_txt: Worksheet,
    log_manager: LogManager
) -> Tuple[int, int, int, Dict[str, int]]:
    """
    扫描指定文件夹下的文件，查找匹配的TXT文件，提取数据并写入Excel。
    """
    total_files_scanned = 0
    found_txt_count = 0
    not_found_txt_count = 0 # <-- 确保这里是 not_found_txt_count
    tag_counts = defaultdict(int)

    for root_str, _, files in os.walk(base_folder_path):
        root = Path(root_str)
        current_txt_files = {
            txt_file.stem.lower(): txt_file
            for txt_file in root.iterdir() if txt_file.suffix.lower() == '.txt'
        }

        for f_name_str in files:
            file_path = root / f_name_str
            
            if file_path.suffix.lower() == '.txt':
                continue

            total_files_scanned += 1

            file_name_without_ext, file_ext = get_file_details(file_path)

            txt_content = ''
            cleaned_data = ''
            prompt_type = ''
            txt_absolute_path = ''
            found_txt = '否'
            cleaned_data_length = 0

            file_abs_path = file_path.resolve() # 获取文件的绝对路径 Path 对象

            link_path_for_excel = str(file_abs_path) 
            
            hyperlink_formula = f'=HYPERLINK("{link_path_for_excel}", "打开文件")'

            if file_name_without_ext in current_txt_files:
                txt_file_path = current_txt_files[file_name_without_ext]
                try:
                    with open(txt_file_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            txt_content = line.strip()
                            cleaned_data, _ = clean_tags(txt_content)
                            cleaned_data_length = len(cleaned_data)
                            prompt_type = detect_types(txt_content)
                            txt_absolute_path = str(txt_file_path.resolve()) # 转为字符串
                            found_txt = '是'
                            found_txt_count += 1

                            for tag in cleaned_data.split(', '):
                                if tag:
                                    tag_counts[tag.strip().lower()] += 1
                            break
                except Exception as e:
                    log_manager.write_log(f"Error reading TXT file {txt_file_path}: {e}")
                    txt_content = f"Error reading TXT: {e}"
                    found_txt = '否 (读取错误)'
                    not_found_txt_count += 1
            else:
                log_manager.write_log(f"No matching TXT file found for: {file_path}")
                not_found_txt_count += 1

            if found_txt == '是':
                ws_matched.append([
                    str(root.resolve()),
                    str(file_abs_path),
                    hyperlink_formula,
                    file_ext,
                    txt_absolute_path,
                    txt_content,
                    cleaned_data,
                    cleaned_data_length,
                    prompt_type,
                    found_txt
                ])
            else:
                ws_no_txt.append([
                    str(root.resolve()),
                    str(file_abs_path),
                    hyperlink_formula,
                    file_ext,
                    found_txt
                ])
    return total_files_scanned, found_txt_count, not_found_txt_count, tag_counts # <-- 修正这里！