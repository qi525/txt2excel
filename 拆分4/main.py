# main.py

import os
import sys
import datetime
import time
from pathlib import Path
from typing import Optional, Dict, Any, List

# 从 loguru 导入 logger，setup_logger 用于配置日志
from loguru import logger
from my_logger import setup_logger

# 从 excel_utilities 导入相关函数和常量
from excel_utilities import FIXED_COLUMN_WIDTH
from excel_utilities import create_empty_workbook, create_sheet_with_headers, set_column_widths, set_hyperlink_and_style, set_fixed_column_widths


from file_system_utils import (
    generate_folder_prefix,
    read_batch_paths,
    normalize_drive_letter,
    create_directory_if_not_exists,
    copy_file
)

# 从重构后的 scanner.py 导入函数
from scanner import scan_files_and_extract_data, ExcelDataWriter

# 导入 HistoryManager 和历史记录相关常量，并导入 _handle_history_caching 函数
from history_execution import HistoryManager, HISTORY_FOLDER_NAME, HISTORY_EXCEL_NAME, _handle_history_caching # 修改导入

# 导入自动打开文件的函数
from file_opener import open_output_files_automatically

# --- Configuration ---
OUTPUT_FOLDER_NAME = "反推记录"
CACHE_FOLDER_NAME = "cache"


# 文件保存重试参数
MAX_SAVE_RETRIES = 5
RETRY_DELAY_SECONDS = 2

# 将 script_dir 和 log_output_folder 移到全局作用域
script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
log_output_folder = script_dir / "logs" # 或者根据你的配置路径

# _handle_history_caching 函数定义已从这里删除，并移动到 history_execution.py 中


def main():
    history_folder_path = script_dir / HISTORY_FOLDER_NAME
    output_base_dir = script_dir / OUTPUT_FOLDER_NAME
    final_history_excel_path = history_folder_path / HISTORY_EXCEL_NAME
    cache_folder_path = script_dir / CACHE_FOLDER_NAME

    # 配置日志系统，并获取错误日志文件路径
    error_warning_log_file_path = setup_logger(log_output_folder)

    logger.info(f"程序启动. 脚本目录: {normalize_drive_letter(str(script_dir))}")
    logger.info(f"日志文件将保存到: {normalize_drive_letter(str(log_output_folder))}")
    logger.info(f"错误和警告日志将写入到独立的日志文件 (仅当出现警告或错误时创建)。")

    if not create_directory_if_not_exists(history_folder_path, logger):
        logger.critical("致命错误: 无法创建历史记录文件夹，程序退出。")
        sys.exit(1)

    if not create_directory_if_not_exists(cache_folder_path, logger):
        logger.critical("致命错误: 无法创建缓存文件夹，程序退出。")
        sys.exit(1)

    if not create_directory_if_not_exists(output_base_dir, logger):
        logger.critical("致命错误: 无法创建输出文件夹 (反推记录)，程序退出。")
        sys.exit(1)

    # --- 开始修改历史管理器初始化和使用方式 ---
    # 定义文件扫描项目的字段结构
    # 这里的顺序决定了Excel中列的顺序
    file_scan_field_definitions: List[Dict[str, Any]] = [
        {"internal_key": "scan_time", "excel_header": "扫描时间", "is_path": False},
        {"internal_key": "folder_path", "excel_header": "文件夹路径", "is_path": True,
         "hyperlink_display_text": "打开文件夹", "hyperlink_not_exist_text": "文件夹不存在"},
        {"internal_key": "total_files", "excel_header": "总文件数", "is_path": False},
        {"internal_key": "found_txt_count", "excel_header": "找到TXT文件数", "is_path": False},
        {"internal_key": "not_found_txt_count", "excel_header": "未找到TXT文件数", "is_path": False},
        {"internal_key": "log_file_abs_path", "excel_header": "Log文件绝对路径", "is_path": True,
         "hyperlink_display_text": "打开Log", "hyperlink_not_exist_text": "Log文件不存在"},
        {"internal_key": "result_xlsx_abs_path", "excel_header": "结果XLSX文件绝对路径", "is_path": True,
         "hyperlink_display_text": "打开结果XLSX", "hyperlink_not_exist_text": "结果XLSX文件不存在"}
    ]

    # 实例化 HistoryManager，传入 field_definitions
    history_manager = HistoryManager(final_history_excel_path, logger, file_scan_field_definitions)
    # --- 结束修改历史管理器初始化和使用方式 ---

    batch_file_path = script_dir / "batchPath.txt"
    folders_to_scan = read_batch_paths(batch_file_path, logger)

    logger.info(f"在 '{normalize_drive_letter(str(batch_file_path))}' 中检测到 {len(folders_to_scan)} 条有效地址。")

    if not folders_to_scan:
        logger.info("没有找到要扫描的文件夹路径，程序终止。")
        logger.close()
        sys.exit(0)

    current_folder_log_sink_id: Optional[int] = None

    final_files_to_open_at_end = []

    for folder_path in folders_to_scan:
        logger.info(f"\n--- 开始处理文件夹: {normalize_drive_letter(str(folder_path))} ---")

        scan_timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_prefix = generate_folder_prefix(folder_path)

        current_scan_log_file = output_base_dir / f"{folder_prefix}_scan_log_{scan_timestamp}.txt"
        current_excel_file = output_base_dir / f"{folder_prefix}_scan_results_{scan_timestamp}.xlsx"

        fallback_excel_file = log_output_folder / f"FALLBACK_{folder_prefix}_scan_results_{scan_timestamp}.xlsx"

        try:
            current_folder_log_sink_id = logger.add(
                str(current_scan_log_file),
                level="INFO",
                rotation="5 MB",
                compression="zip",
                enqueue=True,
                encoding="utf-8",
                format="{time:YYYY-MM-DD HH:mm:ss.SSS} | {level: <8} | {message}"
            )
            logger.info(f"针对当前文件夹的扫描日志将写入: {normalize_drive_letter(str(current_scan_log_file))}")
        except Exception as e:
            logger.error(f"无法为文件夹 '{normalize_drive_letter(str(folder_path))}' 添加扫描日志文件 sink: {e}")
            current_folder_log_sink_id = None

        logger.info(f"开始扫描 {normalize_drive_letter(str(folder_path))}")

        try:
            wb = create_empty_workbook()

            # 定义各个工作表的标题
            matched_headers = ["文件夹路径", "文件绝对路径", "文件链接", "文件扩展名",
                               "TXT文件绝对路径", "TXT文件内容", "清洗后内容", "内容长度",
                               "提示词类型", "找到TXT"]
            unmatched_headers = ["文件夹路径", "文件绝对路径", "文件链接", "文件扩展名", "找到TXT"]
            tag_frequency_headers = ["Tag", "出现次数"]

            # 创建“匹配文件”工作表
            ws_matched = create_sheet_with_headers(wb, "匹配文件", matched_headers, 0)

            # 创建“未匹配文件”工作表
            ws_no_txt = create_sheet_with_headers(wb, "未匹配文件", unmatched_headers, 1)

            # 创建“Tag词频统计”工作表
            ws_tag_frequency = create_sheet_with_headers(wb, "Tag词频统计", tag_frequency_headers, 2)


            # 在调用 scan_files_and_extract_data 之前，创建 ExcelDataWriter 实例
            excel_data_writer = ExcelDataWriter(ws_matched, ws_no_txt, logger)
            total_files, found_txt_count, not_found_txt_count, tag_counts = scan_files_and_extract_data(
                folder_path,
                excel_data_writer,
                logger
            )

            sorted_tags = sorted(tag_counts.items(), key=lambda item: item[1], reverse=True)
            for tag, count in sorted_tags:
                ws_tag_frequency.append([tag, count])

            for worksheet in [ws_matched, ws_no_txt, ws_tag_frequency]:
                set_fixed_column_widths(worksheet, FIXED_COLUMN_WIDTH, logger)

            save_successful = False
            actual_result_file_path = Path("N/A_SAVE_FAILED")

            for attempt in range(MAX_SAVE_RETRIES):
                try:
                    wb.save(str(current_excel_file))
                    logger.info(f"扫描结果已保存到: {normalize_drive_letter(str(current_excel_file))} (尝试 {attempt + 1}/{MAX_SAVE_RETRIES})")
                    actual_result_file_path = current_excel_file
                    save_successful = True
                    break
                except PermissionError as e:
                    logger.warning(f"警告: 无法将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}'，原因: 权限拒绝！请确保该文件未被其他程序（如Excel）打开。尝试 {attempt + 1}/{MAX_SAVE_RETRIES}。错误: {e}")
                    time.sleep(RETRY_DELAY_SECONDS)
                except Exception as e:
                    logger.error(f"错误: 将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}' 失败: {e} (尝试 {attempt + 1}/{MAX_SAVE_RETRIES})")
                    break

            if not save_successful:
                logger.critical(f"严重警告: 经过 {MAX_SAVE_RETRIES} 次尝试后，仍无法将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}'。尝试保存到备用位置。")

                try:
                    wb.save(str(fallback_excel_file))
                    logger.warning(f"成功将扫描结果保存到备用位置: {normalize_drive_letter(str(fallback_excel_file))}")
                    actual_result_file_path = fallback_excel_file
                except Exception as fallback_e:
                    logger.critical(f"致命错误: 尝试将扫描结果保存到备用位置 '{normalize_drive_letter(str(fallback_excel_file))}' 也失败了！错误: {fallback_e}")
                    actual_result_file_path = Path("N/A_SAVE_FAILED")

            # --- 修改 add_history_entry 的调用方式 ---
            new_entry_data: Dict[str, Any] = {
                "scan_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "folder_path": folder_path,
                "total_files": total_files,
                "found_txt_count": found_txt_count,
                "not_found_txt_count": not_found_txt_count,
                "log_file_abs_path": current_scan_log_file,
                "result_xlsx_abs_path": actual_result_file_path
            }
            history_manager.add_history_entry(new_entry_data)
            logger.info(f"本次扫描历史记录已成功添加至内存。")
            # --- 结束修改 add_history_entry 的调用方式 ---

            files_to_open_this_scan = [current_scan_log_file]
            if actual_result_file_path.exists():
                files_to_open_this_scan.append(actual_result_file_path)

            open_output_files_automatically(files_to_open_this_scan, logger)

        except Exception as e:
            logger.error(f"处理文件夹 {normalize_drive_letter(str(folder_path))} 时发生错误: {e}")
        finally:
            if current_folder_log_sink_id is not None:
                logger.remove(current_folder_log_sink_id)
                current_folder_log_sink_id = None
            logger.info(f"--- 完成处理文件夹: {normalize_drive_letter(str(folder_path))} ---\n")

    logger.info(f"所有扫描任务完成，开始将历史记录保存到最终的Excel文件: {normalize_drive_letter(str(final_history_excel_path))}")
    logger.info(f"准备保存 {len(history_manager.history_data)} 条历史记录到Excel。")
    logger.info(f"最初在 '{normalize_drive_letter(str(batch_file_path))}' 中检测到 {len(folders_to_scan)} 条有效地址。")

    save_history_success = history_manager.save_history_to_excel()

    # 调用新封装的缓存处理函数
    _handle_history_caching(
        save_history_success,
        final_history_excel_path,
        cache_folder_path,
        logger,
        final_files_to_open_at_end
    )

    # Add error/warning log file to the list
    if error_warning_log_file_path and error_warning_log_file_path.exists():
        final_files_to_open_at_end.append(error_warning_log_file_path)
    else:
        logger.warning("警告: 错误和警告日志文件不存在或路径无效，无法自动打开。")

    open_output_files_automatically(final_files_to_open_at_end, logger)

    logger.info("所有文件夹处理完毕，程序即将退出。")

if __name__ == "__main__":
    main()