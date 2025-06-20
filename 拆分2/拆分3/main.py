# InterrogateText2Xlsx7.0_main.py

import os
import sys
import datetime
import time
from pathlib import Path
from typing import  Optional

# 从 my_logger 导入 setup_logger 和 _error_log_file_path
from loguru import logger
from my_logger import setup_logger, get_error_log_file_path, _error_log_file_path

# 新增：从 excel_utilities 导入 create_main_workbook 和 setup_excel_sheets
from excel_utilities import FIXED_COLUMN_WIDTH
from excel_utilities import create_main_workbook, setup_excel_sheets 


from file_system_utils import (
    generate_folder_prefix,
    read_batch_paths,
    normalize_drive_letter,
    create_directory_if_not_exists, 
    copy_file
)

from scanner import scan_files_and_extract_data # 导入扫描函数

from history_execution import HistoryManager,HISTORY_FOLDER_NAME,HISTORY_EXCEL_NAME # 导入历史记录相关常量

from file_opener import open_output_files_automatically # 导入自动打开文件的函数

# --- Configuration ---
OUTPUT_FOLDER_NAME = "反推记录"
CACHE_FOLDER_NAME = "cache"


# 新增：文件保存重试参数
MAX_SAVE_RETRIES = 5
RETRY_DELAY_SECONDS = 2

# 将 script_dir 和 log_output_dir 移到全局作用域
script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
log_output_dir = script_dir / "logs"

# 初始化 Loguru 日志系统
setup_logger(log_output_dir)
# --- END MODIFICATION 1/X ---

def main():
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    log_output_dir = script_dir / "logs"
    history_folder_path = script_dir / HISTORY_FOLDER_NAME
    output_base_dir = script_dir / OUTPUT_FOLDER_NAME
    final_history_excel_path = history_folder_path / HISTORY_EXCEL_NAME
    cache_folder_path = script_dir / CACHE_FOLDER_NAME

    # 新增：用于存储错误和警告日志文件路径
    error_warning_log_file_path = get_error_log_file_path()

    logger.info(f"程序启动. 脚本目录: {normalize_drive_letter(str(script_dir))}")
    logger.info(f"日志文件将保存到: {normalize_drive_letter(str(log_output_dir))}")
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

    history_manager = HistoryManager(final_history_excel_path, logger)

    batch_file_path = script_dir / "batchPath.txt"
    folders_to_scan = read_batch_paths(batch_file_path, logger)

    logger.info(f"在 '{normalize_drive_letter(str(batch_file_path))}' 中检测到 {len(folders_to_scan)} 条有效地址。")

    if not folders_to_scan:
        logger.info("没有找到要扫描的文件夹路径，程序终止。")
        print("没有找到要扫描的文件夹路径，程序终止。")
        logger.close()
        sys.exit(0)

    current_folder_log_sink_id: Optional[int] = None

    # --- START MODIFICATION ---
    # Initialize final_files_to_open_at_end here, at the top level of main()
    final_files_to_open_at_end = []
    # --- END MODIFICATION ---

    for folder_path in folders_to_scan:
        # ... (rest of the loop content, no changes needed here) ...
        logger.info(f"\n--- 开始处理文件夹: {normalize_drive_letter(str(folder_path))} ---")
        print(f"\n--- 开始处理文件夹: {folder_path} ---")

        scan_timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_prefix = generate_folder_prefix(folder_path)

        current_scan_log_file = output_base_dir / f"{folder_prefix}_scan_log_{scan_timestamp}.txt"
        current_excel_file = output_base_dir / f"{folder_prefix}_scan_results_{scan_timestamp}.xlsx"

        fallback_excel_file = log_output_dir / f"FALLBACK_{folder_prefix}_scan_results_{scan_timestamp}.xlsx"

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
            wb = create_main_workbook()
            ws_matched, ws_no_txt, ws_tag_frequency = setup_excel_sheets(wb)

            total_files, found_txt_count, not_found_txt_count, tag_counts = scan_files_and_extract_data(
                folder_path, ws_matched, ws_no_txt, logger
            )

            sorted_tags = sorted(tag_counts.items(), key=lambda item: item[1], reverse=True)
            for tag, count in sorted_tags:
                ws_tag_frequency.append([tag, count])

            for worksheet in [ws_matched, ws_no_txt, ws_tag_frequency]:
                from excel_utilities import set_fixed_column_widths
                set_fixed_column_widths(worksheet, FIXED_COLUMN_WIDTH, logger)

            save_successful = False
            actual_result_file_path = Path("N/A_SAVE_FAILED")

            for attempt in range(MAX_SAVE_RETRIES):
                try:
                    wb.save(str(current_excel_file))
                    logger.info(f"扫描结果已保存到: {normalize_drive_letter(str(current_excel_file))} (尝试 {attempt + 1}/{MAX_SAVE_RETRIES})")
                    print(f"扫描结果已保存到: {current_excel_file}")
                    actual_result_file_path = current_excel_file
                    save_successful = True
                    break
                except PermissionError as e:
                    logger.warning(
                        f"警告: 无法将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}'，原因: 权限拒绝！请确保该文件未被其他程序（如Excel）打开。尝试 {attempt + 1}/{MAX_SAVE_RETRIES}。错误: {e}"
                    )
                    print(f"警告: 无法保存结果到 {current_excel_file}！原因: 权限拒绝。尝试 {attempt + 1}/{MAX_SAVE_RETRIES}。等待 {RETRY_DELAY_SECONDS} 秒后重试...")
                    time.sleep(RETRY_DELAY_SECONDS)
                except Exception as e:
                    logger.error(
                        f"错误: 将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}' 失败: {e} (尝试 {attempt + 1}/{MAX_SAVE_RETRIES})",
                        level="ERROR"
                    )
                    print(f"错误: 无法保存结果到 {current_excel_file}. 错误: {e}")
                    break

            if not save_successful:
                logger.critical(
                    f"严重警告: 经过 {MAX_SAVE_RETRIES} 次尝试后，仍无法将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}'。尝试保存到备用位置。"
                )
                print(f"严重警告: 经过 {MAX_SAVE_RETRIES} 次尝试后，仍无法将扫描结果保存到 {current_excel_file}。尝试保存到备用位置...")
                try:
                    wb.save(str(fallback_excel_file))
                    logger.warning(f"成功将扫描结果保存到备用位置: {normalize_drive_letter(str(fallback_excel_file))}")
                    print(f"成功将扫描结果保存到备用位置: {fallback_excel_file}")
                    actual_result_file_path = fallback_excel_file
                except Exception as fallback_e:
                    logger.critical(
                        f"致命错误: 尝试将扫描结果保存到备用位置 '{normalize_drive_letter(str(fallback_excel_file))}' 也失败了！错误: {fallback_e}"
                    )
                    print(f"致命错误: 无法保存结果到任何位置！错误: {fallback_e}")
                    actual_result_file_path = Path("N/A_SAVE_FAILED")

            history_manager.add_history_entry(
                folder_path, total_files, found_txt_count, not_found_txt_count, actual_result_file_path, current_scan_log_file
            )
            logger.info(f"本次扫描历史记录已成功添加至内存。")

            files_to_open_this_scan = [current_scan_log_file]
            if actual_result_file_path.exists():
                files_to_open_this_scan.append(actual_result_file_path)

            open_output_files_automatically(files_to_open_this_scan, logger)

        except Exception as e:
            logger.error(f"处理文件夹 {normalize_drive_letter(str(folder_path))} 时发生错误: {e}")
            print(f"处理文件夹 {folder_path} 时发生错误。错误: {e}")
        finally:
            if current_folder_log_sink_id is not None:
                logger.remove(current_folder_log_sink_id)
                current_folder_log_sink_id = None
            logger.info(f"--- 完成处理文件夹: {normalize_drive_letter(str(folder_path))} ---\n")
            print(f"--- 完成处理文件夹: {folder_path} ---\n")

    logger.info(f"所有扫描任务完成，开始将历史记录保存到最终的Excel文件: {normalize_drive_letter(str(final_history_excel_path))}")
    logger.info(f"准备保存 {len(history_manager.history_data)} 条历史记录到Excel。")
    logger.info(f"最初在 '{normalize_drive_letter(str(batch_file_path))}' 中检测到 {len(folders_to_scan)} 条有效地址。")

    save_history_success = history_manager.save_history_to_excel()

    # The list is now initialized at the beginning of main()
    # final_files_to_open_at_end = [] # This line is moved/removed

    if save_history_success:
        logger.info("历史记录已成功保存到Excel。")
        if create_directory_if_not_exists(cache_folder_path, logger):
            current_timestamp_for_cache = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            cached_history_file_name = f"scan_history_cached_{current_timestamp_for_cache}.xlsx"
            cached_history_file_path = cache_folder_path / cached_history_file_name

            logger.info(f"开始复制历史记录到缓存文件夹: {normalize_drive_letter(str(cached_history_file_path))}")
            copy_cache_success = copy_file(final_history_excel_path, cached_history_file_path, logger)

            if copy_cache_success:
                logger.info("历史记录已成功复制到缓存文件夹。")
                final_files_to_open_at_end.append(cached_history_file_path)
            else:
                logger.error("历史记录复制到缓存文件夹失败。")
        else:
            logger.error(f"无法创建缓存文件夹: {normalize_drive_letter(str(cache_folder_path))}，将无法复制历史记录。")
    else:
        logger.error("历史记录保存到Excel失败，将不会自动打开历史Excel文件。")

    # Add error/warning log file to the list
    if error_warning_log_file_path and error_warning_log_file_path.exists():
        final_files_to_open_at_end.append(error_warning_log_file_path)
    else:
        logger.warning("警告: 错误和警告日志文件不存在或路径无效，无法自动打开。")

    open_output_files_automatically(final_files_to_open_at_end, logger)

    logger.info("所有文件夹处理完毕，程序即将退出。")
    print("所有文件夹处理完毕，程序即将退出。")
    #logger.close()

if __name__ == "__main__":
    main()
