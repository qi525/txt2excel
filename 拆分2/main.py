# InterrogateText2Xlsx7.0_main.py
import os
import sys
import datetime
import subprocess
import re # 导入re模块用于正则表达式
import time # 导入time模块用于延迟重试
from pathlib import Path
from typing import List

# 从utils模块导入所有必要的函数和类
from utils import (
    normalize_drive_letter, 
    generate_folder_prefix, 
    LogManager, 
    HistoryManager, 
    create_directory_if_not_exists, 
    copy_file,
    create_main_workbook, 
    setup_excel_sheets, 
    scan_files_and_extract_data, 
    read_batch_paths,
    HISTORY_FOLDER_NAME,
    HISTORY_EXCEL_NAME,
    OUTPUT_FOLDER_NAME,
    CACHE_FOLDER_NAME,
    MAX_SAVE_RETRIES,
    RETRY_DELAY_SECONDS,
    FIXED_COLUMN_WIDTH # 新增：导入 FIXED_COLUMN_WIDTH
)

def open_output_files_automatically(files_to_open: List[Path], main_log_manager: LogManager):
    """
    自动打开指定的文件。
    Args:
        files_to_open (List[Path]): 要打开的文件路径列表。
        main_log_manager (LogManager): 主日志管理器实例。
    """
    try:
        # 定义一个正则表达式来匹配文件路径中的日期时间戳 (YYYYMMDD_HHMMSS)
        # 允许文件名在时间戳前后有其他字符，例如 scan_results_20240101_123045.xlsx 或 scan_history_backup_20240101_123045.xlsx
        timestamp_pattern = re.compile(r'\d{8}_\d{6}')

        for file_path_to_open in files_to_open:
            if not file_path_to_open.exists():
                main_log_manager.write_log(f"警告: 尝试打开不存在的文件: {normalize_drive_letter(str(file_path_to_open))}", level="WARNING")
                continue

            # 检查文件名是否包含时间戳，或者是否为明确允许打开的临时文件（例如本次扫描结果）
            # 我们只允许打开包含时间戳的文件作为“快照”，不打开作为“数据库”的源文件
            # 这里的逻辑是：如果文件名包含时间戳，或者文件名是本次扫描生成的xlsx或log文件，则允许打开
            file_name = file_path_to_open.name
            
            is_allowed_to_open = False
            
            # 检查是否包含时间戳（例如 cache 中的历史文件，或者备份的历史文件）
            if timestamp_pattern.search(file_name):
                is_allowed_to_open = True
            
            # 检查是否是本次扫描生成的 Excel 或 Log 文件（这些文件本身就包含时间戳和/或文件夹前缀）
            # 新的命名约定：[folder_prefix]_scan_results_YYYYMMDD_HHMMSS.xlsx
            #             [folder_prefix]_scan_log_YYYYMMDD_HHMMSS.txt
            if re.match(r'.*_scan_results_\d{8}_\d{6}\.xlsx', file_name):
                is_allowed_to_open = True
            elif re.match(r'.*_scan_log_\d{8}_\d{6}\.txt', file_name):
                is_allowed_to_open = True
            elif file_name.startswith("error_warning_log_") and file_name.endswith(".txt"): # 允许打开错误日志
                is_allowed_to_open = True
            # [RESTORE] 恢复对以“scan_history_cached_”开头的文件的特殊允许，因为现在需要缓存备份
            elif file_name.startswith("scan_history_cached_") and file_name.endswith(".xlsx"):
                is_allowed_to_open = True
            # [NEW] 允许打开因权限问题而生成的备用结果文件
            elif file_name.startswith("FALLBACK_") and file_name.endswith(".xlsx"): # 修改为更通用的匹配FALLBACK_
                is_allowed_to_open = True


            if not is_allowed_to_open:
                main_log_manager.write_log(f"拒绝自动打开没有时间戳或不符合命名约定的源文件: {normalize_drive_letter(str(file_path_to_open))}", level="WARNING")
                print(f"拒绝自动打开没有时间戳或不符合命名约定的源文件: {file_path_to_open}")
                continue


            if sys.platform.startswith('win'): 
                subprocess.Popen(f'start "" "{file_path_to_open}"', shell=True) 
            elif sys.platform == 'darwin': 
                subprocess.Popen(['open', str(file_path_to_open)])
            else: 
                subprocess.Popen(['xdg-open', str(file_path_to_open)])
            
            print(f"自动打开: {file_path_to_open}")
            main_log_manager.write_log(f"自动打开: {normalize_drive_letter(str(file_path_to_open))}", level="INFO")

    except Exception as e:
        main_log_manager.write_log(f"错误: 自动打开文件失败. 错误: {e}", level="ERROR")
        print(f"错误: 无法自动打开文件。请手动检查。错误: {e}")

def main():
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    log_output_dir = script_dir / "logs"
    history_folder_path = script_dir / HISTORY_FOLDER_NAME # 这里的 HISTORY_FOLDER_NAME 已经改名为 "运行历史记录"
    output_base_dir = script_dir / OUTPUT_FOLDER_NAME
    final_history_excel_path = history_folder_path / HISTORY_EXCEL_NAME 
    cache_folder_path = script_dir / CACHE_FOLDER_NAME

    # 新增：用于记录WARNING和ERROR的独立日志管理器
    error_warning_log_dir = script_dir / "logs"
    error_warning_log_file_name = f"error_warning_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    # 修改：新增 is_error_log_manager=True 参数
    error_warning_log_manager = LogManager(error_warning_log_dir, log_file_name=error_warning_log_file_name, is_error_log_manager=True)
    # 移除这里对 error_warning_log_manager 的初始 INFO 写入
    # error_warning_log_manager.write_log("错误和警告日志文件已创建。", level="INFO", to_error_log=False) 

    # 修改：主日志管理器现在可以将错误日志转发给 error_warning_log_manager
    main_log_manager = LogManager(log_output_dir, error_log_manager=error_warning_log_manager)
    main_log_manager.write_log(f"程序启动. 脚本目录: {normalize_drive_letter(str(script_dir))}", level="INFO")
    main_log_manager.write_log(f"错误和警告日志将写入到: {normalize_drive_letter(str(error_warning_log_manager.log_file_path))} (仅当出现警告或错误时创建)", level="INFO") # 修改提示信息
    
    if not create_directory_if_not_exists(history_folder_path, main_log_manager):
        main_log_manager.write_log("致命错误: 无法创建历史记录文件夹，程序退出。", level="CRITICAL")
        error_warning_log_manager.close() # 确保关闭错误日志
        sys.exit(1)

    # [RESTORE] 恢复缓存文件夹的创建逻辑
    if not create_directory_if_not_exists(cache_folder_path, main_log_manager):
        main_log_manager.write_log("致命错误: 无法创建缓存文件夹，程序退出。", level="CRITICAL")
        error_warning_log_manager.close() # 确保关闭错误日志
        sys.exit(1)
    
    # 新增：确保反推记录文件夹存在
    if not create_directory_if_not_exists(output_base_dir, main_log_manager):
        main_log_manager.write_log("致命错误: 无法创建输出文件夹 (反推记录)，程序退出。", level="CRITICAL")
        error_warning_log_manager.close()
        sys.exit(1)


    # 初始化历史管理器 (现在是Excel版本)
    history_manager = HistoryManager(final_history_excel_path, main_log_manager)

    batch_file_path = script_dir / "batchPath.txt"
    # 调用read_batch_paths函数
    folders_to_scan = read_batch_paths(batch_file_path, main_log_manager)
    
    # 新增的日志提示：显示从batchPath.txt读取到的有效路径数量
    main_log_manager.write_log(f"在 '{normalize_drive_letter(str(batch_file_path))}' 中检测到 {len(folders_to_scan)} 条有效地址。", level="INFO")

    if not folders_to_scan:
        main_log_manager.write_log("没有找到要扫描的文件夹路径，程序终止。", level="INFO")
        print("没有找到要扫描的文件夹路径，程序终止。")
        main_log_manager.close()
        error_warning_log_manager.close() # 确保关闭错误日志
        sys.exit(0)

    for folder_path in folders_to_scan:
        main_log_manager.write_log(f"\n--- 开始处理文件夹: {normalize_drive_letter(str(folder_path))} ---", level="INFO")
        print(f"\n--- 开始处理文件夹: {folder_path} ---")

        scan_timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_prefix = generate_folder_prefix(folder_path) # [NEW] 生成文件夹前缀
        
        # 移除了 current_output_folder 的创建
        # [MODIFIED] 文件名中包含文件夹前缀
        current_scan_log_file = output_base_dir / f"{folder_prefix}_scan_log_{scan_timestamp}.txt" # 直接在 output_base_dir 下
        current_excel_file = output_base_dir / f"{folder_prefix}_scan_results_{scan_timestamp}.xlsx" # 直接在 output_base_dir 下
        
        # [NEW] 定义备用 Excel 文件路径
        fallback_excel_file = log_output_dir / f"FALLBACK_{folder_prefix}_scan_results_{scan_timestamp}.xlsx" # [MODIFIED] 备用文件名也包含文件夹前缀

        # 确保每个扫描日志使用独立的LogManager实例，现在log_directory直接指向 output_base_dir
        # [MODIFIED] log_file_name 包含文件夹前缀
        scan_log_manager = LogManager(output_base_dir, log_file_name=f"{folder_prefix}_scan_log_{scan_timestamp}.txt", error_log_manager=error_warning_log_manager)
        scan_log_manager.write_log(f"开始扫描 {normalize_drive_letter(str(folder_path))}", level="INFO")

        try:
            wb = create_main_workbook()
            ws_matched, ws_no_txt, ws_tag_frequency = setup_excel_sheets(wb)

            total_files, found_txt_count, not_found_txt_count, tag_counts = scan_files_and_extract_data( # 修正这里的变量名
                folder_path, ws_matched, ws_no_txt, scan_log_manager
            )

            sorted_tags = sorted(tag_counts.items(), key=lambda item: item[1], reverse=True)
            for tag, count in sorted_tags:
                ws_tag_frequency.append([tag, count])

            for worksheet in [ws_matched, ws_no_txt, ws_tag_frequency]:
                # 在这里调用utils中的设置列宽函数
                from utils import set_fixed_column_widths # 确保函数被导入
                set_fixed_column_widths(worksheet, FIXED_COLUMN_WIDTH, scan_log_manager)
            
            # --- 主要修改点：尝试保存Excel文件，如果权限拒绝则保存到logs目录 ---
            save_successful = False
            actual_result_file_path = Path("N/A_SAVE_FAILED") # 默认标记为保存失败

            for attempt in range(MAX_SAVE_RETRIES):
                try:
                    wb.save(str(current_excel_file))
                    scan_log_manager.write_log(f"扫描结果已保存到: {normalize_drive_letter(str(current_excel_file))} (尝试 {attempt + 1}/{MAX_SAVE_RETRIES})", level="INFO")
                    print(f"扫描结果已保存到: {current_excel_file}")
                    actual_result_file_path = current_excel_file
                    save_successful = True
                    break # 成功保存，跳出重试循环
                except PermissionError as e:
                    scan_log_manager.write_log(
                        f"警告: 无法将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}'，原因: 权限拒绝！请确保该文件未被其他程序（如Excel）打开。尝试 {attempt + 1}/{MAX_SAVE_RETRIES}。错误: {e}", 
                        level="WARNING"
                    )
                    print(f"警告: 无法保存结果到 {current_excel_file}！原因: 权限拒绝。尝试 {attempt + 1}/{MAX_SAVE_RETRIES}。等待 {RETRY_DELAY_SECONDS} 秒后重试...")
                    time.sleep(RETRY_DELAY_SECONDS) # 等待一段时间后重试
                except Exception as e: # 捕获其他保存错误
                    scan_log_manager.write_log(
                        f"错误: 将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}' 失败: {e} (尝试 {attempt + 1}/{MAX_SAVE_RETRIES})", 
                        level="ERROR"
                    )
                    print(f"错误: 无法保存结果到 {current_excel_file}. 错误: {e}")
                    # 对于非权限错误，可能没有必要重试，直接跳出并转到备用
                    break 
            
            if not save_successful: # 如果所有重试都失败了
                scan_log_manager.write_log(
                    f"严重警告: 经过 {MAX_SAVE_RETRIES} 次尝试后，仍无法将扫描结果保存到 '{normalize_drive_letter(str(current_excel_file))}'。尝试保存到备用位置。", 
                    level="CRITICAL"
                )
                print(f"严重警告: 经过 {MAX_SAVE_RETRIES} 次尝试后，仍无法将扫描结果保存到 {current_excel_file}。尝试保存到备用位置...")
                try:
                    wb.save(str(fallback_excel_file))
                    scan_log_manager.write_log(f"成功将扫描结果保存到备用位置: {normalize_drive_letter(str(fallback_excel_file))}", level="WARNING")
                    print(f"成功将扫描结果保存到备用位置: {fallback_excel_file}")
                    actual_result_file_path = fallback_excel_file # 更新实际保存路径
                except Exception as fallback_e:
                    scan_log_manager.write_log(
                        f"致命错误: 尝试将扫描结果保存到备用位置 '{normalize_drive_letter(str(fallback_excel_file))}' 也失败了！错误: {fallback_e}", 
                        level="CRITICAL"
                    )
                    print(f"致命错误: 无法保存结果到任何位置！错误: {fallback_e}")
                    actual_result_file_path = Path("N/A_SAVE_FAILED") # 标记保存失败
            # --- 修改结束 ---

            # 将本次扫描结果添加到内存中的历史记录
            # 确保传递的是实际保存成功的路径，如果失败则传入None或标记
            history_manager.add_history_entry(
                folder_path, total_files, found_txt_count, not_found_txt_count, actual_result_file_path, current_scan_log_file
            )
            main_log_manager.write_log(f"本次扫描历史记录已成功添加至内存。", level="INFO")

            # 自动打开本次扫描的Excel、Log文件 
            files_to_open_this_scan = [current_scan_log_file]
            if actual_result_file_path.exists(): # 只有当实际结果文件存在时才尝试打开
                files_to_open_this_scan.append(actual_result_file_path)
            
            open_output_files_automatically(files_to_open_this_scan, main_log_manager)

        except Exception as e:
            main_log_manager.write_log(f"处理文件夹 {normalize_drive_letter(str(folder_path))} 时发生错误: {e}", level="ERROR")
            print(f"处理文件夹 {folder_path} 时发生错误。错误: {e}")
        finally:
            scan_log_manager.close() 
            main_log_manager.write_log(f"--- 完成处理文件夹: {normalize_drive_letter(str(folder_path))} ---\n", level="INFO")
            print(f"--- 完成处理文件夹: {folder_path} ---\n")

    # 所有文件夹处理完毕后，将内存中的历史记录保存到最终的Excel文件
    main_log_manager.write_log(f"所有扫描任务完成，开始将历史记录保存到最终的Excel文件: {normalize_drive_letter(str(final_history_excel_path))}", level="INFO")
    main_log_manager.write_log(f"准备保存 {len(history_manager.history_data)} 条历史记录到Excel。", level="INFO") # 确认即将保存的数量
    main_log_manager.write_log(f"最初在 '{normalize_drive_letter(str(batch_file_path))}' 中检测到 {len(folders_to_scan)} 条有效地址。", level="INFO") # 再次提示最初检测到的有效地址数量

    save_history_success = history_manager.save_history_to_excel()

    # 最终需要自动打开的文件列表
    final_files_to_open_at_end = []

    if save_history_success:
        main_log_manager.write_log("历史记录已成功保存到Excel。", level="INFO")
        
        # [RESTORE] 恢复将历史记录复制到缓存文件夹的逻辑
        if create_directory_if_not_exists(cache_folder_path, main_log_manager):
            current_timestamp_for_cache = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            cached_history_file_name = f"scan_history_cached_{current_timestamp_for_cache}.xlsx" # 区分命名
            cached_history_file_path = cache_folder_path / cached_history_file_name

            main_log_manager.write_log(f"开始复制历史记录到缓存文件夹: {normalize_drive_letter(str(cached_history_file_path))}", level="INFO")
            copy_cache_success = copy_file(final_history_excel_path, cached_history_file_path, main_log_manager)
            
            if copy_cache_success:
                main_log_manager.write_log("历史记录已成功复制到缓存文件夹。", level="INFO")
                final_files_to_open_at_end.append(cached_history_file_path) # 如果复制成功，则打开缓存文件
            else:
                main_log_manager.write_log("历史记录复制到缓存文件夹失败。", level="ERROR")
        else:
            main_log_manager.write_log(f"无法创建缓存文件夹: {normalize_drive_letter(str(cache_folder_path))}，将无法复制历史记录。", level="ERROR")
            
    else:
        main_log_manager.write_log("历史记录保存到Excel失败，将不会自动打开历史Excel文件。", level="ERROR")

    # 将错误和警告日志文件添加到需要自动打开的列表
    if error_warning_log_manager.log_file_path and error_warning_log_manager.log_file_path.exists():
        final_files_to_open_at_end.append(error_warning_log_manager.log_file_path)
    else:
        main_log_manager.write_log("警告: 错误和警告日志文件不存在或路径无效，无法自动打开。", level="WARNING")

    # 最终的自动打开操作
    open_output_files_automatically(final_files_to_open_at_end, main_log_manager)

    main_log_manager.write_log("所有文件夹处理完毕，程序即将退出。", level="INFO")
    print("所有文件夹处理完毕，程序即将退出。")
    main_log_manager.close() # 关闭主日志文件句柄
    error_warning_log_manager.close() # 关闭错误日志文件句柄

if __name__ == "__main__":
    main()