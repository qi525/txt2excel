# main.py
import datetime
import sys
from pathlib import Path
import os
import subprocess

# 导入配置
from config import (
    HISTORY_FOLDER_NAME,
    HISTORY_EXCEL_NAME,
    OUTPUT_FOLDER_NAME,
    CACHE_FOLDER_PATH_STR
)

# 导入工具类和核心逻辑
from utils.file_operations import (
    validate_directory,
    create_directory_if_not_exists,
    copy_file
)
from utils.excel_utils import (
    create_main_workbook,
    setup_excel_sheets,
    apply_hyperlink_style
)
from services.log_manager import LogManager
from services.history_manager import HistoryManager
from core.scanner import scan_files_and_extract_data

# 定义Python运行文件的目录
PYTHON_SCRIPT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))

def read_batch_paths(batch_file_path: Path, log_manager: LogManager) -> list[Path]:
    """
    从 batchPath.txt 文件读取批量处理的文件夹路径。
    """
    paths_to_process = []
    if not batch_file_path.exists():
        log_manager.write_log(f"Error: Batch file not found at {batch_file_path}")
        print(f"错误: 未找到批量处理文件 {batch_file_path}。请确保文件存在。")
        return []

    try:
        with open(batch_file_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                path_str = line.strip()
                if not path_str: # 跳过空行
                    continue
                
                path_obj = Path(path_str)
                if validate_directory(path_obj, log_manager):
                    paths_to_process.append(path_obj)
                else:
                    log_manager.write_log(f"Warning: Invalid path in {batch_file_path} on line {line_num}: '{path_str}'. Skipping.")
                    print(f"警告: 批量文件中第 {line_num} 行路径无效: '{path_str}'。已跳过。")
    except Exception as e:
        log_manager.write_log(f"Error reading batch file {batch_file_path}: {e}")
        print(f"错误: 读取批量文件 {batch_file_path} 时发生错误: {e}")
    return paths_to_process

def main():
    """
    程序主入口，协调文件扫描、数据处理、结果保存和日志记录。
    """
    # 1. 初始化日志管理器 (在任何操作前，确保日志系统可用)
    # 主日志文件直接保存在Python脚本运行目录
    # 如果要避免Permission Denied，可以尝试将主日志也放在用户Documents或其他默认有权限的目录
    # 但根据您之前的代码，似乎是允许在脚本同级目录创建的
    
    # 获取主日志文件的路径
    main_log_file_name = f"main_program_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    main_log_file_path = PYTHON_SCRIPT_DIR / main_log_file_name
    
    # 这里的LogManager初始化，不再指定log_directory，而是直接传递文件路径
    # LogManager的__init__需要调整以支持传入具体文件路径
    # 为了简化，我们让LogManager的__init__只接受目录，然后自己生成文件名
    # 所以，这里还是需要一个log_folder，但主log文件放在这里面，而不是脚本根目录
    
    # 重回之前的LogManager初始化方式，但要确保logs目录权限
    log_folder = PYTHON_SCRIPT_DIR / "logs"
    create_directory_if_not_exists(log_folder, None) # 首次创建日志目录不需要log_manager实例
    
    # 主日志管理器，记录整个程序的运行情况
    main_log_manager = LogManager(log_folder, log_file_name=main_log_file_name) # 传入具体文件名
    main_log_manager.write_log("Program started.")


    # 2. 初始化历史记录管理器
    # 历史记录文件保存在Python脚本运行目录下的“反推历史记录”文件夹
    history_folder = PYTHON_SCRIPT_DIR / HISTORY_FOLDER_NAME
    create_directory_if_not_exists(history_folder, main_log_manager) # 确保历史记录文件夹存在
    history_file_path = history_folder / HISTORY_EXCEL_NAME
    history_manager = HistoryManager(history_file_path, main_log_manager)

    # 3. 获取用户输入
    user_input = input("请输入您要处理的文件夹路径 (输入0进行批量扫描): ")

    folders_to_scan = []
    if user_input == '0':
        batch_file_path = PYTHON_SCRIPT_DIR / "batchPath.txt"
        main_log_manager.write_log(f"Batch scan mode selected. Reading paths from {batch_file_path}")
        print(f"已选择批量扫描模式，将从 {batch_file_path} 读取路径。")
        folders_to_scan = read_batch_paths(batch_file_path, main_log_manager)
        if not folders_to_scan:
            main_log_manager.write_log("No valid paths found in batch file. Exiting.")
            print("批量文件中未找到有效路径，程序将退出。")
            sys.exit(0)
    else:
        input_path_obj = Path(user_input.strip())
        if validate_directory(input_path_obj, main_log_manager):
            folders_to_scan.append(input_path_obj)
        else:
            main_log_manager.write_log(f"Error: Invalid directory path entered by user: '{user_input}'. Exiting.")
            print(f"错误: 您输入的路径 '{user_input}' 不是一个有效的文件夹。程序将退出。")
            sys.exit(1)

    # 4. 循环处理每个要扫描的文件夹
    for folder_path in folders_to_scan:
        print(f"\n开始扫描文件夹: {folder_path}")
        main_log_manager.write_log(f"Starting scan for folder: {folder_path}")

        current_time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 定义输出文件名和路径
        output_file_name = f"scan_results_{current_time_str}.xlsx"
        scan_specific_log_file_name = f"scan_log_{current_time_str}.txt" # 为每次扫描单独的日志文件命名

        # 定义主要输出文件 (Python运行目录下的“反推历史记录”子文件夹)
        main_output_xlsx = history_folder / output_file_name
        
        # 定义当前扫描的日志文件路径 (仍在 logs 文件夹内，但名称不同)
        scan_log_file_path = log_folder / scan_specific_log_file_name

        # 定义目标文件夹的日志和xlsx文件路径 (复制一份到目标文件夹的“反推记录”子文件夹)
        target_record_folder = folder_path / OUTPUT_FOLDER_NAME
        create_directory_if_not_exists(target_record_folder, main_log_manager)
        target_output_xlsx = target_record_folder / output_file_name
        target_log_file = target_record_folder / scan_specific_log_file_name # 目标文件夹的log文件

        # 配置当前扫描的日志输出到新文件
        current_scan_log_manager = LogManager(log_folder, log_file_name=scan_specific_log_file_name) # 为当前扫描创建独立的log文件
        current_scan_log_manager.write_log(f"Scanning started for: {folder_path}")


        # 5. 设置Excel工作簿
        wb, ws_matched, ws_no_txt, ws_tag_frequency = setup_excel_sheets()

        # 6. 扫描文件并提取数据
        try:
            total_scanned, found_txt_count, not_found_txt_count, tag_counts_data = scan_files_and_extract_data(
                folder_path, ws_matched, ws_no_txt, current_scan_log_manager # 传入当前扫描的log_manager
            )
            print("文件扫描完成。")
            current_scan_log_manager.write_log("File scan completed.")
        except Exception as e:
            current_scan_log_manager.write_log(f"Error during file scanning: {e}")
            print(f"错误: 文件扫描过程中发生错误: {e}")
            continue # 跳过当前文件夹，处理下一个

        # 7. 写入Tag词频统计
        sorted_tag_counts = sorted(tag_counts_data.items(), key=lambda item: item[1], reverse=True)
        for tag, count in sorted_tag_counts:
            ws_tag_frequency.append([tag, count])
        current_scan_log_manager.write_log("Tag frequency compiled.")

        # 8. 应用超链接样式
        apply_hyperlink_style(ws_matched, 3) # "文件超链接" 在第3列
        apply_hyperlink_style(ws_no_txt, 3)  # "文件超链接" 在第3列
        current_scan_log_manager.write_log("Hyperlink styles applied.")

        # 9. 保存主输出文件
        try:
            wb.save(str(main_output_xlsx))
            print(f'合并完成，已保存至Python运行目录下的“反推历史记录”文件夹: {main_output_xlsx}')
            current_scan_log_manager.write_log(f"Results saved to Python script history directory: {main_output_xlsx}")
        except Exception as e:
            current_scan_log_manager.write_log(f"Error: Could not save results to Python script history directory {main_output_xlsx}. Error: {e}")
            print(f"错误: 无法保存结果到Python运行目录下的“反推历史记录”文件夹 {main_output_xlsx}。错误: {e}")
            continue # 跳过当前文件夹，处理下一个

        # 10. 复制一份到目标文件夹
        try:
            copy_file(main_output_xlsx, target_output_xlsx, current_scan_log_manager)
            print(f'一份副本已保存至目标文件夹: {target_output_xlsx}')
        except Exception as e:
            current_scan_log_manager.write_log(f"Error: Could not copy XLSX to target folder {target_output_xlsx}. Error: {e}")
            print(f"错误: 无法复制 XLSX 到目标文件夹 {target_output_xlsx}。错误: {e}")

        # 11. 复制log文件到目标文件夹
        # 在复制前确保日志文件已关闭并写入完成
        current_scan_log_manager.close() # 在复制前确保日志文件已关闭并写入完成
        try:
            if scan_log_file_path.exists(): # 只有当日志文件实际存在时才尝试复制
                copy_file(scan_log_file_path, target_log_file, main_log_manager) # 使用主日志管理器记录复制log
                print(f'一份log副本已保存至目标文件夹: {target_log_file}')
            else:
                main_log_manager.write_log(f"Warning: Scan specific log file did not exist to copy: {scan_log_file_path}")
                print(f"警告: 本次扫描的日志文件 {scan_log_file_path} 不存在，未能复制到目标文件夹。")
        except Exception as e:
            main_log_manager.write_log(f"Error: Could not copy log to target folder {target_log_file}. Error: {e}")
            print(f"错误: 无法复制log到目标文件夹 {target_log_file}。错误: {e}")

        # 12. 更新历史记录
        try:
            history_manager.update_history(
                folder_path, total_scanned, found_txt_count, not_found_txt_count,
                main_output_xlsx, scan_log_file_path # 传入的是单次扫描的log文件路径
            )
        except Exception as e:
            main_log_manager.write_log(f"Error updating history for {folder_path}: {e}")
            print(f"错误: 更新历史记录失败 for {folder_path}: {e}")

        # 13. 复制历史记录文件到缓存 (如果需要的话，仅复制一份最新的历史记录到缓存)
        history_cache_file_path = None
        cache_folder = PYTHON_SCRIPT_DIR / CACHE_FOLDER_PATH_STR
        create_directory_if_not_exists(cache_folder, main_log_manager)
        try:
            # 复制最新版本的历史记录文件到缓存，以当前时间戳命名
            if history_file_path.exists():
                cache_history_file_name = f"scan_history_{current_time_str}.xlsx"
                history_cache_file_path = cache_folder / cache_history_file_name
                copy_file(history_file_path, history_cache_file_path, main_log_manager)
                print(f"历史记录文件已复制到缓存: {history_cache_file_path}")
            else:
                main_log_manager.write_log(f"History file {history_file_path} does not exist, cannot copy to cache.")
        except Exception as e:
            main_log_manager.write_log(f"Error copying history file to cache: {e}")
            print(f"错误: 无法复制历史记录文件到缓存: {e}")
            history_cache_file_path = None # 复制失败则不尝试打开

        # 14. 自动运行打开文件
        try:
            # 不再使用等待加载"networkidle"。
            # 注意：这里打开的是当前文件夹的输出文件和日志，以及最新的历史记录缓存文件
            files_to_open = [main_output_xlsx] # 总是尝试打开主输出XLSX
            
            # 只有当日志文件确实被创建了，并且目标存在，才尝试打开
            if scan_log_file_path.exists():
                files_to_open.append(scan_log_file_path)
            else:
                main_log_manager.write_log(f"Warning: Attempted to open non-existent scan log file: {scan_log_file_path}")
                print(f"警告: 尝试打开不存在的扫描日志文件: {scan_log_file_path}")


            if history_cache_file_path and history_cache_file_path.exists():
                files_to_open.append(history_cache_file_path)

            for file_path_to_open in files_to_open:
                if not file_path_to_open.exists():
                    main_log_manager.write_log(f"Attempted to open non-existent file: {file_path_to_open}")
                    print(f"警告: 尝试打开不存在的文件: {file_path_to_open}")
                    continue

                if sys.platform.startswith('win'): # Windows
                    os.startfile(str(file_path_to_open))
                elif sys.platform == 'darwin': # macOS
                    subprocess.Popen(['open', str(file_path_to_open)])
                else: # Linux/Unix
                    subprocess.Popen(['xdg-open', str(file_path_to_open)])
                
                print(f"自动打开: {file_path_to_open}")

        except Exception as e:
            main_log_manager.write_log(f"Error automatically opening files. Error: {e}")
            print(f"无法自动打开文件或缓存历史记录。请手动检查。错误: {e}")

        print(f"文件夹 {folder_path} 扫描及处理结束。")
        main_log_manager.write_log(f"Finished processing folder: {folder_path}")
        
    main_log_manager.write_log("Program finished.")
    main_log_manager.close() # 确保主日志文件也关闭
    print("程序运行结束。")

if __name__ == "__main__":
    main()