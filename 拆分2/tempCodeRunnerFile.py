
# ... (保留原有导入，包括 from loguru import logger 和 from my_logger import setup_logger, get_error_log_file_path) ...

def main():
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    log_output_dir = script_dir / "logs"
    history_folder_path = script_dir / HISTORY_FOLDER_NAME
    output_base_dir = script_