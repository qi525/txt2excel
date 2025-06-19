import datetime
import os
import sys
from pathlib import Path
from typing import List, Dict, Any

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet # 用于类型提示

from loguru import logger
from loguru import logger
from my_logger import normalize_drive_letter 

# --- 主要改动点 START ---
# 从新的 excel_utils.py 导入所需的函数和常量
#from excel_utils import FIXED_COLUMN_WIDTH, set_hyperlink_and_style, set_fixed_column_widths 
# --- 主要改动点 END ---
# --- HistoryManager (Excel Version) ---
