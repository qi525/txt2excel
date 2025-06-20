# excel_utilities.py
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Tuple, Optional
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HYPERLINK_FONT = Font(color="0000FF", underline="single")

# 定义固定列宽（以字符为单位）
FIXED_COLUMN_WIDTH = 20

# --- Excel Utilities ---
def create_main_workbook():
    """
    创建主Excel工作簿，包含“匹配文件”和“未匹配文件”工作表。
    """
    wb = Workbook()
    
    # 移除默认创建的Sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
        
    return wb

def setup_excel_sheets(wb: Workbook) -> Tuple[Worksheet, Worksheet, Worksheet]:
    """
    设置Excel工作表及其标题行。
    """
    # 匹配文件工作表
    ws_matched = wb.create_sheet("匹配文件", 0) # 插入到最前面
    ws_matched.append([
        "文件夹路径", "文件绝对路径", "文件链接", "文件扩展名", "TXT文件绝对路径",
        "TXT文件内容", "清洗后内容", "内容长度", "提示词类型", "找到TXT"
    ])

    # 未匹配文件工作表
    ws_no_txt = wb.create_sheet("未匹配文件", 1) # 插入到第二个
    ws_no_txt.append([
        "文件夹路径", "文件绝对路径", "文件链接", "文件扩展名", "找到TXT"
    ])

    # Tag词频统计工作表
    ws_tag_frequency = wb.create_sheet("Tag词频统计", 2) # 插入到第三个
    ws_tag_frequency.append(["Tag", "出现次数"])

    return ws_matched, ws_no_txt, ws_tag_frequency



# --- 辅助函数 ---
# 将 set_hyperlink_and_style 函数粘贴到这里
def set_hyperlink_and_style(
    cell, 
    location: Optional[str], # location 现在可以是 Optional[str]
    display_text: str, 
    logger_obj, 
    source_description: str = "未知源"
):
    """
    封装设置单元格超链接和样式的逻辑。
    Args:
        cell: openpyxl 单元格对象。
        location (Optional[str]): 超链接指向的实际位置（文件路径或URL）。如果为None或空字符串，则不设置超链接。
        display_text (str): 在单元格中显示的文本。
        logger_obj (logger): 日志管理器实例。
        source_description (str): 描述超链接来源，用于日志记录。
    """
    try:
        cell.value = display_text # 首先设置单元格显示文本
        
        # 只有当 location 不为 None 且不为空时才设置超链接
        if location: # 检查 location 是否有效
            cell.hyperlink = location # 然后设置超链接目标
            cell.font = HYPERLINK_FONT # 最后应用预定义的超链接字体样式
            logger_obj.info(
                f"成功设置超链接和样式 for '{display_text}' (Location: '{location}', Source: {source_description})"
            )
        else:
            # 如果没有 location，确保不设置超链接，并移除可能的超链接样式
            cell.hyperlink = None 
            cell.font = Font(color="000000") # 恢复默认字体颜色，去除下划线
            # 这条日志保留，因为仍然是提示没有设置超链接，但级别可以低一些
            logger_obj.info(f"未为 '{display_text}' (Source: {source_description}) 设置超链接，因为location无效或为空。")

    except Exception as e:
        logger_obj.error(
            f"错误: 无法为单元格设置超链接或样式 for '{display_text}' (Location: '{location}', Source: {source_description}). 错误: {e}"
        )
        # 即使出错，也要确保单元格值被设置，即使没有超链接样式
        cell.value = display_text



# --- NEW FUNCTION: Set Fixed Column Widths for a Worksheet ---
def set_fixed_column_widths(worksheet: Worksheet, width: int, logger_obj):
    """
    为给定工作表的所有列设置固定宽度。
    Args:
        worksheet (Worksheet): openpyxl 工作表对象。
        width (int): 要设置的列宽。
        logger_obj (logger): 日志管理器实例。
    """
    try:
        for col_idx in range(1, worksheet.max_column + 1): # 从1开始遍历所有列
            column_letter = get_column_letter(col_idx)
            worksheet.column_dimensions[column_letter].width = width
        logger_obj.info(f"已为工作表 '{worksheet.title}' 设置所有列宽度为 {width}.") # 替换为 Loguru 的 info 方法，to_file_only 行为 Loguru 默认在 setup 时配置
    except Exception as e:
        #logger_obj.info(f"错误: 无法为工作表 '{worksheet.title}' 设置列宽: {e}")#error
        logger_obj.error(f"错误: 无法为工作表 '{worksheet.title}' 设置列宽: {e}") # 替换为 Loguru 的 error 方法
        print(f"错误: 无法为工作表 '{worksheet.title}' 设置列宽. 错误: {e}")
