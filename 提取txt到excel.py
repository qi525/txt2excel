import os
from openpyxl import Workbook

def write_txt_paths_and_content_to_excel(folder_path, excel_file_path):
    """
    将指定文件夹及其子文件夹中所有 .txt 文件的完整路径写入 Excel 表的 A 列，
    并将其内容写入 B 列。如果内容含有换行符，则删除。

    Args:
        folder_path (str): 包含 .txt 文件的根文件夹路径。
        excel_file_path (str): 输出的 Excel 文件路径（例如: "output.xlsx"）。
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "TXT文件信息"

    row_num = 1

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".txt"):
                full_path = os.path.join(root, file)

                # 将路径写入 A 列
                sheet.cell(row=row_num, column=1, value=full_path)

                # 读取 .txt 文件内容
                txt_content = ""
                try:
                    with open(full_path, 'r', encoding='utf-8') as f:
                        txt_content = f.read()
                except Exception as e:
                    print(f"无法读取文件 {full_path}: {e}")
                    txt_content = "读取失败" # 如果读取失败，给出提示

                # 删除换行符
                # replace('\n', '') 替换 Unix/Linux 风格的换行符
                # replace('\r', '') 替换旧 Mac 风格的换行符
                # replace('\r\n', '') 替换 Windows 风格的换行符
                # 按照这个顺序替换可以避免先替换 \r 为空字符串导致 \r\n 变成 \n
                cleaned_content = txt_content.replace('\r\n', '').replace('\n', '').replace('\r', '')

                # 将处理后的内容写入 B 列
                sheet.cell(row=row_num, column=2, value=cleaned_content)

                row_num += 1

    workbook.save(excel_file_path)
    # print(f"所有 .txt 文件的路径和内容已成功写入到 '{excel_file_path}'")
    print(f"The paths and contents of all .txt files were successfully written to'{excel_file_path}'")



### **使用示例**


if __name__ == "__main__":
    # 请将 'D:/test_folder' 替换为你实际要扫描的文件夹路径
    folder_to_scan = "C:\mobile pic\Pictures"

    # # 定义输出的 Excel 文件名和路径
    # output_excel_file = "txt_files_with_content.xlsx"

    cleaned_folder_name = folder_to_scan.replace(":", " ").replace("\\", " ").replace("/", " ")
    output_excel_file = f"{cleaned_folder_name.strip()}.xlsx" # .strip() to remove leading/trailing spaces

    # --- New code for checking file existence ---
    if os.path.exists(output_excel_file):
        # print(f"提示：文件 '{output_excel_file}' 已存在。")
        print(f"Tip: Files '{output_excel_file}' Already exists.")
    # --- End of new code ---

    write_txt_paths_and_content_to_excel(folder_to_scan, output_excel_file)