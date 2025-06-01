import os
from openpyxl import Workbook

def write_txt_paths_and_content_to_excel(folder_path, excel_file_path):
    """
    将指定文件夹及其子文件夹中所有 .txt 文件的完整路径写入 Excel 表的 A 列，
    将其内容写入 B 列，并将 B 列内容复制到 C 列，根据关键词删除 C 列中的指定词组。
    如果内容含有换行符，则删除。

    Args:
        folder_path (str): 包含 .txt 文件的根文件夹路径。
        excel_file_path (str): 输出的 Excel 文件路径（例如: "output.xlsx"）。
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "TXT文件信息"

    # 设置表头，让Excel内容更清晰
    sheet.cell(row=1, column=1, value="文件路径")
    sheet.cell(row=1, column=2, value="原始内容")
    # 错误发生在这里，已更正为 sheet.cell
    sheet.cell(row=1, column=3, value="清洗后的内容") 
    
    row_num = 2  # 从第二行开始写入数据，因为第一行是表头

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".txt"):
                full_path = os.path.join(root, file)

                # 将文件路径写入 A 列
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
                cleaned_content = txt_content.replace('\r\n', '').replace('\n', '').replace('\r', '')

                # 将处理后的内容写入 B 列
                sheet.cell(row=row_num, column=2, value=cleaned_content)

                # --- C 列的新处理逻辑 ---
                # 复制 B 列内容到 C 列进行处理
                content_for_c = cleaned_content

                # 定义需要检查的关键词。任何包含这些关键词的标签都将被移除。
                # 注意：这里不包含 'uncensored'。
                keywords_to_remove = ["censor", "mosaic"] 
                
                # 定义例外词，即使包含关键词也不移除
                exception_words = ["uncensored"]

                # 将B列内容按逗号分割成单个标签，并去除空格
                tags = [tag.strip() for tag in content_for_c.split(',')]
                
                # 过滤掉不需要的标签
                filtered_tags = []
                for tag in tags:
                    is_unwanted = False
                    # 检查是否是例外词
                    if tag in exception_words:
                        filtered_tags.append(tag) # 如果是例外词，直接保留
                        continue # 跳过后续的移除检查
                    
                    # 检查标签是否包含任何需要移除的关键词
                    for keyword in keywords_to_remove:
                        # 使用 `in` 运算符进行子字符串匹配
                        if keyword in tag: 
                            is_unwanted = True
                            break # 找到一个匹配项就跳出内部循环
                    
                    # 如果标签不是不需要的（即不含关键词），并且不为空，则保留
                    if not is_unwanted and tag: 
                        filtered_tags.append(tag)
                
                # 重新将过滤后的标签用逗号连接起来
                content_for_c = ','.join(filtered_tags)
                
                # --- C 列处理逻辑结束 ---

                # 将处理后的内容写入 C 列
                sheet.cell(row=row_num, column=3, value=content_for_c)

                row_num += 1

    workbook.save(excel_file_path)
    print(f"所有 .txt 文件的路径和内容已成功写入到 '{excel_file_path}'") # 修正了变量名



### **使用示例**


if __name__ == "__main__":
    # 请将 'C:\mobile pic\Pictures' 替换为你实际要扫描的文件夹路径
    folder_to_scan = r"C:\mobile pic\Pictures" # 使用原始字符串（r""）来避免反斜杠的转义问题

    # 定义输出的 Excel 文件名和路径
    # 清理文件夹名称，使其适合作为文件名
    cleaned_folder_name = folder_to_scan.replace(":", "").replace("\\", " ").replace("/", " ").strip()
    output_excel_file = f"{cleaned_folder_name}_processed_tags_v2.xlsx" # 添加新后缀以区分版本

    # 检查输出文件是否已存在
    if os.path.exists(output_excel_file):
        print(f"提示：文件 '{output_excel_file}' 已存在。它将被覆盖。")

    # 调用函数将数据写入 Excel
    write_txt_paths_and_content_to_excel(folder_to_scan, output_excel_file)