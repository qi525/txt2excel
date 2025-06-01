import os
from openpyxl import Workbook

def write_txt_paths_and_content_to_excel(folder_path, excel_file_path):
    """
    将指定文件夹及其子文件夹中所有 .txt 文件的完整路径写入 Excel 表的 A 列，
    将其内容写入 B 列，并将 B 列内容复制到 C 列并删除 C 列中的指定词组。
    如果内容含有换行符，则删除。

    Args:
        folder_path (str): 包含 .txt 文件的根文件夹路径。
        excel_file_path (str): 输出的 Excel 文件路径（例如: "output.xlsx"）。
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "TXT文件信息"

    # Set headers for clarity
    sheet.cell(row=1, column=1, value="文件路径")
    sheet.cell(row=1, column=2, value="原始内容")
    sheet.cell(row=1, column=3, value="清洗后的内容")
    
    row_num = 2  # Start from row 2 because row 1 is now headers

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".txt"):
                full_path = os.path.join(root, file)

                # Write path to column A
                sheet.cell(row=row_num, column=1, value=full_path)

                # Read .txt file content
                txt_content = ""
                try:
                    with open(full_path, 'r', encoding='utf-8') as f:
                        txt_content = f.read()
                except Exception as e:
                    print(f"无法读取文件 {full_path}: {e}")
                    txt_content = "读取失败" # If read fails, provide a hint

                # Remove newlines
                cleaned_content = txt_content.replace('\r\n', '').replace('\n', '').replace('\r', '')

                # Write processed content to column B
                sheet.cell(row=row_num, column=2, value=cleaned_content)

                # --- New logic for column C ---
                # Copy content from B to C
                content_for_c = cleaned_content

                # Define the words to be removed (case-insensitive)
                words_to_remove = ["censoring", "censored", "mosaic_censoring"]

                # Replace each word in the content for column C
                for word in words_to_remove:
                    # Use a regex to find whole words and handle cases like "a_censored_b"
                    # For simplicity, we'll do direct string replacement for now,
                    # assuming tags are comma-separated.
                    # If you need more sophisticated word boundary handling,
                    # consider using the `re` module.
                    content_for_c = content_for_c.replace(word, "")
                
                # Clean up any resulting double commas or leading/trailing commas due to replacements
                # This ensures "tag1,,tag2" becomes "tag1,tag2" and ",tag1" becomes "tag1"
                content_for_c = content_for_c.replace(",,", ",").strip(', ')


                # Write processed content to column C
                sheet.cell(row=row_num, column=3, value=content_for_c)
                # --- End of new logic ---

                row_num += 1

    workbook.save(excel_file_path)
    print(f"The paths and contents of all .txt files were successfully written to '{excel_file_path}'")


# ---

### **使用示例 (Usage Example)**

# ```python
if __name__ == "__main__":
    # 请将 'C:\mobile pic\Pictures' 替换为你实际要扫描的文件夹路径
    folder_to_scan = r"C:\mobile pic\Pictures" # Using raw string to avoid issues with backslashes

    # 定义输出的 Excel 文件名和路径
    # Clean up folder name for safe use in file name
    cleaned_folder_name = folder_to_scan.replace(":", "").replace("\\", " ").replace("/", " ").strip()
    output_excel_file = f"{cleaned_folder_name}_processed_tags.xlsx" # Add a suffix to indicate processing

    # Check if the output file already exists
    if os.path.exists(output_excel_file):
        print(f"Tip: File '{output_excel_file}' already exists. It will be overwritten.")

    # Call the function to write data to Excel
    write_txt_paths_and_content_to_excel(folder_to_scan, output_excel_file)