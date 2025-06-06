import os
from openpyxl import Workbook

# 设置要读取的文件夹路径
folder_path = r'C:\sd-webui new\outputs\txt2img-images\_nagihikaru_2025-04-26T05'  # 修改为你的txt文件夹路径
output_xlsx = r'C:\Users\SNOW\Desktop\test.xlsx' # 修改为你想保存的xlsx路径

wb = Workbook()
ws = wb.active
ws.append(['文件夹绝对路径', '图片绝对路径', 'TXT绝对路径', 'TXT内容'])

for root, dirs, files in os.walk(folder_path):
    # 只处理当前目录下的图片和txt
    images = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'))]
    txts = [f for f in files if f.lower().endswith('.txt')]
    for img in images:
        img_name, _ = os.path.splitext(img)
        txt_file = img_name + '.txt'
        if txt_file in txts:
            img_path = os.path.join(root, img)
            txt_path = os.path.join(root, txt_file)
            with open(txt_path, 'r', encoding='utf-8') as f:
                for line in f:
                    ws.append([
                        os.path.abspath(root),
                        os.path.abspath(img_path),
                        os.path.abspath(txt_path),
                        line.strip()
                    ])

wb.save(output_xlsx)
print('合并完成，保存为:', output_xlsx)