import os
from docx import Document

docxDir = input("请输入要提取图片的docx文件目录: ")

def extract_pictures_from_file():
    fileTotal = 0
    for root_dir, sub_dir, files in os.walk(docxDir):
        for file in files:
            if file.endswith('.docx'):
                fileTotal += 1
                file_name = os.path.join(root_dir, file)
                try:
                    doc = Document(file_name)
                    outputpath = input("导出图片到文件夹：")
                    foldername = file.replace('.docx','')
                    newpath = os.path.join(outputpath,foldername)
                    if not os.path.exists(newpath):
                        os.makedirs(newpath)
                    for rel in doc.part.rels.values():
                        if "image" in rel.reltype:
                            img_part = rel.target_part
                            img_data = img_part.blob
                            img_path = os.path.join(newpath,
                                                    f"{file.replace('.docx','')}_{rel.rId}.jpg")
                            with open(img_path, 'wb') as img_file:
                                img_file.write(img_data)
                            print(f"Image extracted: {img_path}")
                except Exception as e:
                    print(f"Error processing file: {file_name}")

extract_pictures_from_file()