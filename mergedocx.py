from docx import Document
import os
import msvcrt
from docxcompose.composer import Composer

DEBUG = True


def crawl(path: str):
    master = Document()
    middle_new_docx = Composer(master)

    for root_dirc, sub_dirc, files in os.walk(path):
        if len(files) == 0:
            print("文件夹中无word文档!")
            break
        for file in files:
            if file.endswith("docx") and "result" not in file:
                try:
                    word_document = Document(file)

                    middle_new_docx.append(word_document)

                    print(f"读取文件: {file}")
                except PermissionError:
                    print(f"{file}读取失败!")

    out_path = os.path.join(path, "result.docx")
    middle_new_docx.save(out_path)
    print(f"文件输出到: {out_path}")


if __name__ == "__main__":
    try:
        dirc = os.getcwd()
        crawl(dirc)
    finally:
        os.system('pause')

