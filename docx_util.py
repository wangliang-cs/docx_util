import os.path

from docx import Document
from docxcompose.composer import Composer

import datetime


# refer to: https://blog.csdn.net/m0_46219714/article/details/127694673
def merge_docs(source_filepath_list, output_filepath):
    page_break_doc = Document()
    page_break_doc.add_page_break()

    assert len(source_filepath_list) > 0
    for idx, source_path in enumerate(source_filepath_list):
        print(f"合并[{idx + 1}/{len(source_filepath_list)}]:  {source_path}")
    output_composer = Composer(Document(source_filepath_list[0]))

    for idx, source_path in enumerate(source_filepath_list):
        if idx == 0:
            continue
        output_composer.append(page_break_doc)
        output_composer.append(Document(source_path))

    output_composer.save(output_filepath)


def load_source_list():
    source_filepath_list = []
    source_dir = ""
    with open("./mergelist.txt", "r", encoding="utf-8") as f:
        for idx, line in enumerate(f.read().splitlines()):
            if idx == 0:
                source_dir = line
            else:
                source_filepath_list.append(f"{source_dir}{line}")
    return source_filepath_list


def generate_output_path():
    time_of_execution = datetime.datetime.now().strftime('%Y年%m月%d日%H时%M分%S秒')
    filepath = f"./{time_of_execution}_合并文件.docx"
    count = 1
    while os.path.exists(filepath):
        filepath = f"./{time_of_execution}_合并文件_{count}.docx"
        count = count + 1
    return filepath


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    source_filepath_list = load_source_list()
    output_filepath = generate_output_path()
    merge_docs(source_filepath_list, output_filepath)
    print("===========================")
    print(f"输出: {output_filepath}")
