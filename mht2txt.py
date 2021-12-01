import os
import win32com.client

def mht2doc2txt(filepath):  # doc转docx
    filepath_doc = filepath.strip('.mhtml')+'.doc'
    os.rename(filepath,filepath_doc)  # 重命名成 .doc
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(filepath_doc)
    filepath_docx = filepath_doc.strip('.doc')+'.txt'
    doc.SaveAs(filepath_docx, 2)  # 12表示docx格式
    doc.Close()
    word.Quit()


if __name__ == '__main__':
    # 注意：目录的格式必须写成双反斜杠
    filepath = r'E:\测试.mhtml'
    mht2doc2txt(filepath)