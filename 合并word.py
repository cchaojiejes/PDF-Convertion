from os.path import abspath
from win32com import client
from docx2pdf import convert
from PyPDF2 import PdfFileMerger
import time, os, warnings
#若要重新进行合并，需要删除文件"CombinedWord.docx"；"CombinedWord.pdf"; "AllInOne.pdf"
#CombinedWord.docx：把所有的word识别并保存到这个文件里
#CombinedWord.pdf：把上面这个合并了的word转化成pdf格式
#AllInOne.pdf：整合整个文件夹下所有的word和pdf到了这个文件里

def combine_word(path):
    #合并所有word
    files = list()
    filelist = os.listdir(path)
    for file in filelist:
        if (file.endswith(".doc") or file.endswith(".docx")) and ("~$" not in file):
            filePath = path+file
            files.append(filePath)
    # 启动word应用程序
    word = client.gencache.EnsureDispatch("Word.Application")

    word.Visible = True
    # 新建空白文档
    new_document = word.Documents.Add()

    for fn in files[::-1]:
        print ("正在合并：", fn)
        time.sleep(1)  # 每次间隔1s
        fn = abspath(fn)
        new_document.Application.Selection.Range.InsertFile(fn)
    # 保存最终文件，关闭Word应用程序
    combined_word_path = path+"CombinedWord.docx"
    new_document.SaveAs(combined_word_path)
    print("合并word完成，哦吼！")
    word.Documents.Close()
    word.Quit()

def word_pdf(path):
    # word转pdf
    print("正在转换word->pdf")
    combined_word_path = path + "CombinedWord.docx"
    word2pdf = path + "CombinedWord.pdf"
    convert(combined_word_path, word2pdf)
    print("所有word文件转PDF文件已完成！！！已存至路径：“CombinedWord.pdf”。")


def combine_pdf(path):
    # 合并所有pdf
    print("开始合并pdf")
    warnings.filterwarnings("ignore")
    merger = PdfFileMerger()  # 调用PDF文件合并模块
    filelist = os.listdir(path)  # 读取文件夹所有文件
    for file in filelist:
        if file.endswith(".pdf"):
            print("识别到pdf文件，进行合并：", file)
            time.sleep(1)  # 每次间隔1s
            merger.append(path + file)  # 合并PDF文件

    merger.write("AllInOne.pdf")  # 写入PDF文件
    merger.close()
    print("合并pdf完成，已存至路径：“AllInOne.pdf”。")

if __name__ == "__main__":
    path = "C:/Users/cchaojie/Desktop/New folder1111/"  # 读取word的文件夹
    # combine_word(path)
    # word_pdf(path)
    # combine_pdf(path)