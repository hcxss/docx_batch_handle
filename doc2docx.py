# encoding=utf-8
import os

from win32com import client as wc


def doc2docxFun(path):
    w = wc.gencache.EnsureDispatch("Word.Application")
    # doc = w.Documents.Open(r"D:/data/测试.doc")  # 读取指定路径的doc
    # doc.SaveAs2(r"D:/data/测试.docx", 12)
    doc = w.Documents.Open(r"" + path + "")  # 读取指定路径的doc
    doc.SaveAs2(r"" + path + "x", 12)


path = "D:/data/"  # 需处理文件夹路径
files = []
for file in os.listdir(path):
    if file.endswith(".doc"):  # 排除文件夹内的其它干扰文件，只获取word文件
        files.append(path + file)
        doc2docxFun(path + file)
# doc2docxFun(path)
