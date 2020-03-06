# 内置模块
import os
import json

# 第三方模块
from win32com.client import gencache, constants


def word2pdf():
    '''
    (1) 转换当前目录下的 word 为 pdf
    '''
    folder = './output'
    if not os.path.isdir(folder):
        os.mkdir(folder)
    print('---正在转换为 pdf 文件...')
    path = os.getcwd()
    print(path)
    file_list = os.listdir(path)
    word_list = [
        name for name in file_list if name.endswith(('.doc', '.docx'))]
    pdf_list = [name for name in file_list if name.endswith('.pdf')]
    word = gencache.EnsureDispatch('Word.Application')
    for word_name in word_list:
        pdf_name = os.path.splitext(word_name)[0]+'.pdf'
        if pdf_name in pdf_list:
            continue
        word_path = path+'\\'+word_name
        pdf_path = path+'\\'+folder+'/'+pdf_name
        try:
            doc = word.Documents.Open(word_path, ReadOnly=1)
            doc.ExportAsFixedFormat(pdf_path,
                                    constants.wdExportFormatPDF,
                                    Item=constants.wdExportDocumentWithMarkup,
                                    CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
        except Exception as e:
            print('---转换失败，异常是：{}'.format(e))
        finally:
            word.Quit(constants.wdDoNotSaveChanges)
    print('---转换完毕')


def name2json():
    '''
    (2) 读取当前目录下的文件名称，存为 json 文件
    '''
    print('---正在生成 json 文件...')
    print('---生成完毕')


def main():
    '''
    (3) 程序入口
    '''
    print('---请输入以下序号选择功能：')
    print('---1.转换当前目录下的 word 为 pdf')
    print('---2.读取当前目录下文件名称，存为 json 文件')
    print('---3.功能 1 和 2 同时进行')
    option = input()
    if option == '1':
        word2pdf()
    elif option == '2':
        name2json()
    elif option == '3':
        word2pdf()
        name2json()


if __name__ == '__main__':
    main()
