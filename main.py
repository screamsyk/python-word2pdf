# 内置模块
import os
import configparser
import json

# 第三方模块
from win32com.client import gencache, constants

# 全局变量
g_config = None


def read_config():
    '''
    (1)读取配置文件
    '''
    global g_config
    config = configparser.ConfigParser()
    config.read('./conf.ini', encoding="utf-8-sig")
    g_config = config['config'] or {}


def word2pdf():
    '''
    (1) 转换指定目录下的 word 为 pdf
    '''
    print('---正在转换为 pdf 文件...')
    path = os.getcwd() or g_config['path'] or './'
    # output = g_config['output'] or 'output'
    file_list = os.listdir(path)
    word_list = [
        name for name in file_list if name.endswith(('.doc', '.docx'))]
    pdf_list = [name for name in file_list if name.endswith('.pdf')]
    print(word_list)
    word = gencache.EnsureDispatch('Word.Application')
    for word_name in word_list:
        pdf_name = os.path.splitext(word_name)[0]+'.pdf'
        if pdf_name in pdf_list:
            continue
        word_path = path+'\\'+word_name
        pdf_path = path+'\\'+pdf_name
        print(pdf_path)
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
    (2) 读取指定目录下的文件名称，存为 json 文件
    '''
    print('---正在生成 json 文件...')
    print('---生成完毕')


def main():
    '''
    (3) 程序入口
    '''
    read_config()
    print('---请输入以下序号选择功能（目录默认为程序所在的当前目录，可在 conf.ini 中进行修改）：')
    print('---1.转换指定目录下的 word 为 pdf')
    print('---2.读取指定目录下文件名称，存为 json 文件')
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
