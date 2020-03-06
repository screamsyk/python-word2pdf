# 内置模块
import os
import json

# 第三方模块
from win32com.client import gencache, constants

# 全局变量
g_path = ''
g_output_path = ''


def word2pdf():
    '''
    (1) 转换当前目录下的 word 为 pdf
    '''
    if not os.path.isdir(g_output_path):
        os.mkdir(g_output_path)  # 创建目录
    names = os.listdir(g_path)
    word_list = [name for name in names if name.endswith(('.doc', '.docx'))]
    pdf_list = [name for name in names if name.endswith('.pdf')]
    word = gencache.EnsureDispatch('Word.Application')  # 调用 word 的 API
    for word_name in word_list:
        print('---开始转换：', word_name)
        pdf_name = os.path.splitext(word_name)[0]+'.pdf'
        if pdf_name in pdf_list:
            continue
        word_path = g_path+'/' + word_name
        pdf_path = g_output_path+'/'+pdf_name
        try:
            doc = word.Documents.Open(word_path, ReadOnly=1)
            doc.ExportAsFixedFormat(pdf_path,
                                    constants.wdExportFormatPDF,
                                    Item=constants.wdExportDocumentWithMarkup,
                                    CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
        except Exception as e:
            print('---转换失败，异常是：{}'.format(e))
        finally:
            print('---转换成功：', pdf_name)
    word.Quit(constants.wdDoNotSaveChanges)  # 退出
    print('---全部转换完毕')


def name2json():
    '''
    (2) 读取当前目录下的文件名称，存为 json 文件
    '''
    print('---正在生成 json 文件...')
    files = os.listdir(g_path)
    names = [os.path.splitext(file)[0]
             for file in files if not file.endswith('.exe')]
    obj = {'files': files, 'names': names}
    with open(f'output.json', 'w', encoding='utf-8') as f:
        f.write(json.dumps(obj, ensure_ascii=False))
    print('---生成完毕')


def main():
    '''
    (3) 程序入口
    '''
    global g_path
    global g_output_path
    g_path = os.getcwd()
    g_output_path = g_path+'/output'
    print('---请输入以下序号选择功能：')
    print('---1.转换当前目录下的 word 为 pdf')
    print('---2.读取当前目录下文件名称，存为 json 文件')
    option = input()
    if option == '1':
        word2pdf()
    elif option == '2':
        name2json()


if __name__ == '__main__':
    main()
