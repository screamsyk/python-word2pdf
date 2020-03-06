# 内置模块
import os
import configparser
import json

# 全局变量
g_config = None


def get_config():
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
    get_config()
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
