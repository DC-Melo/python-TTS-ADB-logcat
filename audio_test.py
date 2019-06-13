# -*- coding: utf-8 -*-
import time
import os

import openpyxl
import xlrd
import xlwt
import pandas
import pyttsx3
import logging

debug_dir = os.path.join(os.getcwd(), "debug")
config_dir = os.path.join(os.getcwd(), "config")
output_dir = os.path.join(os.getcwd(), "output")

if not os.path.exists(debug_dir):
    os.mkdir(debug_dir)
if not os.path.exists(config_dir):
    os.mkdir(config_dir)
if not os.path.exists(output_dir):
    os.mkdir(output_dir)
logging.basicConfig(level=logging.DEBUG,  # 控制台打印的日志级别
                    filename=os.path.join(debug_dir, "log.log"),
                    filemode='a',  # 模式，有w和a，w就是写模式，每次都会重新写日志，覆盖之前的日志 a是追加模式，默认如果不写的话，就是追加模式
                    format='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'  # 日志格式
                    )
file_list = []
while len(file_list) == 0:
    for root, dirs, files in os.walk(config_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.txt':
                file_list.append(os.path.join(root, file))
    if len(file_list) > 0:
        for index, xlsx_file in enumerate(file_list):
            print("%s:%s"%(index+1, xlsx_file.strip()))
    else:
        print("请将您的语料文件放到", config_dir,"目录。")
        time.sleep(10)
global TEST_FILE
test_file_index = str(input("请输入："))
print("您选择语料文件如下：".center(100, "="))
print(test_file_index, file_list[int(test_file_index)-1])
TEST_XLSX = os.path.join(config_dir, file_list[int(test_file_index)-1])
engine = pyttsx3.init()   # 初始化
engine.setProperty('voice','zh')
with open(TEST_XLSX, encoding='gbk') as f:
    for line in f.readlines():
        print(line)
        engine.say(line)
        engine.runAndWait()
        print("检查结果")




