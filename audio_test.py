# -*- coding: utf-8 -*-
import time
import os
import string

import xlrd
import xlwt
import pyttsx3
import logging
from py3adb import ADB

CYCLETIME=2

def str_count_zh(str):
    '''找出字符串中的中英文、空格、数字、标点符号个数'''
    count_en = count_dg = count_sp = count_zh = count_pu = 0
    for s in str:
        # 英文
        if s in string.ascii_letters:
            count_en += 1
        # 数字
        elif s.isdigit():
            count_dg += 1
        # 空格
        elif s.isspace():
            count_sp += 1
        # 中文
        elif s.isalpha():
            count_zh += 1
        # 特殊字符
        else:
            count_pu += 1
    return count_zh
def printChooseList(title,list,begin,body,end):
    n=110
    if len(list)>0:
        print(title.center(n-str_count_zh(title),body))
        for i, val in enumerate(list):
            if len(val)>0:
                print(begin,i,":",val.ljust(n-3-len(str(i))-str_count_zh(val)-4),end)
        print(n*body)
def printChooseDict(title,dict,begin,body,end):
    n=110
    if len(dict)>0:
        print(title.center(n-str_count_zh(title),body))
        for key in dict:
            print(begin, key, ":", dict[key].ljust(n - 3 - len(str(i)) - str_count_zh(dict[key]) - 4), end)
        print(n * body)
input_dir = os.path.join(os.getcwd(), "input")
output_dir = os.path.join(os.getcwd(), "output")
outputLog_dir = os.path.join(os.getcwd(), "output", "logcat")
debug_dir = os.path.join(os.getcwd(), "debug")
if not os.path.exists(debug_dir):
    os.mkdir(debug_dir)
if not os.path.exists(input_dir):
    os.mkdir(input_dir)
if not os.path.exists(output_dir):
    os.mkdir(output_dir)
if not os.path.exists(outputLog_dir):
    os.mkdir(outputLog_dir)
logging.basicConfig(level=logging.DEBUG,  # 控制台打印的日志级别
                    filename=os.path.join(debug_dir, "log.log"),
                    filemode='a',  # 模式，有w和a，w就是写模式，每次都会重新写日志，覆盖之前的日志 a是追加模式，默认如果不写的话，就是追加模式
                    format='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'  # 日志格式
                    )
file_dict = {}
while len(file_dict) == 0:
    i=0
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.xlsx':
                    i+=1
                    file_dict[str(i)] = os.path.join(root, file)
    if len(file_dict) > 0:
        printChooseDict("请选择测文件（.xlsx）", file_dict,"|","-","|")
    else:
        print("请将测试用例(.xlsx)放到", input_dir, "目录下！！！")
        time.sleep(10)
select_file_index = str(input("请输入序号（例如：1）："))
SELECT_FILE_DICT = {}
if select_file_index != '0':
    select_file_index = [item for item in select_file_index.split("&")]
    for item in select_file_index:
        SELECT_FILE_DICT[item] = file_dict[item]
else:
    TEST_TYPE = 0
printChooseDict("您选择了以下测试文件", SELECT_FILE_DICT,"|","-","|")
confirm = input("请确认测试文件，y/Y:确认 其他任意键:退出：")
if confirm not in ["y", "Y"]:
    print("退出测试")
    time.sleep(5)
    sys.exit(1)

TEST_FILE = os.path.join(input_dir, list(SELECT_FILE_DICT.values())[0])
book = xlrd.open_workbook(TEST_FILE)
sheet_dict ={}
for index, sheet in enumerate(book.sheet_names()):
    sheet_dict[str(index)]=sheet.strip()
printChooseDict(TEST_FILE+",请选择测试表单",sheet_dict,"|","-","|")


select_sheet_index = str(input("请输入表单序号（例如：1）："))
SELECT_SHEET_DICT = {}
if select_sheet_index != '0':
    select_sheet_index = [item for item in select_sheet_index.split("&")]
    for item in select_sheet_index:
        SELECT_SHEET_DICT[item] = sheet_dict[item]
else:
    TEST_TYPE = 0
printChooseDict("您选择了以下测试表单(sheet)", SELECT_SHEET_DICT,"|","-","|")
confirm = input("请确认测试表单sheet，y/Y:确认 其他任意键:退出：")
if confirm not in ["y", "Y"]:
    print("退出测试")
    time.sleep(5)
    sys.exit(1)


adbConn = ADB()
adbConn.set_adb_path(os.path.join(os.getcwd(), "tools", "adb.exe"))
adbConn.get_adb_path()
device_dict ={}
while True:
    device_tuple = adbConn.get_devices()
    if not device_tuple[1]:
        printChooseDict("请插入USB，开启ADB调试模式", device_dict, "!", "!", "!")
        continue
    else:
        for i,dev in enumerate(device_tuple[1]):
            device_dict[str(i)] = dev
        break
printChooseDict("请选择被测设备",device_dict,"|", "-", "|")
select_device_index = str(input("请输入被测设备序号（例如：1）："))
adbConn.set_target_device(device_dict[select_device_index])
adbConn.set_adb_root()
adbConn.connect_remote()
adbConn.start_server()
adbConn.get_state()
while not adbConn.check_path():
    adbPath = str(input("当前adb路径" + adbConn.get_adb_path() + "错误，请从新输入adb路径："))
    adbConn.set_adb_path(adbPath)
engine = pyttsx3.init()   # 初始化
engine.setProperty('voice', 'zh')
engine.say("开始测试")
engine.runAndWait()
for sheetName in SELECT_SHEET_DICT.values():
    select_sheet=book.sheet_by_name(sheetName)
    for row in range(1, select_sheet.nrows):
        tts_text = str(select_sheet.cell(row, 6).value)
        if len(tts_text.strip()) < 2 or tts_text.strip().startswith("#") or tts_text.strip().startswith('//'):
            continue
        print(tts_text.strip())
        engine.say("你好，小安")
        engine.runAndWait()
        engine.say(tts_text.replace("\n", ""))
        engine.runAndWait()
        adbConn.get_logcat("-c")
        time.sleep(30)
        log = adbConn.get_logcat("-d")
        print(log)
        logFileName = str(row).zfill(4) + "-" + time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime(time.time())) +".txt"
        fl = open(os.path.join(output_dir, "logcat", logFileName ), 'w', encoding='utf-8')
        fl.write("语音输入：" + tts_text + "\n")
        fl.write('\n'.join(str(log)))
        fl.close()