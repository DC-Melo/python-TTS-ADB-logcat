# -*- coding: utf-8 -*-
import os
import sys
import time
import string
import platform
import signal
try:
    import logging
except ImportError:
    os.system('pip3 install logging')
    import logging


class Logger(object):
    level_relations = {
        'debug': logging.DEBUG,
        'info': logging.INFO,
        'warning': logging.WARNING,
        'error': logging.ERROR,
        'crit': logging.CRITICAL
    }  # 日志级别关系映射

    def __init__(self, filename, level='info', when='D', back_count=3, fmt='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'):
        self.logger = logging.getLogger(filename)
        format_str = logging.Formatter(fmt)  # 设置日志格式
        self.logger.setLevel(self.level_relations.get(level))  # 设置日志级别
        sh = logging.StreamHandler()  # 往屏幕上输出
        sh.setFormatter(format_str)  # 设置屏幕上显示的格式
        th = handlers.TimedRotatingFileHandler(filename=filename, when=when, backupCount=back_count, encoding='utf-8')#往文件里写入#指定间隔时间自动生成文件的处理器
        # 实例化TimedRotatingFileHandler
        # interval是时间间隔，backupCount是备份文件的个数，如果超过这个个数，就会自动删除，when是间隔的时间单位，单位有以下几种：
        # S 秒
        # M 分
        # H 小时、
        # D 天、
        # W 每星期（interval==0时代表星期一）
        # midnight 每天凌晨
        th.setFormatter(format_str)  # 设置文件里写入的格式
        self.logger.addHandler(sh)  # 把对象加到logger里
        self.logger.addHandler(th)


def str_count_zh(str_word):
    # 找出字符串中的中英文、空格、数字、标点符号个数
    count_en = count_dg = count_sp = count_zh = count_pu = 0
    for s in str_word:
        if s in string.ascii_letters:  # 英文
            count_en += 1
        elif s.isdigit():  # 数字
            count_dg += 1
        elif s.isspace():  # 空格
            count_sp += 1
        elif s.isalpha():  # 中文
            count_zh += 1
        else:              # 特殊字符
            count_pu += 1
    return count_zh


def print_choose_list(title, format_list, begin, body, end):
    n = 110
    if len(format_list) > 0:
        print(title.center(n-str_count_zh(title), body))
        for i, val in enumerate(format_list):
            if len(val) > 0:
                print(begin, i, ":", val.ljust(n-3-len(str(i))-str_count_zh(val)-4), end)
        print(n*body)


def print_choose_dict(title, choose_dict, begin, body, end):
    n = 110
    if len(choose_dict) > 0:
        print(title.center(n-str_count_zh(title), body))
        for k in choose_dict.keys():
            print(begin, k, ":", choose_dict[k].ljust(n - 3 - len(str(k))- str_count_zh(str(choose_dict[k])) - 4), end)
        print(n * body)


def print_confirm_dict(title, confirm_dict, begin, body, end):
    n = 110
    if len(confirm_dict) > 0:
        # print(title.center(n-str_count_zh(title),body))
        for key in confirm_dict.keys():
            print(key, ":", confirm_dict[key].ljust(n - 3 - len(str(i)) - str_count_zh(confirm_dict[key]) - 4))
        # print(n * body)


if __name__ == '__main__':
    # config folder.
    from logging import handlers
    input_dir = os.path.join(os.getcwd(), "2-input")
    debug_dir = os.path.join(os.getcwd(), "5-debug")
    output_dir = os.path.join(os.getcwd(), "6-output")
    if not os.path.exists(debug_dir):
        os.mkdir(debug_dir)
    if not os.path.exists(input_dir):
        os.mkdir(input_dir)
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    log_name = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime(time.time()))+".log"
    log = Logger(os.path.join(debug_dir, log_name), level='debug')
    log.logger.info('sys.version_info:'+str(sys.version_info))
    log.logger.info('platform.system():'+str(platform.system()))
    log.logger.info('platform.python_version():'+str(platform.python_version()))
    log.logger.info('platform.architecture():'+str(platform.architecture()))
    # import package except install package
    try:
        import xlrd
    except ImportError as err:
        logging.error(err)
        f = os.popen('pip3 install xlrd', "r")
        logging.info(f.read())
        if f.read().upper().find("success".upper()):
            import xlrd
        else:
            logging.fatal("could not install this package.")
    try:
        import xlwt
    except ImportError as err:
        logging.error(err)
        f = os.popen('pip3 install xlwt', "r")
        logging.info(f.read())
        if f.read().upper().find("success".upper()):
            import xlwt
        else:
            logging.fatal("could not install this package.")
    try:
        import pyttsx3
    except ImportError as err:
        logging.error(err)
        f = os.popen('pip3 install pyttsx3', "r")
        logging.info(f.read())
        if f.read().upper().find("success".upper()):
            import pyttsx3
        else:
            logging.fatal("could not install this package.")
    try:
        from py3adb import ADB
    except ImportError as err:
        logging.error(err)
        f = os.popen('pip3 install py3adb', "r")
        logging.info(f.read())
        if f.read().upper().find("success".upper()):
            import py3adb
        else:
            logging.fatal("could not install this package.")
        from py3adb import ADB
    adbConn = ADB()
    adbConn.set_adb_path(os.path.join(os.getcwd(), "4-tool", "adb.exe"))
    if not adbConn.check_path():
        adbConn.set_adb_path(os.path.join(os.getcwd(), "4-tool", "adb"))
    while not adbConn.check_path():
        log.logger.error("there is no adb/adb.exe in the 4-tool folder")
        adbPath_input = str(input("当前adb路径" + str(adbConn.get_adb_path()) + "错误，请从新输入adb路径："))
        adbPath_input_list = adbPath_input.split('/')
        adbPath = ''
        for path in adbPath_input_list:
            adbPath = os.path.join(adbPath, path)
        adbConn.set_adb_path(adbPath)
        log.logger.info("set adb path:"+adbPath)
    #  find .xlsx files in 2-input
    file_dict = {}
    while len(file_dict) == 0:
        i = 0
        for root, dirs, files in os.walk(input_dir):
            for file in files:
                if os.path.splitext(file)[1] == '.xlsx':
                        i += 1
                        file_dict[str(i)] = os.path.join(root, file)
        if len(file_dict) > 0:
            print_choose_dict("请选择测试文件（.xlsx）", file_dict, "|", "-", "|")
        else:
            print("请将测试文件(.xlsx)放到", input_dir, "目录下！！！,等待10秒重新检测文件...")
            time.sleep(10)

    SELECT_FILE_DICT = {}
    while True:
        select_file_index = str(input("请输入序号（例如：1）："))
        select_file_index = [item for item in select_file_index.split("&")]
        if set(select_file_index) <= set(file_dict.keys()) and len(select_file_index) == 1:
            for item in select_file_index:
                SELECT_FILE_DICT[item] = file_dict[item]
            break
        else:
            print("输入有误，请重新输入！")

    print_confirm_dict("您选择了以下测试文件", SELECT_FILE_DICT, "|", "-", "|")
    confirm = input("y/Y/space/enter:确认 e/E:保存结果并退出")
    if confirm in ['e', "E"]:
        print("保存测试结果，并退出")
        time.sleep(5)
        sys.exit(1)

    TEST_FILE = os.path.join(input_dir, SELECT_FILE_DICT[select_file_index[0]])
    book = xlrd.open_workbook(TEST_FILE)
    SHEET_DICT ={}
    for index, sheet in enumerate(book.sheet_names()):
        SHEET_DICT[str(index + 1)] = sheet.strip()
    print_choose_dict(SELECT_FILE_DICT[select_file_index[0]] + ",请选择测试表单", SHEET_DICT, "|", "-", "|")

    SELECT_SHEET_DICT = {}
    while True:
        select_sheet_index = str(input("请输入表单序号（例如：1）："))
        select_sheet_index = [item for item in select_sheet_index.split("&")]
        if set(select_sheet_index) <= SHEET_DICT.keys():
            for item in select_sheet_index:
                SELECT_SHEET_DICT[item] = SHEET_DICT[item]
            break
        else:
            print("输入有误，请重新输入！")

    print_choose_dict("您选择了以下测试表单(sheet)", SELECT_SHEET_DICT, "|", "-", "|")
    confirm = input("y/Y/space/enter:确认 e/E:保存结果并退出")
    if confirm in ['e', "E"]:
        print("保存测试结果，并退出")
        time.sleep(5)
        sys.exit(1)

    DEVICE_DICT ={}
    while True:
        device_tuple = adbConn.get_devices()
        if not device_tuple[1]:
            log.logger.warning("未插入USB/未开启ADB调试模式")
            print_choose_dict("请插入USB，开启ADB调试模式", DEVICE_DICT, "!", "!", "!")
            continue
        else:
            log.logger.info(str(device_tuple))
            for index, dev in enumerate(device_tuple[1]):
                DEVICE_DICT[str(index + 1)] = dev
            break
    print_choose_dict("请选择被测设备", DEVICE_DICT, "|", "-", "|")
    select_device_index = str(input("请输入被测设备序号（例如：1）："))
    while True:
        select_device_index = [item for item in select_device_index.split("&")]
        if set(select_device_index) <= DEVICE_DICT.keys() and len(select_device_index) == 1:
            log.logger.info('devices:'+str(DEVICE_DICT))
            log.logger.info('select device:'+str(DEVICE_DICT[select_device_index[0]]))
            break
        else:
            print("设备选择错误，请重新输入!")
    # 链接设备
    adbConn.set_target_device(DEVICE_DICT[select_device_index[0]])
    if not adbConn.set_adb_root():
        log.logger.info("device could not root")
    if not adbConn.connect_remote():
        log.logger.info("device could not remote")
    adbConn.start_server()
    adbConn.get_state()
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
            time.sleep(10)
            logcat_log = os.popen(adbConn.get_adb_path()+" logcat -d").read()
            print(logcat_log)
            logFileName = str(row).zfill(4) + "-" + time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime(time.time())) +".txt"
            fl = open(os.path.join(output_dir, "logcat", logFileName ), 'w', encoding='utf-8')
            fl.write("语音输入：" + tts_text + "\n")
            fl.write(logcat_log)
            fl.close()