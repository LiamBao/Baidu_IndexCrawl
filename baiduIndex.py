#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import traceback
import time
import xlwt
import logging
import copy
import re
import requests
import urllib
import json
from datetime import datetime, timedelta
from selenium import webdriver
import queue
import sys
import threading
import hashlib 
import io
from PIL import Image
import ctypes  


###########################
########################### config.py  常量参数配置
###########################

global start_date,end_date
browser_driver = 'Firefox'  
# 百度用户名
user_name = '18964715224'
# 百度密码
password = 'cicdata02'
# 百度登陆链接
login_url = ('https://passport.baidu.com/v2/?login&tpl=mn&u='
             'http%3A%2F%2Fwww.baidu.com%2F')
# 一周
one_week_trend_url = ('http://index.baidu.com/?tpl=trend&type=0'
                      '&area=0&time=12&word={word}')
# 区间
time_range_trend_url = ('http://index.baidu.com/?tpl=trend&type=0&area=0'
                        '&time={start_date}|{end_date}&word={word}')
# api
all_index_url = ('http://index.baidu.com/Interface/Search/getAllIndex/'
                 '?res={res}&res2={res2}&startdate={start_date}'
                 '&enddate={end_date}')
# 图片信息的api
index_show_url = ('http://index.baidu.com/Interface/IndexShow/show/?res='
                  '{res}&res2={res2}&classType=1&res3[]={enc_index}'
                  '&className=view-value&{t}'
                  )
# 判断登陆状态的地址
user_center_url = 'http://i.baidu.com/'
# 判断登陆的标记
login_sign = 'http://passport.baidu.com/?logout'
# 线程数
num_of_threads = 40
# 关键词index的默认区间开始  在传参里修改
start_date = '2016-01-01'
# 关键词index的区间结束
end_date = '2016-08-31'

# 输出的格式，暂时只支持excel
# extension = 'excel'
# 输出的文件夹路径，可以自定义
out_file_path = './output'
# 关键词任务的文件路径，可以自定义
keywords_task_file_path = './task.txt'

# 要获取趋势的类别，默认是三种趋势都获取。all代表整体趋势，pc代表PC趋势, wise代表移动趋势
index_type_list = ['all', 'pc', 'wise']


STD_INPUT_HANDLE = -10  
STD_OUTPUT_HANDLE= -11  
STD_ERROR_HANDLE = -12  
  
FOREGROUND_BLACK = 0x0  
FOREGROUND_BLUE = 0x01 # text color contains blue.  
FOREGROUND_GREEN= 0x02 # text color contains green.  
FOREGROUND_RED = 0x04 # text color contains red.  
FOREGROUND_INTENSITY = 0x08 # text color is intensified.  
  
BACKGROUND_BLUE = 0x10 # background color contains blue.  
BACKGROUND_GREEN= 0x20 # background color contains green.  
BACKGROUND_RED = 0x40 # background color contains red.  
BACKGROUND_INTENSITY = 0x80 # background color is intensified.  

class Color:  
    ''''' See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winprog/winprog/windows_api_reference.asp 
    for information on Windows APIs.'''  
    std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)  
      
    def set_cmd_color(self, color, handle=std_out_handle):  
        """(color) -> bit 
        Example: set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | FOREGROUND_INTENSITY) 
        """  
        bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)  
        return bool  
      
    def reset_color(self):  
        self.set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)  
      
    def print_red_text(self, print_text):  
        self.set_cmd_color(FOREGROUND_RED | FOREGROUND_INTENSITY)  
        print (print_text)  
        self.reset_color()  
          
    def print_green_text(self, print_text):  
        self.set_cmd_color(FOREGROUND_GREEN | FOREGROUND_INTENSITY)  
        print (print_text)  
        self.reset_color()  
      
    def print_blue_text(self, print_text):   
        self.set_cmd_color(FOREGROUND_BLUE | FOREGROUND_INTENSITY)  
        print (print_text)
        self.reset_color()   


###########################
###########################  log.py  日志格式
###########################

# 定义handler的输出格式formatter
formatter = logging.Formatter(
    # '[%(levelname)1.1s %(asctime)s %(module)s:%(lineno)d] %(message)s'
    '[%(levelname)1.1s %(asctime)s %(module)s] %(message)s'
)

logger = logging.getLogger(__name__)
channel = logging.StreamHandler()
logger.setLevel(logging.DEBUG)
logger.addHandler(channel)

channel.setFormatter(formatter)
channel.setFormatter(formatter)


###########################
###########################api.py  百度API抓包调用
###########################


UserAgent = ('Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36')

HUMAN_HEADERS = {
    'Accept': ('text/html,application/xhtml+xml,application/xml;q=0.9,'
               'image/webp,*/*;q=0.8'),
    'User-Agent': UserAgent,
    'Accept-Language':'en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4',
    'Connection':'keep-alive',
    'Accept-Encoding': 'gzip, deflate, sdch'
}


class Api(object):
    def __init__(self, cookie):
        self.headers = copy.deepcopy(HUMAN_HEADERS)
        self.headers.update({'Cookie': cookie})

    def get_all_index_html(self, all_index_url):
        r = requests.get(all_index_url, headers=self.headers)
        return r.json()

    def get_index_show_html(self, index_show_url):
        r = requests.get(index_show_url, headers=self.headers)
        content = r.json()['data']['code'][0]
        img_url = re.findall('(?is)"(/Interface/IndexShow/img/[^"]*?)"', content)
        img_url = "http://index.baidu.com%s" % img_url[0]

        regex = ('(?is)<span class="imgval" style="width:(\d+)px;">'
                 '<div class="imgtxt" style="margin-left:-(\d+)px;">')
        result = re.findall(regex, content)
        skip_info = result if result else list()
        return img_url, skip_info

    def get_value_from_url(self, img_url, index_skip_info):
        r = requests.get(img_url, headers=self.headers)
        return get_num(r.content, index_skip_info)



###########################
########################### browser.py  SELENIUM包，浏览器切图
###########################


class BaiduBrowser(object):
    def __init__(self, cookie_json='', check_login=True):
        if not browser_driver:
            browser_driver_name = 'Firefox'
        else:
            browser_driver_name = browser_driver
        browser_driver_class = getattr(webdriver, browser_driver_name)
        self.browser = browser_driver_class()
        # 设置超时时间
        self.browser.set_page_load_timeout(20)
        # 设置脚本运行超时时间
        self.browser.set_script_timeout(20)
        # 百度用户名
        self.user_name = user_name
        # 百度密码
        self.password = password
        self.cookie_json = cookie_json
        self.api = None
        self.cookie_dict_list = []

        self.init_api(check_login=check_login)

    def is_login(self):
        # 如果初始化BaiduBrowser时传递了cookie信息，则检测一下是否登录状态
        self.login_with_cookie(self.cookie_json)
        # 访问待检测的页面
        self.browser.get(user_center_url)
        html = self.browser.page_source
        # 检测是否有登录成功标记
        return login_sign in html

    def init_api(self, check_login=True):
        # 判断是否需要登录
        need_login = False
        if not self.cookie_json:
            need_login = True
        elif check_login and not self.is_login():
            need_login = True
        # 执行浏览器自动填表登录，登录后获取cookie
        if need_login:
            self.login(self.user_name, self.password)
            self.cookie_json = self.get_cookie_json()
        cookie_str = self.get_cookie_str(self.cookie_json)
        # 获取到cookie后传给api
        self.api = Api(cookie_str)

    def get_date_info(self, start_date, end_date):
        # 如果start_date和end_date中带有“-”，则替换掉
        if start_date.find('-') != -1 and end_date.find('-') != -1:
            start_date = start_date.replace('-', '')
            end_date = end_date.replace('-', '')
        # start_date和end_date转换成datetime对象
        start_date = datetime.strptime(start_date, '%Y%m%d')
        end_date = datetime.strptime(end_date, '%Y%m%d')

        # 循环start_date和end_date的差值，获取区间内所有的日期
        date_list = []
        temp_date = start_date
        while temp_date <= end_date:
            date_list.append(temp_date.strftime("%Y-%m-%d"))
            temp_date += timedelta(days=1)
        start_date = start_date.strftime("%Y-%m-%d")
        end_date = end_date.strftime("%Y-%m-%d")
        return start_date, end_date, date_list

    def get_one_day_index(self, date, url):
        try_num = 0
        try_max_num = 5
        while try_num < try_max_num:
            try:
                try_num += 1

                img_url, val_info = self.api.get_index_show_html(url)

                value = self.api.get_value_from_url(img_url, val_info)
                break
            except:
                pass
        logger.info('date:%s,  value:%s' % (date, value))
        return value.replace(',', '')

    def get_baidu_index_by_date_range(self, keyword, start_date, end_date,
                                      type_name):
        # 根据区间获取关键词的索引值
        url = time_range_trend_url.format(
            start_date=start_date, end_date=end_date,
            word=urllib.parse.quote(keyword.encode('gbk'))
        )
        self.browser.get(url)
        # 执行js获取后面所需的res和res2的值
        res = self.browser.execute_script('return PPval.ppt;')
        res2 = self.browser.execute_script('return PPval.res2;')


        if  res == None  or res2 == None:
            return None

        start_date, end_date, date_list = self.get_date_info(
            start_date, end_date
        )

        # 拼接api的url
        url = all_index_url.format(
            res=res, res2=res2, start_date=start_date, end_date=end_date
        )

        try:

            all_index_info = self.api.get_all_index_html(url)
            indexes_enc = all_index_info['data'][type_name][0]['userIndexes_enc']
            enc_list = indexes_enc.split(',')
            wm = WorkManager(num_of_threads)
        except Exception as e:
            print(e)
        


        for index, _ in enumerate(enc_list):
            url = index_show_url.format(
                res=res, res2=res2, enc_index=_, t=int(time.time()) * 1000
            )

            date = date_list[index]

            wm.add_job(date, self.get_one_day_index, date, url)

        wm.start()
        wm.wait_for_complete()

        baidu_index_dict = wm.get_all_result_dict_from_queue()

        return baidu_index_dict

    def _get_index_period(self, keyword):
        # 拼接一周趋势的url
        url = one_week_trend_url.format(

            word=urllib.parse.quote(keyword.encode('gbk'))
        )
        self.browser.get(url)
        # 获取下方api要用到的res和res2的值
        res = self.browser.execute_script('return PPval.ppt;')
        res2 = self.browser.execute_script('return PPval.res2;')
        start_date, end_date = self.browser.execute_script(
            'return BID.getParams.time()[0];'
        ).split('|')
        start_date, end_date, date_list = self.get_date_info(
            start_date, end_date
        )
        url = all_index_url.format(
            res=res, res2=res2, start_date=start_date, end_date=end_date
        )
        all_index_info = self.api.get_all_index_html(url)
        start_date,end_date = all_index_info['data']['all'][0]['period'].split('|')
        # 重置start_date, end_date，以api返回的为准
        start_date, end_date, date_list = self.get_date_info(
            start_date, end_date
        )
        logger.info('all_start_date:%s, all_end_date:%s' %(start_date, end_date))
        return date_list

    def get_baidu_index(self, keyword, type_name):
        global start_date,end_date
        if (start_date and end_date):
            _, _, date_list = self.get_date_info(start_date,end_date)
        else:
            date_list = self._get_index_period(keyword)
        baidu_index_dict = dict()
        start = 0
        skip = 180
        end =len(date_list)
        while start < end:
            try:
                start_date = date_list[start]
                if start + skip >= end -1:
                    end_date = date_list[-1]
                else:
                    end_date = date_list[start + skip]
                result = self.get_baidu_index_by_date_range(
                    keyword, start_date, end_date, type_name
                )
                if result == None:
                    return  None
                baidu_index_dict.update(result)
                start += skip + 1
            except:
                import traceback

                print (traceback.format_exc())
        return baidu_index_dict

    def login(self, user_name, password):
        # login_url = login_url

        self.browser.get(login_url)

        # 自动填写表单并提交，如果出现验证码需要手动填写
        user_name_obj = self.browser.find_element_by_id(
            'TANGRAM__PSP_3__userName'
        )
        user_name_obj.send_keys(user_name)
        ps_obj = self.browser.find_element_by_id('TANGRAM__PSP_3__password')
        ps_obj.send_keys(password)
        sub_obj = self.browser.find_element_by_id('TANGRAM__PSP_3__submit')
        sub_obj.click()

        # 如果页面的url没有改变，则继续等待
        while self.browser.current_url == login_url:
            time.sleep(3)

    def close(self):
        self.browser.quit()

    def get_cookie_json(self):
        return json.dumps(self.browser.get_cookies())

    def get_cookie_str(self, cookie_json=''):
        if cookie_json:
            cookies = json.loads(cookie_json)
        else:
            cookies = self.browser.get_cookies()
        return '; '.join(['%s=%s' % (item['name'], item['value'])
                          for item in cookies])

    def login_with_cookie(self, cookie_json):
        self.browser.get('https://www.baidu.com/')
        for item in json.loads(cookie_json):
            try:
                self.browser.add_cookie(item)
            except:
                continue




###########################
########################### threadid.py 线程处理
###########################


class Worker(threading.Thread):  # 处理工作请求
    def __init__(self, work_queue, result_queue, **kwargs):
        threading.Thread.__init__(self, **kwargs)
        self.setDaemon(True)
        self.work_queue = work_queue
        self.result_queue = result_queue

    def run(self):
        while 1:
            try:
                task_key, func, args, kwargs = self.work_queue.get(False)  # get task
                res = func(*args, **kwargs)
                self.result_queue.put((task_key, res))  # put result
            except queue.Empty:
                break


class WorkManager:  # 线程池管理,创建
    def __init__(self, num_of_workers=10):
        self.work_queue = queue.Queue() # 请求队列
        self.result_queue = queue.Queue()  # 输出结果的队列
        self.workers = []
        self.init_threads(num_of_workers)

    def init_threads(self, num_of_workers):
        for i in range(num_of_workers):
            worker = Worker(self.work_queue, self.result_queue)  # 创建工作线程
            self.workers.append(worker)  # 加入到线程队列

    def start(self):
        for w in self.workers:
            w.start()

    def wait_for_complete(self):
        while len(self.workers):
            worker = self.workers.pop()  # 从池中取出一个线程处理请求
            worker.join()
            if worker.isAlive() and not self.work_queue.empty():
                self.workers.append(worker)  # 重新加入线程池中

    def add_job(self, task_key, func, *args, **kwargs):
        self.work_queue.put((task_key, func, args, kwargs))  # 向工作队列中加入请求

    def get_all_result_dict_from_queue(self):
        all_result_dict = {}
        while not self.result_queue.empty():
            task_key, result = self.result_queue.get(False)
            all_result_dict[task_key] = result
        return all_result_dict


###########################
########################### img_util
###########################

WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
IMG_MODEL_FOLDER = os.path.join(os.path.dirname(__file__), 'img_model')
img_value_dict = dict()


def get_num(img_data, index_skip_info):
    try:
        f = io.BytesIO(img_data)
        #得到一个图像的实例对象 img
        img = Image.open(f)
        width, height = img.size
        counter = 0
        last_width = 0
        for skip_w, skip_x in index_skip_info:
            counter += 1
            skip_w = int(skip_w)
            skip_x = int(skip_x)
            box = (skip_x, 0, skip_x + skip_w, height)
            new_img = img.crop(box)
            if counter == 1:
                end_img = Image.new('RGB', (100, new_img.size[1]))
            end_img.paste(new_img, (last_width, 0))
            last_width += new_img.size[0]
    except Exception as err:
        print(err)
    return get_value_from_img(img=end_img)



def get_value_from_img(fp=None, img=None):
    if not fp and not img:
        raise Exception('param error')
    if not img and fp:
        img = Image.open(fp)
    img = img.convert('RGB')
    img_data = img.load()
    img_width, img_height = img.size
    for x in range(img_width):
        for y in range(img_height):
            if img_data[x, y] != WHITE:
                img_data[x, y] = WHITE
            else:
                img_data[x, y] = BLACK
    small_imgs = split_img(img, img_data, img_width, img_height)
    return get_value_from_small_imgs(small_imgs)


def get_value_from_small_imgs(small_imgs):
    global img_value_dict
    value = []
    for img in small_imgs:
        key = get_md5(img)
        value.append(img_value_dict[key])
    return "".join(value)


def split_img(img, img_data, img_width, img_height):
    imgs = []
    split_info = []
    left = right = top = bottom = 0
    y_set = set()
    for x in range(img_width):
        all_is_white = True
        for y in range(img_height):
            if img_data[x, y] == WHITE:
                continue
            all_is_white = False
            if not left:
                left = x
            y_set.add(y)
        if all_is_white and left and not right:
            right = x
            top = min(y_set)
            bottom = max(y_set)
            split_info.append((left, right, top, bottom))
            left = right = top = bottom = 0
            y_set = set()
    for left, right, top, bottom in split_info:
        box = (left, top - 1, right, bottom + 1)
        new_img = img.crop(box)
        imgs.append(new_img)
    return imgs


def get_md5(img):
    content_list = []
    img = img.convert('RGB')
    img_data = img.load()
    img_width, img_height = img.size
    for x in range(img_width):
        for y in range(img_height):
            content = 'x:{0},y:{1},{2}'.format(x, y, img_data[x, y])
            content_list.append(content)
    lllll="".join(content_list)
    lllll = hashlib.md5(lllll.encode(encoding='gb2312'))
    # return hashlib.md5("".join(content_list)).hexdigest()
    return  lllll.hexdigest()

def _load_imgs():
    global img_value_dict
    file_name_list = os.listdir(IMG_MODEL_FOLDER)
    for file_name in file_name_list:
        value = file_name.split('.')[0]
        file_path = os.path.join(IMG_MODEL_FOLDER, file_name)
        img = Image.open(file_path)
        key = get_md5(img)
        img_value_dict[key] = value


_load_imgs()


def write_excel(excel_file, data_list):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(u'工作表1')
    row = 0
    ws.write(row, 0, u'关键词')
    ws.write(row, 1, u'日期')
    ws.write(row, 2, u'类型')
    ws.write(row, 3, u'指数')
    row = 1
    for result in data_list:
        col = 0
        for item in result:
            ws.write(row, col, item)
            col += 1
        row += 1

    wb.save(excel_file)




print('*'*40)
print('##  Python  3.4')
print('##  Author  Liam')
print('##  Data    8/31/2016')
print('##  Crawl   BaiduIndexData')
print('*'*40)
print('\r\n')
time.sleep(1.5)
logger.info(u'确保账号能登录百度, 注意输入页面验证码!')
s = BaiduBrowser()
fp = open(keywords_task_file_path, 'rb')
task_list = fp.readlines()
fp.close()


root = os.path.dirname(os.path.realpath(__file__))
result_folder = os.path.join(root, out_file_path)
if not os.path.exists(result_folder):
    os.makedirs(result_folder)
type_list = index_type_list

for keyword in task_list:
    try:
        keyword = keyword.strip()
        if not keyword:
            continue
        keyword_unicode = keyword.decode('utf-8')
        keyword_unicode = keyword_unicode.split('#*#',3)
        start_date = keyword_unicode[1]
        end_date = keyword_unicode[2]
        print ('\r\nStarting Parsing Keyword: '+str(keyword_unicode) )
        for type_name in type_list:

                baidu_index_dict = s.get_baidu_index(
                    keyword_unicode[0], type_name
                )
                if baidu_index_dict == None:
                    logger.info(' 未收录该词条')
                    continue
                date_list = sorted(baidu_index_dict.keys())

                file_name = '%s_%s.xls' % (keyword_unicode[0], type_name)
                file_path = os.path.join(result_folder, file_name)

                data_list = []
                for date in date_list:
                    value = baidu_index_dict[date]
                    data_list.append((keyword_unicode[0], date, type_name, value))
                write_excel(file_path, data_list)
    except:
        # print (traceback.format_exception())
        print('Parsing Error')
s.close()
