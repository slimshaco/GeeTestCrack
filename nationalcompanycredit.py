import random
import re
import time
import math
from PIL import Image, ImageChops
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from bs4 import BeautifulSoup
import xlsxwriter as wx
import pandas as pd
from xygdcx import *
from threading import Thread
from urllib.parse import quote,unquote



class NationalCompanyCredit():

    def __init__(
            self,
            url,
            search_word = '腾讯',
            input_id = 'keyword',
            search_elemnt_id = 'btn_query',
            gt_info_class_name = 'gt_info_content',
            result_class_name = 'search_result',
            result_list_class_name = 'a.search_list_item',
            distance = 7
                 ):
        '''

        :param url:  企业信用信息系统的网址
        :param search_word: 搜索关键词
        :param input_id: 搜索框id
        :param search_elemnt_id: 搜索按钮id
        :param gt_info_class_name: 验证的结果的class名
        :param result_class_name: 搜索结果的class name
        :param result_list_class_name: 搜索名单的class name
        :param distance: 需要减去的距离
        '''

        self.url = url
        self.input_id = input_id
        self.result_class_name = result_class_name
        self.search_element_id = search_elemnt_id
        self.search_word = search_word
        self.gt_info_class_name = gt_info_class_name
        self.result_list_class_name = result_list_class_name
        self.distance = distance
        dcap = dict(DesiredCapabilities.PHANTOMJS)
        dcap["phantomjs.page.settings.userAgent"] = (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.98 Safari/537.36"
        )
        #self.driver = webdriver.PhantomJS(desired_capabilities=dcap)
        self.driver = webdriver.Chrome('d:/work/chromedriver.exe')
        #self.driver.set_window_size(1280, 800)


    def get_crop_img(self):
        screenshot = self.driver.get_screenshot_as_png()
        screenshot = Image.open(BytesIO(screenshot))
        # screenshot.show()
        captcha_el = self.driver.find_element_by_class_name("gt_box")
        location = captcha_el.location
        size = captcha_el.size
        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']
        box = (left, top, right, bottom)
        print(box)
        if box[0] == 0:
            raise (Exception('======='))
        captcha_image = screenshot.crop(box)
        #captcha_image.save()  # "%s.png" % uuid.uuid4().hex
        print('截图成功')
        return captcha_image

    def find_diff_crop(self, im1, im2):
        diff = ImageChops.difference(im1, im2).convert('L')  # 把两张照片融合在一起， 相同的地方会全部变成黑色
        w, h = diff.size
        diff = diff.crop((62, 0, w, h))
        # diff.save('%s.png' %time.time())
        for i in range(w):
            for j in range(h):
                if diff.getpixel((i, j)) > 60:
                    # if diff.getpixel((i, j))[0] >= 60 and diff.getpixel((i, j))[1] >= 60 and diff.getpixel((i, j))[2] >=60 :
                    print('找到像素点为', i)
                    return i + 62

    def find_diff(self, im1, im2):
        diff = ImageChops.difference(im1, im2).convert('L')
        w, h = diff.size
        for i in range(w):
            for j in range(h):
                if diff.getpixel((i, j)) > 60:
                    return i


    def get_offset(self,offset):
        array_x = [ 3 / 5, 4 / 7, 5 / 9]  # 滑动的距离的乘积，最好控制在递减一半的情况
        # array_x = [1.0 / 2, 3.0 / 5, 4.0 / 7, 5.0 / 9]
        total = 0
        while total <= offset:
            left = offset - total
            if left < 5 :  # 当最后滑动距离小于5时，不再分割滑动距离
                for i in range(left):
                    yield 1
                break
            next_move = math.ceil(random.choice(array_x) * left)
            total += next_move
            yield next_move

    def get_page(self):
        self.driver.get(self.url)
        element = WebDriverWait(self.driver, 8).until(lambda x: x.find_element_by_id(self.input_id))
        element.clear()
        element.send_keys(self.search_word)
        # 必须要设置延时，不然图片加载不了
        time.sleep(0.5)
        # element = driver.find_element_by_id('btn_query')
        # element.click()
        element = WebDriverWait(self.driver, 8).until(lambda x: x.find_element_by_id(self.search_element_id))
        element.click()
        WebDriverWait(self.driver, 8).until(lambda x: x.find_element_by_class_name("gt_box"))
        time.sleep(0.5)
        im1 = self.get_crop_img()
        knob = self.driver.find_element_by_class_name("gt_slider_knob")
        action = ActionChains(self.driver)
        action.move_to_element_with_offset(knob, 21, 21).click_and_hold().perform()
        time.sleep(0.5)
        im2 = self.get_crop_img()

        return im1, im2


    def crack(self, distance):
        distance = self.distance
        im1, im2 = self.get_page()
        offset = self.find_diff_crop(im1, im2) - distance
        if offset == 55:
            action = ActionChains(self.driver)
            action.move_by_offset(240, 0).perform()
            time.sleep(0.8)
            im2 = self.get_crop_img()
            offset = self.find_diff(im1, im2) - 7
            action.release().perform()
            # 必须加延时，否则第二次滑动不了，原因是要等滑块恢复原位才能二次滑动
            time.sleep(1.8)
        knob = self.driver.find_element_by_class_name("gt_slider_knob")
        ActionChains(self.driver).move_to_element_with_offset(knob, 21, 21).click_and_hold().perform()
        for o in self.get_offset(offset):
            y = -1  # 关键点就在y轴上面，设为-1成功率最高
            ActionChains(self.driver).move_by_offset(o, y).perform()
            print('移动了{}距离'.format(o))
            time.sleep(random.choice([1,2]) / 100)
        ActionChains(self.driver).release().perform()
        time.sleep(0.8)
        element = self.driver.find_element_by_class_name(self.gt_info_class_name)
        status = element.text
        print("破解验证码的结果: ", status)
        if status.find("超过") != -1:
            element = WebDriverWait(self.driver, 8).until(lambda x: x.find_element_by_class_name(self.result_class_name))
            print(element.text)
            return self.driver.page_source
            # 使用phantomjs时不要使用refresh语句会报错

        else:
            print('破解失败')



    def get_company_url(self):
        source = self.crack(self.distance)
        if source:
            soup = BeautifulSoup(source, 'lxml')
            urls = soup.select(self.result_list_class_name)
            return urls[0].get('href')

    def get_company_info(self, url):
        global data,title
        self.driver.get(url)
        element = self.driver.find_element_by_id('addmore')
        element.click()
        soup = BeautifulSoup(self.driver.page_source, 'lxml')
        #print(soup)
        datas = soup.select('dd')
        titles = soup.select('dt')
        [data.append(re.sub(r'\s','',i.get_text())) for i in datas]
        [title.append(re.sub(r'\s','',i.get_text())) for i in titles]



def gjxy(name):
    print(time.ctime())
    c = NationalCompanyCredit(url='http://www.gsxt.gov.cn/index.html', search_word=name, input_id='keyword',
                              search_elemnt_id='btn_query', result_class_name='search_result',
                              result_list_class_name='.search_list_item.db',distance=7)
    url = 'http://www.gsxt.gov.cn' + c.get_company_url()
    c.get_company_info(url)


def write_excel(data, name):
    workbook = wx.Workbook('./{}.xlsx'.format(name),{'strings_to_urls':False})
    format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': \
            'white', 'font_size': 10, 'font_name': '微软雅黑', 'text_wrap': True})
    worksheet = workbook.add_worksheet()
    width = 2
    worksheet.set_column(0, width, 25)
    row = 0
    col = 0
        # .write方法  write（行,列,写入的内容,样式）
    for i in range(len(title)-1):
        worksheet.write(row, col, data[i], format)  # 在第一列的地方写入item
        worksheet.write(row, col + 1, data[i], format)  # 在第二列的地方写入data
        row += 1
    workbook.close()

data = []
title = []
gjxy('华为')
print(data)
