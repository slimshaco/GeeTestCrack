import random
import time
import math
from io import BytesIO
from PIL import Image, ImageChops
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

class NationalCompanyCredit():

    def __init__(
            self,
            url,
            search_word = '腾讯',
            input_id = 'keyword',
            search_elemnt_id = 'btn_query',
            gt_info_class_name = 'gt_info_content',
            result_class_name = 'search_result'
                 ):
        '''

        :param url:
        :param search_word:
        :param input_id:
        :param search_elemnt_id:
        :param result_class_name:
        '''

        self.url = url
        self.input_id = input_id
        self.result_class_name = result_class_name
        self.search_element_id = search_elemnt_id
        self.search_word = search_word
        self.gt_info_class_name = gt_info_class_name

        dcap = dict(DesiredCapabilities.PHANTOMJS)
        dcap["phantomjs.page.settings.userAgent"] = (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.98 Safari/537.36"
        )

        self.driver = webdriver.PhantomJS(desired_capabilities=dcap)
        # self.driver = webdriver.Chrome()
        self.driver.set_window_size(1280, 800)


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

    def find_diff(self, im1, im2):
        diff = ImageChops.difference(im1, im2).convert('L')  # 把两张照片融合在一起， 相同的地方会全部变成黑色
        w, h = diff.size
        for i in range(w):
            for j in range(h):
                if diff.getpixel((i, j)) > 70:
                    # if diff.getpixel((i, j))[0] >= 60 and diff.getpixel((i, j))[1] >= 60 and diff.getpixel((i, j))[2] >=60 :
                    print('找到像素点为', i)
                    return i

    def get_offset(self,offset):
        array_x = [1.0 / 2, 3.0 / 5, 4.0 / 7, 5.0 / 9]  # 滑动的距离的乘积，最好控制在递减一半的情况
        # array_x = [1.0 / 2, 3.0 / 5, 4.0 / 7, 5.0 / 9, 2.0 / 3]
        total = 0
        while total <= offset:
            left = offset - total
            if left <= 4:  # 当最后滑动距离小于4时，不再分割滑动距离
                yield left
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
        action.move_by_offset(240, 0).perform()
        time.sleep(0.8)
        im2 = self.get_crop_img()
        action.release().perform()
        # 必须加延时，否则第二次滑动不了，原因是要等滑块恢复原位才能二次滑动
        time.sleep(1.8)
        return im1, im2


    def crack(self):
        im1, im2 = self.get_page()
        offset = self.find_diff(im1, im2) - 7
        knob = self.driver.find_element_by_class_name("gt_slider_knob")
        ActionChains(self.driver).move_to_element_with_offset(knob, 21, 21).click_and_hold().perform()
        for o in self.get_offset(offset):
            y = -1  # 关键点就在y轴上面，设为-1成功率最高
            ActionChains(self.driver).move_by_offset(o, y).perform()
            print('移动了{}距离'.format(o))
            time.sleep(random.choice([2, 4]) / 100)
        ActionChains(self.driver).release().perform()
        time.sleep(0.6)
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

