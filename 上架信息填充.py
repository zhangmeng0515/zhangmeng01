import pandas as pd
from selenium import webdriver
import time
import os

class zidongSpider(object):
    def __init__(self):
        # 打开登录页面
        self.option = webdriver.ChromeOptions()
        # 去掉Chrome 正受到自动测试软件的控制
        self.option.add_argument('disable-infobars')
        # 创建Chrome驱动器对象
        chromedriver = '{}'.format(os.path.abspath('chromedriver.exe'))
        self.driver = webdriver.Chrome(chromedriver, chrome_options=self.option)
        self.driver.maximize_window()
        time.sleep(0.5)
        self.driver.get('https://abdpop.secoo.com/login?redirectUrl=')
        # self.driver.find_element_by_id("code").send_keys("3498")
        # time.sleep(0.5)
        # self.driver.find_element_by_id("password").send_keys("Fm159147")
        time.sleep(60)

    def __del__(self):
        # 当本对象销毁的时候退出浏览器
        #  print("over%", time.strftime('%Y.%m.%d.%H.%I.%M.%S', time.localtime(time.time())))

        self.driver.quit()

    def get_data_list(self, i, skuid, html, length, width, height, weight):
        try:
            # 1. 准备URL个发送请求获取页面内容
            self.driver.get('https://abdpop.secoo.com/productInfo/modify/{}.do'.format(skuid))
            # 物流信息插入
            time.sleep(1)
            self.driver.find_element_by_xpath("//input[@name='length']").clear()
            time.sleep(0.25)
            self.driver.find_element_by_xpath("//input[@name='length']").send_keys(length)
            time.sleep(0.25)
            self.driver.find_element_by_xpath("//input[@name='width']").clear()
            time.sleep(0.25)
            self.driver.find_element_by_xpath("//input[@name='width']").send_keys(width)
            time.sleep(0.25)
            self.driver.find_element_by_xpath("//input[@name='height']").clear()
            time.sleep(0.25)
            self.driver.find_element_by_xpath("//input[@name='height']").send_keys(height)
            time.sleep(0.25)
            self.driver.find_element_by_xpath("//input[@name='weight']").clear()
            time.sleep(0.25)
            self.driver.find_element_by_xpath("//input[@name='weight']").send_keys(weight)
            time.sleep(0.25)
            # 包装清单插入
            # self.driver.find_element_by_xpath("//textarea[@class='ng-pristine ng-valid']").clear()
            # time.sleep(2)
            # self.driver.find_element_by_xpath("//textarea[@class='ng-valid ng-dirty']").send_keys("111")
            # 清除描述页内容
            time.sleep(0.5)
            self.driver.execute_script("scrollTo(0, 5500)")
            self.driver.switch_to.frame('ueditor_0')
            time.sleep(0.5)
            self.driver.find_element_by_tag_name('body').clear()
            self.driver.switch_to.default_content()
            time.sleep(0.5)
            self.driver.find_element_by_xpath("//div[@id='edui4']").click()
            time.sleep(0.5)
            # 商品描述插入
            # print type(HTML)
            time.sleep(0.5)
            self.driver.find_element_by_xpath("//textarea[@autocapitalize='off']").send_keys(u'{}'.format(html))
            time.sleep(0.5)
            self.driver.find_element_by_xpath("//div[@class='detailed_sub']/a[1]").click()
            time.sleep(0.5)
            self.f.write("Yes{}&{}".format(i+1, time.strftime('%Y.%m.%d.%H.%I.%M.%S', time.localtime(time.time()))) + "\n")
            self.f.close()

        except:
            self.f.write("No{}&{}".format(i+1, time.strftime('%Y.%m.%d.%H.%I.%M.%S', time.localtime(time.time()))) + "\n")
            self.f.close()

    def run(self):
        # 新建运行日志文档

        # 读取文件所有内用
        info = pd.read_excel(r"{}".format(os.path.abspath('moban.xlsx')))
        # 1. 准备参数
        info = info.fillna('')
        print(os.path.abspath('moban.xlsx'))
        for i in range(len(info['spuid'])):
            self.f = open(os.path.dirname(__file__) + "/logging.txt", "a+")
            skuid = str(info['spuid'][i])
            html = info['html'][i]
            length = str(info['length'][i])
            width = str(info['width'][i])
            height = str(info['height'][i])
            weight = str(info['weight'][i])
            self.get_data_list(i, skuid, html, length, width, height, weight)


# 测试
if __name__ == '__main__':
    zfs = zidongSpider()
    zfs.run()