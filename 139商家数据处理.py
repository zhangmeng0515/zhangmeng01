# 更新139库存

import pandas as pd
from selenium import webdriver
import time
import os
import shutil
import numpy as np

os.chdir(r'E:\0000~急需处理')

class download_139info():
    def __init__(self, account, password):      
        self.option = webdriver.ChromeOptions()  
        # 去掉Chrome 正受到自动测试软件的控制
        self.option.add_argument('disable-infobars')
        # 创建Chrome驱动器对象
        self.chromedriver = '{}'.format(os.path.abspath('chromedriver.exe'))  # 需要先下载一个与谷歌浏览器版本匹配的chromedriver
        
        
        
     # 1. 登陆页面
    def login(self, account, password):
        self.driver = webdriver.Chrome(self.chromedriver, chrome_options=self.option)
        self.driver.maximize_window()  # 窗口最大化
        time.sleep(0.5)
        self.driver.get('http://101.251.111.14:8082/tplus/view/login.html')  # 打开登陆页面链接
        self.driver.find_element_by_id("userName").send_keys(account)  # 定位账号输入局域网，输入账号
        time.sleep(0.5)
        self.driver.find_element_by_id("password").send_keys(password)  # 定位密码输入局域网，输入密码
        time.sleep(1)
        self.driver.find_element_by_xpath("//button[@type='submit']").click()  # 点击确定,提交账号
        time.sleep(5)
        
        
     # 2.下载库存表
    def down_table(self):
        mouse_loc = self.driver.find_element_by_xpath('//li[@data-code="ST"]')  # 找到鼠标悬停位置
        webdriver.ActionChains(self.driver).move_to_element(mouse_loc).perform() # 移动鼠标到我的用户名
        time.sleep(2)
        self.driver.find_element_by_xpath('//a[@code="#ST3011"]').click()  # 点击库存核查
        time.sleep(15)
        self.driver.switch_to.frame("ST3011_iframe")  # 转换iframe
        imp_loc = self.driver.find_element_by_xpath('//div[@class="toolbar"]/button[@class="ant-btn toolbarItem ant-dropdown-trigger"]') 
        webdriver.ActionChains(self.driver).move_to_element(imp_loc).perform() # 移动鼠标找出下拉菜单
        time.sleep(3)
        self.driver.find_element_by_xpath('//ul[@role="menu"]/li[2]').click()  # 点击导出Excel表格
        time.sleep(15)
        self.driver.quit()

    # 运行类方法
    def run(self, account, password):
        self.login(account, password)
        self.down_table()
        # self.move_and_rename(account)

if __name__ == '__main__':
        temp = download_139info('陈志锋', 'a665097100')
        temp.run('陈志锋', 'a665097100')

shutil.move(r'C:\Users\Administrator\Downloads\现存量查询.xls', '现存量查询.xls')

a = pd.read_excel('现存量查询.xls')
a.columns = list(a.loc[1])  # loc按索引”名称“, 索引表格的行，   iloc按索引的位置序号
a.drop(a.columns.tolist()[a.columns.tolist().index('可用量') + 1:], axis=1, inplace=True)
a.drop(a.columns.tolist()[0], axis=1, inplace=True)
a.dropna(how='all',inplace=True)
a.drop(1, inplace=True)
a.dropna(subset=['可用量'], inplace=True)  # subset值必须用列表表示
a.dropna(subset=['货号'], inplace=True)
a.reset_index(drop=True, inplace=True)
a = a.fillna('').astype(str)


def get_col_list_after_insert(raw_col, insert_before_which_col, the_seq_to_insert):
    for each_field in the_seq_to_insert:
        raw_col.insert(raw_col.index(insert_before_which_col), each_field)
    return raw_col
col_seq = get_col_list_after_insert(a.columns.tolist(), '规格型号', ['品类']) 
col_seq = get_col_list_after_insert(col_seq, '材质', ['去格式货号(含商家编号)', '3498是否有', '4083是否有', '共有商品', '去格式货号&尺码']) 
col_seq = get_col_list_after_insert(col_seq, '厂家颜色编码', ['清洗尺码']) 
a = a.reindex(columns=col_seq)

def clean_category(source_col):
    temp_list = []
    for total_item in source_col:
        if '包' in total_item or '钱夹' in total_item or '皮套' in total_item or (('夹' in total_item or '夾' in total_item) and '夹克' not in total_item):
            temp_list.append('包袋')
        else:
            temp_list.append('其他')
    return temp_list
a['品类'] = clean_category(a['存货'])

import re
def remove_format(col_name):
    fuhao1 = re.compile('-')
    fuhao2 = re.compile('\/')
    fuhao3 = re.compile('\.')
    fuhao4 = re.compile(' ')
    fuhao5 = re.compile('\*')
    fuhao6 = re.compile('O')
    fuhao7 = re.compile('o')
    temp_list = []
    for each_term in col_name:
        c_step1 = fuhao1.sub('', each_term)
        c_step2 = fuhao2.sub('', c_step1)
        c_step3 = fuhao3.sub('', c_step2)
        c_step4 = fuhao4.sub('', c_step3)
        c_step5 = fuhao5.sub('', c_step4)
        c_step6 = fuhao6.sub('0', c_step5)
        c_step7 = fuhao7.sub('0', c_step6)
        c_step7 = c_step7 + '139'
        temp_list.append(c_step7)
    return temp_list
a['去格式货号(含商家编号)'] = remove_format(a['货号'])


gi3498 = pd.read_excel('3498' + '#' + time.strftime('%Y%m%d',time.localtime(time.time())) + '商品信息表.xlsx').fillna('').astype(str)
gi4083 = pd.read_excel('4083' + '#' + time.strftime('%Y%m%d',time.localtime(time.time())) + '商品信息表.xlsx').fillna('').astype(str)
gi3498 = gi3498[gi3498['商家ID']=='139']
gi4083 = gi4083[gi4083['商家ID']=='139']
gi3498.reset_index(drop=True, inplace=True)
gi4083.reset_index(drop=True, inplace=True)

def vlookup(find_value, start_col, object_col):
    try:
        num = list(start_col).index(find_value)
        return object_col[num]
    except:
        return ''

a['3498是否有'] = [vlookup(code, gi3498['去格式货号'], gi3498['去格式货号']) for code in a['去格式货号(含商家编号)']]
a['4083是否有'] = [vlookup(code, gi4083['去格式货号'], gi4083['去格式货号']) for code in a['去格式货号(含商家编号)']]
a['共有商品'] = ['' if a['3498是否有'][i]=='' and a['4083是否有'][i]=='' else a['去格式货号(含商家编号)'][i] for i in range(len(a['3498是否有']))]
def clean_size(source_col1, source_col2, refcol):
    temp_list = []
    for i,item in enumerate(refcol):
        if item=='包袋' or 'x' in source_col1[i]:
            temp_list.append('')
        else:
            c = re.compile('[\u4e00-\u9fa5]')
            match = c.match(source_col1[i])
            if source_col1[i]=='':
                value = source_col2[i]
            elif match:
                value = source_col2[i]
            else:
                value = source_col1[i]
            value = value.replace(' ','').replace('TU', '').replace('UNI', '').replace('cm', '')
            temp_list.append(value)
    return temp_list

a['清洗尺码'] = clean_size(a['规格型号'], a['尺码'], a['品类'])

a['去格式货号&尺码'] = ['' if item=='' else a['共有商品'][i] + a['清洗尺码'][i] for i,item in enumerate(a['共有商品'])]

a = a[a['共有商品']!='']
a.to_excel('139#' + time.strftime('%Y%m%d', time.localtime(time.time())) + '最新库存表.xlsx', index=False)

# 匹配库存
a['去格式货号&尺码大写配对列'] = [each_term.upper() for each_term in a['去格式货号&尺码']]
a.drop_duplicates('去格式货号&尺码大写配对列', keep='first', inplace=True)
for list_name in [gi3498, gi4083]:
    list_name['去格式货号&尺码大写配对列'] = [each_term.upper() for each_term in list_name['去格式货号&尺码']]
    list_name['最新库存'] = list_name.merge(a, on='去格式货号&尺码大写配对列', how='left')['可用量'].fillna('0')  # 缺失值填0
    del list_name['去格式货号&尺码大写配对列']

# 生成后台库存更新表
def import_storage_updata(which_list):
    dataframe = pd.DataFrame()
    if which_list=='3498':
        commodity_info_list = gi3498
    elif which_list=='4083':
        commodity_info_list = gi4083
    commodity_info_list = commodity_info_list[commodity_info_list['商家ID']=='139']
    dataframe['商品编码'] = commodity_info_list['商品编码']
    dataframe['库存'] = commodity_info_list['最新库存'].astype(float)
    num = int(np.ceil(len(dataframe)/999))
    for n in range(num):
        dataframe[n*999:(n+1)*999].to_excel(which_list + '后台库存更新汇总'+str(n+1)+ '.xlsx', index=False, float_format='%.0f')

toltal_list = ['3498', '4083']
for each_list in toltal_list:
    import_storage_updata(each_list)

# 删除多余表格
os.remove('现存量查询.xls')







