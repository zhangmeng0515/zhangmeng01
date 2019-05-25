import pandas as pd
from selenium import webdriver
import time
import os
import numpy as np
import shutil

os.chdir(r'E:\0000~急需处理')

class download_goodsinfo(object):
    def __init__(self, account, password):      
        self.option = webdriver.ChromeOptions()  
        # 去掉Chrome 正受到自动测试软件的控制
        self.option.add_argument('disable-infobars')
        # 创建Chrome驱动器对象
        self.chromedriver = '{}'.format(os.path.abspath('chromedriver.exe'))  # 需要先下载一个与谷歌浏览器版本匹配的chromedriver
        self.driver = webdriver.Chrome(self.chromedriver, chrome_options=self.option)
        
        
     # 1. 登陆页面
    def login(self, account, password):
        self.driver.maximize_window()  # 窗口最大化
        time.sleep(0.5) 
        self.driver.get('https://abdpop.secoo.com/login?redirectUrl=')  # 打开登陆页面链接
        self.driver.find_element_by_xpath('//input[@placeholder="Merchant ID"]').send_keys(account)  # 定位账号输入局域网，输入账号
        time.sleep(0.5)
        self.driver.find_element_by_xpath('//input[@placeholder="password"]').send_keys(password)  # 定位密码输入局域网，输入密码
        time.sleep(40)
        
     # 2. 导出和下载表格 
    def download_table(self, account):
        self.driver.get('https://abdpop.secoo.com/productInfo/spuManageList')  # 登陆成功后跳转页面
        time.sleep(5)
        self.driver.find_element_by_xpath("//a[@ng-click='exportSkuList()']").click()  # 定位导出商品按钮，点击
        time.sleep(10)
        number = str(int(np.random.rand()*100000000))  # 设置一个随机数
        self.driver.find_element_by_xpath("//input[@id='remarkVal']").send_keys(number)  # 定位输入框，输入随机数
        time.sleep(1)
        self.driver.find_element_by_xpath("//div[@class='pm_cont_btn']/a[1]").click()  # 点击确定
        self.driver.get('https://abdpop.secoo.com:443/vendorFile/toVendorFilePage')  # 跳转到下载页面
        if account == '3498':
            time.sleep(200)
        else:
            time.sleep(20)
        self.driver.refresh()  # 刷新页面
        time.sleep(10)
        self.filename = self.driver.find_element_by_xpath("//tbody/tr[2]/td[1]").text  # 获取下载文件的原始名称
        # if driver.find_element_by_xpath("//tbody/tr[2]/td[3]").text == number:
        self.driver.find_element_by_xpath("//tbody/tr[2]/td[6]/a[1]".format(number)).click()  # 点击下载
        time.sleep(60)
        self.driver.quit()
     
     # 3. 移动和重命名下载的文件
    def move_and_rename(self, account):
        shutil.move('C:\\Users\\Administrator\\Downloads\\' + self.filename, 'E:\\0000~急需处理\\' + account + '.xlsx')
    
     # 运行方法
    def run(self, account, password):
        self.login(account, password)
        self.download_table(account)
        self.move_and_rename(account)


# password = {'4083':'LlFm2259','3498':'Fmll2260'}
# accounts = list(password.keys())

if __name__ == '__main__':
        temp = download_goodsinfo('4083', 'Ll456852')
        temp.run('4083', 'Ll456852')
        temp = download_goodsinfo('3498', 'Fm159753')
        temp.run('3498', 'Fm159753')
        
# 出现chrome not reachable, 运行太多次，需要关闭掉之前的chromedriver
# os.rename 无法跨磁盘， 改用shutil.move
    
#  提审表自动化测试
filename = '【提审表】146#190410-3498-陈'
costname = '最新实库+虚库成本表汇总190415(实库)'
table = pd.read_excel('E:/0170~腾讯QQ接收及下载文件/' + filename + '.xlsx')
table = table.dropna(how='all').reset_index(drop=True)
def get_col_list_after_insert(dataframe, insert_before_which_col, the_seq_to_insert):
    col_name = dataframe.columns.tolist()
    for each_field in the_seq_to_insert:
        col_name.insert(col_name.index(insert_before_which_col), each_field)
    return col_name
col_seq = get_col_list_after_insert(table,'型号',['货号去#', '商家', '去格式货号', '货号去格式去商家'])
col_seq.extend(['账户', '最新销售状态', '最新结算价', '最新市场价', '最新匹配成本价', '最新运费', '最新毛利', '最新毛利率', '最新情况说明',
                '发布库存(不是实际库存)'])
table = table.reindex(columns=col_seq)
table['情况说明'] = table['情况说明'].fillna('').astype(str)
table['账户'] = table['账户'].fillna('')
for num,i in enumerate(table['情况说明']):
    if '3498' in i:
        table['账户'][num] = '3498'
    elif '4083' in i:
        table['账户'][num] = '4083'
table['账户'].replace('',method='ffill', inplace=True)

import re
def remove_hashtag(col_name):
    qu_jing = re.compile(r'#.*')
    qu_kuohao1 = re.compile('\(.*\)')
    qu_kuohao2 = re.compile('（.*）')
    qu_da_o = re.compile('O')
    qu_xiao_o = re.compile('o')
    temp_list = []
    for each_term in col_name:
        clean_step1 = qu_jing.sub('', each_term)
        if clean_step1 == '':
            temp_list.append('')
        else:
            clean_step2 = qu_kuohao1.sub('', clean_step1)
            clean_step3 = qu_kuohao2.sub('', clean_step2)
            clean_step4 = qu_da_o.sub('0', clean_step3)
            clean_step5 = qu_xiao_o.sub('0', clean_step4)
            temp_list.append(clean_step5)
    return temp_list
table['货号去#'] = remove_hashtag(table['货号'])
      
def get_merchant_id(source_col):
    temp_list = []
    for each_term in source_col:
        if '-' in each_term[-5:] and each_term[-1]!='-':
            n = 5-each_term[-5:].find('-')-1
            merchant_id = each_term[-n:]
            if n<1:
                temp_list.append('')
            elif n==1 and merchant_id in ['s','S']:
                temp_list.append('S')
            elif n==2:
                temp_list.append('')
            elif n==3 and merchant_id in ['001','002','003','004','005','006','007','008','009','010','012','013','015','016','020','022','025','034','036',
                                          '039','041','042','050','057','077','078','080','092','099','117','118','139','141','144','146','148']:
                temp_list.append(merchant_id)
            elif n==4 and merchant_id in ['025S', '025s', '012w', '012W']:
                temp_list.append(merchant_id.upper())            
            else:
                  temp_list.append('')
        else:
            temp_list.append('')
    return temp_list
table['商家'] = get_merchant_id(table['货号去#'])

def remove_format(col_name):
    fuhao1 = re.compile('-')
    fuhao2 = re.compile('\/')
    fuhao3 = re.compile('\.')
    fuhao4 = re.compile(' ')
    fuhao5 = re.compile('\*')
    temp_list = []
    for each_term in col_name:
        c_step1 = fuhao1.sub('', each_term)
        c_step2 = fuhao2.sub('', c_step1)
        c_step3 = fuhao3.sub('', c_step2)
        c_step4 = fuhao4.sub('', c_step3)
        c_step5 = fuhao5.sub('', c_step4)
        temp_list.append(c_step5)
    return temp_list
table['去格式货号'] = remove_format(table['货号去#'])

def remove_merchant_and_format(remove_format_list, merchant_id_list):
    temp_list = []
    for num in range(len(merchant_id_list)):
        n = len(remove_format_list[num])-len(merchant_id_list[num])
        goods_num = remove_format_list[num][:n]
        temp_list.append(goods_num)
    return temp_list
table['货号去格式去商家'] = remove_merchant_and_format(table['去格式货号'], table['商家']) 
    

gi3498 = pd.read_excel('3498.xlsx').fillna('').astype(str)
gi4083 = pd.read_excel('4083.xlsx').fillna('').astype(str)
gi3498.drop_duplicates('商品编码', keep='first', inplace=True)
gi4083.drop_duplicates('商品编码', keep='first', inplace=True)

def vlookup(find_value, start_col, object_col):
    try:
        num = list(start_col).index(find_value)
        return object_col[num]
    except:
        return ''
table['商品编码'] = table['商品编码'].fillna('').astype(str)
table['最新销售状态'] = [vlookup(code, gi3498['商品编码'], gi3498['销售状态']) if table['账户'][i]=='3498' else vlookup(code, gi4083['商品编码'], gi4083['销售状态']) 
                                                                                                                            for i,code in enumerate(table['商品编码'])]
table['最新结算价'] = [vlookup(code, gi3498['商品编码'], gi3498['结算价']) if table['账户'][num]=='3498' else vlookup(code, gi4083['商品编码'], gi4083['结算价']) 
                                                                                                                            for num,code in enumerate(table['商品编码'])]
table['最新市场价'] = [vlookup(code, gi3498['商品编码'], gi3498['市场价']) if table['账户'][num]=='3498' else vlookup(code, gi4083['商品编码'], gi4083['市场价']) 
                                                                                                                            for num,code in enumerate(table['商品编码'])]
# 设置index列为索引 table.set_index(['index'], inplace=True)
costinfo = pd.read_excel(costname + '.xlsx').fillna('').astype(str)
costinfo = costinfo[costinfo['商家'].isin(list(set(table['商家'])))]
costinfo1 = costinfo.sort_values(by=['去格式货号', '成本价（HK)'], ascending=(True, False))
costinfo1.drop_duplicates('去格式货号', keep='first', inplace=True)
table['最新匹配成本价'] = [vlookup(goodscode, costinfo['去格式货号'], costinfo['成本价（HK)']) for goodscode in table['去格式货号']]
table['最新运费'] = [200 if cost>2000 else 165 for cost in table['最新匹配成本价']]
table['最新毛利'] = [saleprice - table['最新匹配成本价'][i] - table['最新运费'][i] for i,saleprice in enumerate(table['结算价'])]
table['最新毛利率'] = ['{:.2f}'.format(np.round(table['最新毛利'][i] / saleprice, 2) * 100) for i,saleprice in enumerate(table['结算价'])]
costinfo['上架库存'] = costinfo['上架库存'].replace('', 0).astype(float)
costinfo2 = costinfo.groupby(by='去格式货号')['上架库存'].sum()
costinfo2 = costinfo2.reset_index()
table['发布库存(不是实际库存)'] = [vlookup(goodscode, costinfo2['去格式货号'], costinfo2['上架库存']) for goodscode in table['去格式货号']]
table.to_excel(filename + '已审.xlsx', index=False)




















    
    
    
    
    