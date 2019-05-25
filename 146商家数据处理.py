import os
import pandas as pd
from openpyxl import load_workbook
import time
import re
import sys
sys.path.append(r'E:\0210~处理代码\代码')
import goodslist

start = time.perf_counter()
os.chdir(r'E:\0210~处理代码')


def excelAddSheet(dataframe, excelpath, Sheet_name='AddSheet'):
   excelWriter=pd.ExcelWriter(excelpath,engine='openpyxl')
   book = load_workbook(excelWriter.path)
   excelWriter.book = book
   dataframe.to_excel(excel_writer=excelWriter, sheet_name=Sheet_name, index=False)
   excelWriter.close()

def deal_original_data(excelpath, source_col):
    whether_pic_exist = []
    for i in source_col:
        if i =='':
            whether_pic_exist.append('0')
        else:
            whether_pic_exist.append('1')
    return whether_pic_exist

def get_brand_list(source_col):
    find_chinese_str = re.compile('[\u4e00-\u9fa5]')
    remove_redundant_space = re.compile(' +')
    brand_list = []
    for i in source_col:
        match = find_chinese_str.search(i)
        if match:
            temp1 = i[0:match.start()].replace('amp;', '').replace('&', ' ')
            temp2 = remove_redundant_space.sub(' ',temp1)
            brand_list.append(temp2.strip())
        else:
            temp3 = i.replace('amp;', '').replace('&', ' ')
            temp4 = temp3.replace('DELUXE BRAND','').strip()
            temp5 = remove_redundant_space.sub(' ',temp4)
            brand_list.append(temp5)
    return brand_list

def remove_merchant_and_format_remove_o(source_col):
    fuhao1 = re.compile('-')
    fuhao2 = re.compile('\/')
    fuhao3 = re.compile('\.')
    fuhao4 = re.compile(' ')
    fuhao5 = re.compile('\*')
    fuhao6 = re.compile('O')
    fuhao7 = re.compile('o')
    remove_format_and_merchant = []
    for i in source_col:
        c_step1 = fuhao1.sub('', i)
        c_step2 = fuhao2.sub('', c_step1)
        c_step3 = fuhao3.sub('', c_step2)
        c_step4 = fuhao4.sub('', c_step3)
        c_step5 = fuhao5.sub('', c_step4)
        c_step6 = fuhao6.sub('0', c_step5)
        c_step7 = fuhao7.sub('0', c_step6)
        remove_format_and_merchant.append(c_step7)
    return remove_format_and_merchant 

def round_to_int(x):
    if x-int(x)>=0.5:
        return int(x)+1
    else:
        return int(x)

def get_size_list(source_col):
    specification_attributes = []
    re_clean_size = re.compile('[0-9XMLSxmls][\w\*\.\/]*')
    for i in source_col:
        match = re_clean_size.search(i)
        if match:
            temp1 = match.group(0).replace('cm', '')
            if temp1[0]=='0'and len(temp1)>=2:
                specification_attributes.append(temp1[1:])
            else:
                specification_attributes.append(temp1)
        else:
            specification_attributes.append('')
    return specification_attributes

def get_sell_storage_list(source_col):
    storage_for_sell = []
    for i in source_col:
        if eval(i) > 0:
            storage_for_sell.append(str(eval(i)-1))
        else:
            storage_for_sell.append('0')
    return storage_for_sell

def get_whether_storage_change_list(number_col):
    whether_add_storage = []
    for num in number_col:
        if concat_table['匹配列'][num]=='0' and concat_table['上架库存_y'][num]=='0' and eval(concat_table['上架库存_x'][num])>0:
            whether_add_storage.append('1')
        else:
            whether_add_storage.append('0')
    return whether_add_storage

def whether_sale_in_account(account):
    temp_list = []
    Account = pd.read_excel(account +'#' + time.strftime('%Y%m%d',time.localtime(time.time())) + '商品信息表.xlsx')
    Account = Account.astype(str).replace('nan', '')
    Account.drop_duplicates('去格式货号&尺码', keep='first', inplace=True)
    Account['匹配列'] = '0'
    temp_table = add_sheet.merge(Account, on='去格式货号&尺码', how='left')
    temp_table['匹配列'] = temp_table['匹配列'].fillna('1')
    for num in range(len(temp_table['匹配列'])):
        if temp_table['匹配列'][num]=='0' and temp_table['销售状态'][num]=='上架可销售':
            temp_list.append('1')
        else:
            temp_list.append('0')
    return temp_list

def whether_cost_change(number_col):
    cost_change = []
    for num in number_col:
        if concat_table['匹配列'][num]=='0' and concat_table['成本价（HK)_x'][num]!=concat_table['成本价（HK)_y'][num]:
            cost_change.append('1')
        else:
            cost_change.append('0')
    return cost_change

def whether_make_list_and_its_raw_goodsnum(number_col):
    make_list = []
    raw_goodsnum = []
    for num in number_col:
        if add_sheet['成本价是否变动'][num] == '1' or ((eval(add_sheet['是否货号新增'][num]) + 
                    eval(add_sheet['是否上次库存为0此次库存大于0'][num]))>0 and (add_sheet['100180是否在售'][num]==
                    add_sheet['3498是否在售'][num]==add_sheet['100247是否在售'][num]==add_sheet['4083是否在售'][num]=='0'
                    and eval(add_sheet['上架库存'][num])>0)):
            make_list.append('1')
            raw_goodsnum.append(add_sheet['原始货号'][num])
        else:
            make_list.append('0')
            raw_goodsnum.append('')
    return (make_list, raw_goodsnum)

def final_list():
    make_final_list = []
    final_raw_goodsnum = []
    for num in range(len(add_sheet['上架清单的原始货号'])):
        if add_sheet['原始货号'][num] in list(add_sheet['上架清单的原始货号']):
            raw_goodsnum_and_size = add_sheet['原始货号'][num] + add_sheet['尺码'][num]
            make_final_list.append(add_sheet['原始货号'][num])
            final_raw_goodsnum.append(raw_goodsnum_and_size)
        else:
            make_final_list.append('')
            final_raw_goodsnum.append('')
    return (make_final_list, final_raw_goodsnum)

def concat_detail_pic_path(reference_col):
    detail_pic_path = []
    for num in range(len(reference_col)):
        if reference_col[num]=='':
            detail_pic_path.append('')
        else:
            detail_pic_path.append(add_sheet['图片路径(imgUrl)'][num] + reference_col[num])
    return detail_pic_path



old_table = '20190111-欧派数据.xlsx'
new_table = '20190114-欧派数据原数据.xlsx'
# 读取今日库存表，处理格式和缺失值，添加字段，写到到该表的一个新sheet中
newest_storage_list = pd.read_excel(new_table)
newest_storage_list = newest_storage_list.astype(str).replace('nan','')
newest_storage_list['是否有图'] = deal_original_data(new_table, newest_storage_list['货品主图(goodsImage)'])
excelAddSheet(newest_storage_list, new_table, Sheet_name='总数据')

# 筛选符合要求的数据，添加需要的字段列，处理列的格式和缺失值
add_sheet = newest_storage_list[(newest_storage_list['是否删除（0：未删除，1：删除）(isDel)']=='0')&
                            (newest_storage_list['是否上架（1：上架0：下架）(goodsShow)']!='0')&
                            (newest_storage_list['规格是否删除（0：未删除，1：删除）(isDel)']=='0')&
                            (newest_storage_list['规格否上架（1：上架0：下架）(specIsOpen)']=='1')&
                            (newest_storage_list['香港自提价(specGoodsPrice)']!='0')&
                            (newest_storage_list['库存'].astype(int)>0)&
                            (newest_storage_list['是否有图']=='1')]
final_col_name = add_sheet.columns.tolist()
temp_col_name = ['归属表', '商家', '品牌', '品类', '原始货号', '尺码', '去格式去商家货号', '去格式货号', '去格式去商家货号&尺码',
                  '去格式货号&尺码', '去格式去商家货号&尺码（运营专用）', '成本价（HK)', '库存1', '上架库存', '是否货号新增',
            '是否上次库存为0此次库存大于0', '100180是否在售', '3498是否在售', '100247是否在售', '4083是否在售', '成本价是否变动',
            '是否制作上架清单', '上架清单的原始货号', '最终上架清单(原始货号)', '最终上架清单(货号+尺码)','图片地址', '图片保存文件夹',
            '详情图1地址', '详情图2地址', '详情图3地址', '详情图4地址', '图片分列1', '图片分列2','图片分列3', '图片分列4']
final_col_name.extend(temp_col_name)
add_sheet = add_sheet.reindex(columns = final_col_name,fill_value='') 

# 处理列：归属表~上架库存
add_sheet['图片路径(imgUrl)']=[i.replace('img/','img') for i in add_sheet['图片路径(imgUrl)']]
add_sheet['归属表'] = add_sheet['归属表'].apply(lambda x:time.strftime('%Y-%m-%d', time.localtime()) + '更新ZM')
add_sheet['品类'] = add_sheet['一级分类名称(gcName2)']
add_sheet['商家'] = add_sheet['商家'].apply(lambda x:'146')
add_sheet['品牌'] = get_brand_list(add_sheet['货品品牌名称(brandName)'])
add_sheet['原始货号'] = add_sheet['商品货号(goodsSerial)']
add_sheet['去格式去商家货号'] = remove_merchant_and_format_remove_o(add_sheet['原始货号'])
add_sheet['去格式货号'] = Goodsinfo.concat_two_col(add_sheet['去格式去商家货号'], add_sheet['商家'])
add_sheet['成本价（HK)'] = (add_sheet['香港自提价(specGoodsPrice)'].astype(float)/0.88).apply(lambda x:round_to_int(x)).astype(str)
add_sheet['库存1'] = add_sheet['库存']
add_sheet['尺码'] = get_size_list(add_sheet['规格属性值(specGoodsSpec)'])
add_sheet.reset_index(drop=True, inplace=True)
add_sheet['去格式去商家货号&尺码'] = Goodsinfo.concat_two_col(add_sheet['去格式去商家货号'], add_sheet['尺码'])
add_sheet['去格式货号&尺码'] = Goodsinfo.concat_third_col(add_sheet['去格式去商家货号'], add_sheet['商家'], add_sheet['尺码'])
add_sheet['上架库存'] = get_sell_storage_list(add_sheet['库存1'])

# 处理列：是否货号新增~成本价是否变动
last_storage_list = pd.read_excel(old_table, sheet_name = '剔除不上架无图无库存无成本价')
last_storage_list = last_storage_list.astype(str).replace('nan', '')
last_storage_list.drop_duplicates('去格式货号&尺码', keep='first', inplace=True)
last_storage_list['匹配列'] = '0'
concat_table = add_sheet.merge(last_storage_list, on='去格式货号&尺码', how='left')
add_sheet['是否货号新增'] = concat_table['匹配列'] = concat_table['匹配列'].fillna('1')
add_sheet['是否上次库存为0此次库存大于0'] = get_whether_storage_change_list(range(len(concat_table['匹配列'])))
add_sheet['100180是否在售'] = whether_sale_in_account('100180')
add_sheet['3498是否在售'] = whether_sale_in_account('3498')
add_sheet['100247是否在售'] = whether_sale_in_account('100247')
add_sheet['4083是否在售'] = whether_sale_in_account('4083')
add_sheet['成本价是否变动'] = whether_cost_change(range(len(concat_table['匹配列'])))
(add_sheet['是否制作上架清单'],add_sheet['上架清单的原始货号']) = whether_make_list_and_its_raw_goodsnum(range(len(add_sheet['上架库存'])))
(add_sheet['最终上架清单(原始货号)'], add_sheet['最终上架清单(货号+尺码)']) = final_list()

# 处理列：图片地址~最后
add_sheet['图片地址'] = Goodsinfo.concat_two_col(add_sheet['图片路径(imgUrl)'], add_sheet['货品主图(goodsImage)'])
'''注意：分列后空值显示为Nonetype~~~~~~~~~~变成字符串None！！！！！！！！'''
separate_pic_path = add_sheet['货品多图(goodsImageMore)'].str.split(',', expand=True).astype(str).replace('None','')
add_sheet['图片分列1'] = separate_pic_path[1] 
add_sheet['图片分列2'] = separate_pic_path[2]
add_sheet['图片分列3'] = separate_pic_path[3]
add_sheet['图片分列4'] = separate_pic_path[4]
add_sheet['详情图1地址'] = concat_detail_pic_path(add_sheet['图片分列1'])
add_sheet['详情图2地址'] = concat_detail_pic_path(add_sheet['图片分列2'])
add_sheet['详情图3地址'] = concat_detail_pic_path(add_sheet['图片分列3'])           
add_sheet['详情图4地址'] = concat_detail_pic_path(add_sheet['图片分列4'])           


# 保存到新sheet下
excelAddSheet(add_sheet, new_table, Sheet_name='剔除不上架无图无库存无成本价')
end = time.perf_counter()
print ('欧派表处理时间:{:.2f}秒'.format(end-start))
    





# =============================================================================
# 上架清单部分
# =============================================================================
os.chdir(r'E:\0210~处理代码')
import urllib.request
start = time.perf_counter()
def download_picture(path, url_list, name_list):
    if not os.path.exists(path):
        os.makedirs(path)
        print(path + '——文件夹创建成功')
    else:
        print(path + '——目录已存在')
    for num in range(len(url_list)):
        try:
            urllib.request.urlretrieve(url_list[num], path +'/' + name_list[num] + '.jpg')
            print('下载成功')
        except:
            print('下载错误')
            
sale_table = add_sheet[add_sheet['最终上架清单(货号+尺码)']!=''] # 筛选赋值后注意索引需要
sale_table.reset_index(drop=True, inplace=True)
download_picture('图片', sale_table['图片地址'], sale_table['原始货号'])
    
end = time.perf_counter()
print ('图片下载时间:{:.2f}秒'.format(end-start))  
    