import pandas as pd
import re
import os
import time

os.chdir(r'E:\0000~急需处理')  
def get_col_list_after_insert(dataframe, insert_before_which_col, the_seq_to_insert):
    col_name = dataframe.columns.tolist()
    for each_field in the_seq_to_insert:
        col_name.insert(col_name.index(insert_before_which_col), each_field)
    return col_name

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

def remove_merchant_and_format(remove_format_list, merchant_id_list):
    temp_list = []
    for num in range(len(merchant_id_list)):
        n = len(remove_format_list[num])-len(merchant_id_list[num])
        goods_num = remove_format_list[num][:n]
        temp_list.append(goods_num)
    return temp_list

def clean_specification_and_size(source_col):
    temp_list = []
    find_num = re.compile('[0-9\.]+')
    split_by_space = re.compile(' ')
    for each_term in source_col:
        match = split_by_space.split(each_term)
        if len(match)==1:
            if match[0] in ['S','M','L','XS','XXS','XL','XXL','XXXL']:
                temp_list.append(match[0])
            elif find_num.search(match[0]):
                temp_list.append(find_num.search(match[0]).group(0))
            else:   
                temp_list.append('')
        else:
            chima = match[-1]
            if 'cm' in chima or '寸' in chima or '号' in chima:
                chima = chima.replace('cm', '')
                chima = chima.replace('寸', '')
                chima = chima.replace('号', '')
                temp_list.append(chima)
            elif chima in ['7P', '4A', 'BKT', '其它']:
                temp_list.append('')
            elif chima[-1] in ['色', '纹','形','花','形','粉', '蓝', '彩', '节','系','红','金','黑','白','灰','棕','橙','黄','绿','兰','紫','子']:
                temp_list.append('')
            elif '码' in chima:
                if chima[-1]=='码':  
                    temp_list.append('')
                else:
                    n = len(chima) - chima.find('码')-1
                    temp_list.append(chima[-n:])  
            elif chima=='00':
                temp_list.append('0')
            else:
                temp_list.append(chima)
    return temp_list

def concat_two_col(first_col, second_col):
    temp_list = []
    for num in range(len(first_col)):
        temp_list.append(first_col[num] + second_col[num])
    return temp_list

def concat_third_col(first_col, second_col, third_col):
    temp_list = []
    for num in range(len(first_col)):
        temp_list.append(first_col[num] + second_col[num] +third_col[num])
    return temp_list
    
def main():
    toltal_storage_table = pd.read_excel('最新实库+虚库成本表汇总190513(146).xlsx', sheet_name='Sheet1')
    toltal_storage_table['去格式货号&尺码大写配对列'] = [each_term.upper() for each_term in toltal_storage_table['去格式货号&尺码']]
    toltal_storage_table.drop_duplicates('去格式货号&尺码大写配对列', keep='first', inplace=True)
    for list_name in ['3498','4083']:
        # 读取文件
        full_name = list_name +'.xlsx'
        ysf = pd.read_excel(full_name, sheet_name='Sheet0')
        
        # 定义新的列名
        new_col = ['尺码','去格式货号','去商家去格式货号','货号去#','商家ID','商编','结算价1','上架状态','去商家去格式货号与尺码','去格式货号&尺码','与最新欧派数据匹配库存']
        col_name = get_col_list_after_insert(ysf, '商品编码', new_col)
        col_name.append('最新库存')
        ysf = ysf.reindex(columns=col_name)
        
        # 格式规整处理
        ysf = ysf.astype(str).replace('nan','')
      
        # 货号去#列处理
        ysf['货号去#'] = remove_hashtag(ysf['货号'])

        # 去格式货号列处理
        ysf['去格式货号']= remove_format(ysf['货号去#'])
           
        # 获取商家ID
        ysf['商家ID'] = get_merchant_id(ysf['货号去#'])
        
        # 获取去商家去格式货号
        ysf['去商家去格式货号'] = remove_merchant_and_format(ysf['去格式货号'], ysf['商家ID'])        
        
        # 清洗尺码
        ysf['尺码'] = clean_specification_and_size(ysf['规格'])

        # 其他字段简单处理
        ysf['商编'] = ysf['商品编码']
        ysf['结算价1'] = ysf['结算价'].apply(lambda x:int(eval(x))).astype(str)
        ysf['上架状态'] = ysf['销售状态']
        ysf['去商家去格式货号与尺码'] = concat_two_col(ysf['去商家去格式货号'], ysf['尺码'])
        ysf['去格式货号&尺码'] = concat_two_col(ysf['去格式货号'], ysf['尺码'])
        
        # 匹配最新库存
        '''这里为数值格式，注意~~~~~~~~~~~~~~！！！！！！！！！！'''
        ysf['去格式货号&尺码大写配对列'] = [each_term.upper() for each_term in ysf['去格式货号&尺码']]
        ysf['最新库存'] = ysf.merge(toltal_storage_table, on='去格式货号&尺码大写配对列', how='left')['上架库存'].fillna(0)
        del ysf['去格式货号&尺码大写配对列']
        
        # 保存文件
        ysf.to_excel(list_name + '#' + time.strftime('%Y%m%d',time.localtime(time.time())) + '商品信息表.xlsx', index=False)

if __name__=='__main__':                  
    main()
