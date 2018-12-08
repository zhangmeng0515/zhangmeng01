import requests
import math
import pandas as pd
import time

def get_json(url, num):  # 获取JSON，获取第几页数据
    headers = {
            'User-Agent':'Mozilla/5.0(windows NT 10.0; Win64; x64) AppleWebKit/537.36(KHTML,like Gecko) Chorme/63.0.3239.132 Safari/537.36',
            'Host':'www.lagou.com',
            'Referer':'https://www.lagou.com/jobs/list_%E6%95%B0%E6%8D%AE%E5%88%86%E6%9E%90?city=%E6%B7%B1%E5%9C%B3&cl=false&fromSearch=true&labelWords=&suginput=&isSchoolJob=1'
            }   # 请求头，URL中设置了城市名
    form_data = {'first':'true','pn':num,'kd':'数据分析'}  # 网页From Data参数
    try:
        r=requests.post(url, headers=headers, data=form_data)  # 抓取职位信息的JOIN文件
        r.raise_for_status()
        r.encoding = 'utf-8'
        return r.json()
    except:
        return ''

def get_a_page(result):  # 获取一页的信息
    field = ['companyFullName','companyShortName','createTime','companySize','financeStage','district','positionName',
             'firstType','secondType','thirdType','industryField','positionLables','skillLables','jobNature',
             'resumeProcessDay','workYear','education','salary','positionAdvantage']  # 要获取哪些字段信息
    page_info = []
    try:
        for i in result:
            info = []
            [info.append(i[j]) for j in field]
            page_info.append(info)
        return page_info
    except:
        return ''
    
def get_pages(count):
    num = math.ceil(count/15)
    if num >= 100:
        return 100
    else:
        return num
    
def main():
    url = 'https://www.lagou.com/jobs/positionAjax.json?city=%E6%B7%B1%E5%9C%B3&needAddtionalResult=false'
    page_1 = get_json(url, 1)  # 总职位数在每页的json字典中都有显示，这里随便选了第一页
    total_count = page_1['content']['positionResult']['totalCount'] 
    num=get_pages(total_count)  
    total_info=[]
    print('职位总数:{},页数:{}'.format(total_count, num))
    start= time.time()
    time.sleep(20)
    for n in range(1, num+1):  # 获取每页的json,汇总数据，转成一个嵌套列表,每个子列表为一个职位的所有信息
        page=get_json(url, n)
        result = page['content']['positionResult']['result']
        a_page_info = get_a_page(result)
        total_info += a_page_info
        print('已经抓取第{}页,职位总数:{}，当前总耗时{:.2f}秒'.format(n, str(len(total_info)), time.time()-start))
        time.sleep(10)# 暂停10秒，防止被服务器拉黑
    # 把list转为DataFrame
    df=pd.DataFrame(data=total_info,columns=['公司全名','公司简称','创办时间','公司规模','融资阶段','区域','职位名称',
                                             '岗位类型1级','岗位类型2级','岗位类型3级','行业领域','职位标签','技能标签','工作形式',
                                             '简历处理时间','工作经验','学历要求','工资','职位福利'])
    df.to_csv(r'C:\Users\a\AppData\Local\Programs\Python\Python37\lagou_jobs2323.csv',index=False)
    print('CSV保存成功')

if __name__=='__main__': #在其他文件import这个py文件时,不会自动运行主函数
    main()






