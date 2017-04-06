# 高德地图：http://ditu.amap.com/  高德地图poi：http://lbs.amap.com/api/webservice/guide/api/search/#text
# coding:utf-8
# github:https://github.com/tianyu8969/python-to-amap

import json
import xlwt
from datetime import datetime
from urllib import request
from urllib.parse import quote
import sys
import time

# 获取当前日期
today = datetime.today()
# 将获取到的datetime对象仅取日期如：2017-4-6
today_date = datetime.date(today)

json_name = 'data_amap.json'
# 高德地图poi：http://lbs.amap.com/api/webservice/guide/api/search/#text
# 请替换为自己申请的key值：申请Web服务API类型KEY http://lbs.amap.com/dev/
# 090000医疗
url_amap = 'http://restapi.amap.com/v3/place/text?key=6159ef91602ee2dbd718fc7c30601397&keywords=卫生服务中心&types=090000&city=上海&citylimit=true&children=1&offset=20&page=pageindex&extensions=all'
page_size = 20  # 每页记录数据，强烈建议不超过25，若超过25可能造成访问报错
page_index = r'page=1'  # 显示页码
total_record = 0  # 定义全局变量，总行数
# Excel表头
hkeys = ['id', '行业类型', '医院名称', '医院类型', '医院地址', '联系电话', 'location', '省份代码', '省份名称', '城市代码', '城市名称', '区域代码', '区域名称',
         '所在商圈']
# 获取数据列
bkeys = ['id', 'biz_type', 'name', 'type', 'address', 'tel', 'location', 'pcode', 'pname', 'citycode', 'cityname',
         'adcode', 'adname', 'business_area']


# 获取数据
def get_data(pageindex):
    global total_record
    # 暂停500毫秒，防止过快取不到数据
    time.sleep(0.5)
    print('解析页码： ' + str(pageindex) + ' ... ...')
    url = url_amap.replace('pageindex', str(pageindex))
    # 中文编码
    url = quote(url, safe='/:?&=')
    html = ""
    with request.urlopen(url) as f:
        html = f.read()
    rr = json.loads(html)
    if total_record == 0:
        total_record = int(rr['count'])
    return rr['pois']


def getPOIdata():
    global total_record
    print('获取POI数据开始')
    josn_data = get_data(1)
    if (total_record % page_size) != 0:
        page_number = int(total_record / page_size) + 2
    else:
        page_number = int(total_record / page_size) + 1

    with open(json_name, 'w') as f:
        # 去除最后]
        f.write(json.dumps(josn_data).rstrip(']'))
        for each_page in range(2, page_number):
            html = json.dumps(get_data(each_page)).lstrip('[').rstrip(']')
            if html:
                html = "," + html
            f.write(html)
            print('已保存到json文件：' + json_name)
        f.write(']')
    print('获取POI数据结束')


# 写入数据到excel
def write_data_to_excel(name):
    # 从文件中读取数据
    fp = open(json_name, 'r')
    result = json.loads(fp.read())
    # 实例化一个Workbook()对象(即excel文件)
    wbk = xlwt.Workbook()
    # 新建一个名为Sheet1的excel sheet。此处的cell_overwrite_ok =True是为了能对同一个单元格重复操作。
    sheet = wbk.add_sheet('Sheet1', cell_overwrite_ok=True)

    # 创建表头
    # for循环访问并获取数组下标enumerate函数
    for index, hkey in enumerate(hkeys):
        sheet.write(0, index, hkey)

    # 遍历result中的没个元素。
    for i in range(len(result)):
        values = result[i]
        n = i + 1
        for index, key in enumerate(bkeys):
            val = ""
            # 判断是否存在属性key
            if key in values.keys():
                val = values[key]
            sheet.write(n, index, val)
    wbk.save(name + str(today_date) + '.xls')
    print('保存到excel文件： ' + name + str(today_date) + '.xls ！')


if __name__ == '__main__':
    # 写入数据到json文件，第二次运行可注释
    getPOIdata()
    # 读取json文件数据写入到excel
    write_data_to_excel("上海卫生服务中心-高德地图")
