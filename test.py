import os
import arcpy
import requests
import json
from urllib.parse import quote
import xlwt
from Coordin_transformlat import gcj02towgs84


def get_poi(key, city, keywords, page):
    """
    这是一个能够从高德地图获取poi数据的函数
    key：为用户申请的高德密钥
    city：目标城市
    types：POI数据的类型
    page：当前页数
    """

    # 设置header
    header = {
        'User-Agent': "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50"}

    # 构建url
    # {}在链接里表示占位符，也就是占住位置先不填写，.format()里就是往刚刚占位符的地方填写变量，按照顺序一一对应，第一个{}里是key，第二个{}里是keywords
    url = 'https://restapi.amap.com/v3/place/text?key={}&keywords={}&city={}&page={}&output=josn'.format(key, keywords,
                                                                                                         quote(city),
                                                                                                         page)

    # 用get函数请求数据
    r = requests.get(url, headers=header)
    # 设置数据的编码为'utf-8'
    r.encoding = 'utf-8'
    # 将请求得到的数据按照'utf-8'编码成字符串
    data = r.text
    return data


def get_times(key, city, keywords):
    """
    这是一个控制申请次数的函数
    """
    page = 1
    # 创建一个poilist的空列表
    poilist = []
    # 执行以下代码，直到count为0的时候跳出循环
    while True:
        # 调用第一个函数来获取数据
        result = get_poi(key, city, keywords, page)

        # json.loads可以对获取回来JSON格式的数据进行解码
        content = json.loads(result)

        # content的样子其实跟返回结果参数是一样的，可以想象成有很多个字段的excel表格，下面这个语句就是提取出pois那个字段
        pois = content['pois']

        # pois的信息写入空列表里，这里由于不知道返回的数据长什么样子，所以会难以理解些
        for i in range(len(pois)):
            poilist.append(pois[i])

        # 递增page
        page = page + 1

        # 判断当前页下的count是否等于0
        if content['count'] == '0':
            break
    # 将写好poi信息的列表返回
    return poilist


def write_to_excel(poilist, city, keywords):
    """
    这是一个可以将列表写入excel的函数
    poilist -- 上面得到的poilist
    city -- 城市名，这个参数是保存excel文件的名字用的
    # types -- poi类别，这个也是为了保存excel文件，可直接看代码最后一行
    keywords -- poi类别，这个也是为了保存excel文件，可直接看代码最后一行
    """
    # 我们可以把这两行代码理解成新建一个excel表，第一句是新建excel文件
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    # 往这个excel文件新建一个sheet表格
    sheet = book.add_sheet(keywords, cell_overwrite_ok=True)

    # 第一行(列标题)
    sheet.write(0, 0, 'x')
    sheet.write(0, 1, 'y')
    sheet.write(0, 2, 'count')
    sheet.write(0, 3, 'name')
    sheet.write(0, 4, 'address')
    sheet.write(0, 5, 'adname')

    # 最难理解的地方应该是这里了，放到代码后面讲解
    for i in range(len(poilist)):
        name = poilist[i]['name']
        location = poilist[i]['location']
        address = poilist[i]['address']
        adname = poilist[i]['adname']
        lng = str(location).split(",")[0]
        lat = str(location).split(",")[1]

        # 这里是坐标系转换，也放到代码后面详解
        result = gcj02towgs84(location)
        lng = result[0]
        lat = result[1]

        # 每一行写入
        sheet.write(i + 1, 0, lng)
        sheet.write(i + 1, 1, lat)
        sheet.write(i + 1, 2, 1)
        sheet.write(i + 1, 3, name)
        sheet.write(i + 1, 4, address)
        sheet.write(i + 1, 5, adname)
    # 最后，将以上操作保存到指定的Excel文件中
    book.save(os.path.join(r'C:\Users\Administrator\Desktop\选修课作业\pois', city + "_" + keywords + '.xls'))


def txt_to_shp(file, fc, fc_type):
    """支持画多线"""
    with open(file, "r") as f:
        with arcpy.da.InsertCursor(fc, ["SHAPE@"]) as cur:
            if fc_type == "POINT":
                for line in f.readlines():
                    if len(line.split(" ")) == 3:
                        point = arcpy.Point()
                        point.X = line.split(" ")[-2]
                        point.Y = line.split(" ")[-1]
                        point_Germetry = arcpy.PointGeometry(point)
                        cur.insertRow([point_Germetry])


if __name__ == '__main__':
    # 这里修改为自己的高德密钥
    key = '4571a8bb06a2632eadc39d93cf5d4951'

    # 这里修改自己的poi类型
    types = ['核酸采样点']

    # 建议将大区域切分成几个小区域来获取，保证获取的数据齐全
    city_list = ['二七区', '中原区', '管城区', '金水区', '惠济区']
    # city_list = ['郑东新区', '经开区', '高新区', '航空港区']

    # 先遍历city_list里面的每个区域
    for city in city_list:
        # 再遍历types里的每个类别
        for type in types:
            poi = get_times(key, city, type)
            print('当前城市：' + str(city) + ', 分类：' + str(type) + ", 总的有" + str(len(poi)) + "条数据")
            write_to_excel(poi, city, type)
            print('*' * 50 + '分类：' + str(type) + "写入成功" + '*' * 50)

    # # 这样搜索只能得到200行POI数据
    # key = 'c27e7fcfb55e45d68aff120c05bbc62e'
    # city = '郑州市'
    # keywords = ['二七区核酸采样点', '中原区核酸采样点', '管城区核酸采样点', '金水区核酸采样点', '上街区核酸采样点', '惠济区核酸采样点']
    # poi = get_times(key, city, keywords)
    # for keyword in keywords:
    #     print('当前城市：' + str(city) + str(keyword) + ", 总的有" + str(len(poi)) + "条数据")
    #     write_to_excel(poi, city, keyword)
    #     print('*' * 50 + '分类：' + str(keyword) + "写入成功" + '*' * 50)
    # print('====爬取完成====')

