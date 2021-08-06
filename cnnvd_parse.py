# -*- coding:utf-8 -*-
import importlib
import sys

# print (u'系统默认编码为',sys.getdefaultencoding())
default_encoding = 'utf-8'  # 重新设置编码方式为uft-8
if sys.getdefaultencoding() != default_encoding:
    importlib.reload(sys)
    sys.setdefaultencoding(default_encoding)
# print (u'系统默认编码为',sys.getdefaultencoding())
import requests
from bs4 import BeautifulSoup
import xlwt


def getURLDATA(url):
    # url = 'http://www.cnnvd.org.cn/web/xxk/ldxqById.tag?CNNVD=CNNVD-201901-1014'
    header = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.80 Safari/537.36',
        'Connection': 'keep-alive', }
    r = requests.get(url, headers=header, timeout=30)
    # r.raise_for_status()抛出异常
    html = BeautifulSoup(r.content.decode(), 'html.parser')

    link = html.find(class_='detail_xq w770')  # 漏洞信息详情
    link_introduce = html.find(class_='d_ldjj')  # 漏洞简介
    link_others = html.find_all(class_='d_ldjj m_t_20')  # 其他
    # print(len(link_introduce))
    try:
        #print("危害等级:"+link.contents[3].contents[5].find('a').text.lstrip().rstrip())#危害等级
        list4.append(str(link.contents[3].contents[5].find('a').text.lstrip().rstrip()))
    except:
        #print("危害等级:is empty")
        list4.append("")
    try:
        #print("CVE编号:"+link.contents[3].contents[7].find('a').text.lstrip().rstrip())#CVE编号
        list5.append(str(link.contents[3].contents[7].find('a').text.lstrip().rstrip()))
    except:
        #print("CVE编号:is empty")
        list5.append("")
    try:
        #print("漏洞类型:"+link.contents[3].contents[9].find('a').text.lstrip().rstrip())#漏洞类型
        list6.append(str(link.contents[3].contents[9].find('a').text.lstrip().rstrip()))
    except:
        #print("漏洞类型:is empty")
        list6.append("")
    try:
        #print("发布时间:"+link.contents[3].contents[11].find('a').text.lstrip().rstrip())#发布时间
        list7.append(str(link.contents[3].contents[11].find('a').text.lstrip().rstrip()))
    except:
        #print("发布时间:is empty")
        list7.append("")
    try:
        #print("威胁类型:"+link.contents[3].contents[13].find('a').text.lstrip().rstrip())#威胁类型
        list8.append(str(link.contents[3].contents[13].find('a').text.lstrip().rstrip()))
    except:
        #print("威胁类型:is empty")
        list8.append("")
    try:
        #print("更新时间:"+link.contents[3].contents[15].find('a').text.lstrip().rstrip())#更新时间
        list9.append(str(link.contents[3].contents[15].find('a').text.lstrip().rstrip()))
    except:
        #print("更新时间:is empty")
        list9.append("")
    try:
        #print("厂商:"+link.contents[3].contents[17].find('a').text.lstrip().rstrip())#厂商
        list10.append(str(link.contents[3].contents[17].find('a').text.lstrip().rstrip()))
    except:
        #print("厂商:is empty")
        list10.append("")

        # link_introduce=html.find(class_='d_ldjj')#漏洞简介
    try:
        link_introduce_data = BeautifulSoup(link_introduce.decode(), 'html.parser').find_all(name='p')
        s = ""
        for i in range(0, len(link_introduce_data)):
            ##print (link_introduce_data[i].text.lstrip().rstrip())
            s = s + str(link_introduce_data[i].text.lstrip().rstrip())
        # print(s)
        list11.append(s)
    except:
        list11.append("")

    if (len(link_others) != 0):
        # link_others=html.find_all(class_='d_ldjj m_t_20')
        # print(len(link_others))
        try:
            # 漏洞公告
            link_others_data1 = BeautifulSoup(link_others[0].decode(), 'html.parser').find_all(name='p')
            s = ""
            for i in range(0, len(link_others_data1)):
                ##print (link_others_data1[i].text.lstrip().rstrip())
                s = s + str(link_others_data1[i].text.lstrip().rstrip())
            # print(s)
            list12.append(s)
        except:
            list12.append("")

        try:
            # 参考网址
            link_others_data2 = BeautifulSoup(link_others[1].decode(), 'html.parser').find_all(name='p')
            s = ""
            for i in range(0, len(link_others_data2)):
                ##print (link_others_data2[i].text.lstrip().rstrip())
                s = s + str(link_others_data2[i].text.lstrip().rstrip())
            # print(s)
            list13.append(s)
        except:
            list13.append("")

        try:
            # 受影响实体
            link_others_data3 = BeautifulSoup(link_others[2].decode(), 'html.parser').find_all('a', attrs={
                'class': 'a_title2'})
            s = ""
            for i in range(0, len(link_others_data3)):
                ##print (link_others_data3[i].text.lstrip().rstrip())
                s = s + str(link_others_data3[i].text.lstrip().rstrip())
            # print(s)
            list14.append(s)
        except:
            list14.append("")

        try:
            # 补丁
            link_others_data3 = BeautifulSoup(link_others[3].decode(), 'html.parser').find_all('a', attrs={
                'class': 'a_title2'})
            s = ""
            for i in range(0, len(link_others_data3)):
                ##print (link_others_data3[i].text.lstrip().rstrip())
                s = s + str(link_others_data3[i].text.lstrip().rstrip())
            # print(s)
            list15.append(s)
        except:
            list15.append("")
    else:
        list12.append("")
        list13.append("")
        list14.append("")
        list15.append("")


if __name__ == "__main__":
    global list4
    global list5
    global list6
    global list7
    global list8
    global list9
    global list10
    global list11
    global list12
    global list13
    global list14
    global list15

    list1 = []  # 网站的url
    list2 = []  # 漏洞的名称
    list3 = []  # cnnvd编号
    list4 = []  # 危害等级
    list5 = []  # CVE编号
    list6 = []  # 漏洞类型
    list7 = []  # 发布时间
    list8 = []  # 威胁类型
    list9 = []  # 更新时间
    list10 = []  # 厂商
    list11 = []  # 漏洞简介
    list12 = []  # 漏洞公告
    list13 = []  # 参考网址
    list14 = []  # 受影响实体
    list15 = []  # 补丁
    start = 1
    last =1
    # url = 'http://www.cnnvd.org.cn/web/xxk/ldxqById.tag?CNNVD=CNNVD-201901-1014'
    # getURLDATA(url)
    f = xlwt.Workbook()  # 创建EXCEL工作簿
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    sheet1.write(0, 0, "漏洞名称")
    sheet1.write(0, 1, "网址")
    sheet1.write(0, 2, "CNNVD编号")
    sheet1.write(0, 3, "危害等级")
    sheet1.write(0, 4, "CVE编号")
    sheet1.write(0, 5, "漏洞类型")
    sheet1.write(0, 6, "发布时间")
    sheet1.write(0, 7, "威胁类型")
    sheet1.write(0, 8, "更新时间")
    sheet1.write(0, 9, "厂商")
    sheet1.write(0, 10, "漏洞简介")
    sheet1.write(0, 11, "漏洞公告")
    sheet1.write(0, 12, "参考网址")
    sheet1.write(0, 13, "受影响实体")
    sheet1.write(0, 14, "补丁")

    for j in range(start, last + 1):
        # url='http://www.cnnvd.org.cn/web/vulnerability/querylist.tag?pageno=1&repairLd='
        url = 'http://www.cnnvd.org.cn/web/vulnerability/querylist.tag?pageno=' + str(j) + '&repairLd='
        print("page" + str(j))
        header = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.80 Safari/537.36',
            'Connection': 'keep-alive', }
        r = requests.get(url, headers=header, timeout=30)
        # r.raise_for_status()抛出异常
        html = BeautifulSoup(r.content.decode(), 'html.parser')
        link = html.find_all(class_='a_title2')
        for i in link:
            ##print (i.text.lstrip())
            try:
                list1.append(i.text.lstrip())
                ##print ("http://www.cnnvd.org.cn"+i.attrs['href'])
                k = str(i.attrs['href'])
                list2.append("http://www.cnnvd.org.cn" + k)
                list3.append(k[28:])
                # print("http://www.cnnvd.org.cn"+k)
                getURLDATA("http://www.cnnvd.org.cn" + k)
            except:
                print("http://www.cnnvd.org.cn" + k)
                break

    for i in range(len(list15)):
        sheet1.write(i + 1, 0, list1[i])
        sheet1.write(i + 1, 1, list2[i])
        sheet1.write(i + 1, 2, list3[i])
        sheet1.write(i + 1, 3, list4[i])
        sheet1.write(i + 1, 4, list5[i])
        sheet1.write(i + 1, 5, list6[i])
        sheet1.write(i + 1, 6, list7[i])
        sheet1.write(i + 1, 7, list8[i])
        sheet1.write(i + 1, 8, list9[i])
        sheet1.write(i + 1, 9, list10[i])
        sheet1.write(i + 1, 10, list11[i])
        sheet1.write(i + 1, 11, list12[i])
        sheet1.write(i + 1, 12, list13[i])
        sheet1.write(i + 1, 13, list14[i])
        sheet1.write(i + 1, 14, list15[i])
    f.save(str(start) + "-" + str(last) + ".xls")  # 保存文件