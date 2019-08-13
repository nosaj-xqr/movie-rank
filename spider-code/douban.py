import xlwt
import re
import urllib.request
import urllib.error
import time
from lxml import etree
import requests

#使用xlwt建立excel文件
workbook=xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('movie')
row0=[u'电影',u'豆瓣评分',u'豆瓣排名',u'豆瓣评分人数']
for k in range(len(row0)):
    worksheet.write(0,k,row0[k])

#伪装浏览器并爬取
header=('User-Agent','Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:67.0) Gecko/20100101 Firefox/67.0')
opener=urllib.request.build_opener()
opener.addheaders=[header]
urllib.request.install_opener(opener)
url='https://movie.douban.com/top250?start=0'
num=0
for i in range(4):
    req=requests.get(url)
    html=etree.HTML(req.text)
    name=html.xpath('//ol/li/div/div[@class="info"]/div[@class="hd"]/a/span[1]/text()')
    vote=html.xpath('//ol/li/div/div[@class="info"]/div[@class="bd"]/div/span[2]/text()')
    rank=html.xpath('//ol/li/div/div[@class="pic"]/em/text()')
    people=html.xpath('//ol/li/div/div[@class="info"]/div[@class="bd"]/div/span[4]/text()')
    for j in range(0,len(name)):
        worksheet.write(i*len(name)+j+1,0,name[j])
        worksheet.write(i*len(name)+j+1,1,vote[j])
        worksheet.write(i*len(name)+j+1,2,rank[j])
        pat_num='(.*?)人评价'
        people_num=re.compile(pat_num).findall(people[j])
        worksheet.write(i*len(name)+j+1,3,people_num)
    num += 25
    url='https://movie.douban.com/top250?start='+str(num)
    time.sleep(10)
workbook.save('douban.xlsx')
