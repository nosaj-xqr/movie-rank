#注：在分析网站时，发现网页上有但源代码信息里没有评分人数显示，之后采用fiddler进行抓包分析也未找到，不知道什么情况。。
import xlwt
import re
import urllib.request
import urllib.error
import time
from lxml import etree
import requests

workbook=xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('movie')
row0=[u'电影',u'猫眼评分',u'猫眼排名']
for k in range(len(row0)):
    worksheet.write(0,k,row0[k])

header=('User-Agent','Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:67.0) Gecko/20100101 Firefox/67.0')
opener=urllib.request.build_opener()
opener.addheaders=[header]
urllib.request.install_opener(opener)
url='https://maoyan.com/board/4?offset=0'
num=0
for i in range(10):
    req=requests.get(url)
    html=etree.HTML(req.text)
    name=html.xpath('//p[@class="name"]/a/text()')
    vote=html.xpath('//p[@class="score"]')
    for m in range(len(vote)):
        vote[m] = vote[m].xpath('string(.)')
    rank=html.xpath('//dd/i/text()')
    for j in range(0,len(name)):
        worksheet.write(i*len(name)+j+1,0,name[j])
        worksheet.write(i*len(name)+j+1,1,vote[j])
        worksheet.write(i*len(name)+j+1,2,rank[j])
    num += 10
    url='https://maoyan.com/board/4?offset='+str(num)
    time.sleep(10)
workbook.save('maoyan.xlsx')
