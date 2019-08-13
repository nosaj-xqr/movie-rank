import xlwt
import re
import urllib.request
import urllib.error
import time
from lxml import etree
import requests

workbook=xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('movie')
row0=[u'电影',u'时光网评分',u'时光网排名',u'时光网评分人数']
for k in range(len(row0)):
    worksheet.write(0,k,row0[k])

header=('User-Agent','Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:67.0) Gecko/20100101 Firefox/67.0')
opener=urllib.request.build_opener()
opener.addheaders=[header]
urllib.request.install_opener(opener)
url='http://www.mtime.com/top/movie/top100/'
num=1
for i in range(10):
    req=requests.get(url)
    html=etree.HTML(req.text)
    name=html.xpath('//div[@class="mov_pic"]/a/@title')
    vote=html.xpath('//div[@class="mov_point"]/b')
    for m in range(len(vote)):
        vote[m] = vote[m].xpath('string(.)')
    rank=html.xpath('//div[@class="number"]/em/text()')
    people=html.xpath('//div[@class="mov_point"]/p/text()')
    for j in range(0,len(name)):
        pat_name='(.*?)/'
        name_cn=re.compile(pat_name).findall(name[j])
        worksheet.write(i*len(name)+j+1,0,name_cn)
        #在分析网站时，发现此部电影没有评分信息，遂手动添加了一个。。
        if name_cn == ['疯狂动物城']:
            vote.insert(j,'8.6')
        worksheet.write(i*len(name)+j+1,1,vote[j])
        worksheet.write(i*len(name)+j+1,2,rank[j])
        pat_num='(.*?)人评分'
        people_num=re.compile(pat_num).findall(people[j])
        worksheet.write(i*len(name)+j+1,3,people_num)
    num += 1
    url='http://www.mtime.com/top/movie/top100/index-'+str(num)+'.html'
    time.sleep(10)
workbook.save('mtime.xlsx')
