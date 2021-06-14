# -*- coding: utf-8 -*-
import requests
import json
import xlrd
import openpyxl
import xlsxwriter
import time
from xlutils.copy import copy
import os
from lxml import etree
import pymysql

conn = pymysql.connect(
    host='172.16.30.18',
    user='root',
    passwd='hua1315579747',
    db='heganghua_copy',
    port=3306,
    charset='utf8')
cursor = conn.cursor()


url = "http://218.76.40.80:9000/hnxyfw/channel/searchJSComInfo.jspx?queryType=1"
hosts = "http://218.76.40.80:9000"

headers = {
    "Cache-Control": "max-age=0",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9,und;q=0.8,en;q=0.7",
    # "Cookie": "JSESSIONID=C60787A55A58115F17E7E1E318AA7397; _site_id_cookie=1; clientlanguage=zh_CN",
    "Connection": "keep-alive"
}
cookies = {
    "JSESSIONID": "C60787A55A58115F17E7E1E318AA7397",
    "_site_id_cookie": "1",
    "clientlanguage": "zh_CN"
}


path = os.path.abspath(os.path.dirname(__file__))
workbook = xlrd.open_workbook(path + "/company.xlsx")
sheet = workbook.sheet_by_name('Sheet1')
col = sheet.col_values(0)

for j in range(0, len(col)):

    item = col[j]
    time.sleep(0.5)
    print(item)


# resp = requests.get(url=url, headers=headers, cookies=cookies)
# # print(resp.text)
# html = etree.HTML(resp.text)
#
# # 翻页
# for j in range(32, 41):
#     print(j, '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>')
#     time.sleep(3)
#     trList = html.xpath('//div[@class="mc_content"]/table/tbody/tr')
#     for i in range(1, len(trList)):
#         tr = trList[i]
#         code = tr.xpath('td[1]/text()')
#         code = code[0] if code else ""
#         name = tr.xpath('td[2]/a/text()')[0]
#         href = tr.xpath('td[2]/a/@href')[0]
#         print(code, name, href, "")
#         print("-------------------------------------------\n\n")
#         # 数据
#         messageResp = requests.get(
#             url=hosts+href, headers=headers, cookies=cookies)
#         messageHtml = etree.HTML(messageResp.text)
#         messageTrList = messageHtml.xpath(
#             '//div[@id="yjxxContent"]/div[2]/table/tbody/tr')
#
#         for k in range(1, len(messageTrList)):
#             messageTr = messageTrList[k]
#             detailHref = messageTr.xpath('td/a/@href')[0]
#             detailResp = requests.get(
#                 url=hosts+detailHref, headers=headers, cookies=cookies)
#             detailHtml = etree.HTML(detailResp.text)
#             # 工程名称
#             detailTr2 = detailHtml.xpath(
#                 '//div[@class="mc_content"]/table/tr[2]')[0]
#             projectName = detailTr2.xpath('td[2]/text()')
#             projectName = projectName[0] if projectName else ""
#             # 合同价
#             detailTr3 = detailHtml.xpath(
#                 '//div[@class="mc_content"]/table/tr[3]')[0]
#             price = detailTr3.xpath('td[2]/text()')
#             price = price[0] if price else ""
#             # 技术等级
#             detailTr4 = detailHtml.xpath(
#                 '//div[@class="mc_content"]/table/tr[4]')[0]
#             level = detailTr4.xpath('td[2]/text()')
#             level = level[0] if level else ""
#             # 交工日期
#             detailTr5 = detailHtml.xpath(
#                 '//div[@class="mc_content"]/table/tr[5]')[0]
#             date = detailTr5.xpath('td[4]/text()')
#             date = date[0] if date else ""
#             # 开始桩号
#             detailTr7 = detailHtml.xpath(
#                 '//div[@class="mc_content"]/table/tr[7]')[0]
#             bigen = detailTr7.xpath('td[2]/text()')
#             bigen = bigen[0] if bigen else ""
#             # 结束桩号
#             end = detailTr7.xpath('td[4]/text()')
#             end = end[0] if end else ""
#             # 主要工程量
#             detailTr8 = detailHtml.xpath(
#                 '//div[@class="mc_content"]/table/tr[9]')[0]
#             text = detailTr8.xpath('td[2]/text()')
#             text = text[0] if text else ""
#
#             print(name, projectName, price, level, date, bigen, end, text)
#             sql = "insert into dada values(%s, %s, %s, %s, %s, %s, %s, %s)"
#             values = (name, projectName, price, level, date, bigen, end, text)
#             try:
#                 cursor.execute(sql, values)
#                 conn.commit()
#             except Exception as e:
#                 print(e)
#
#     # next page
#     nextUrl = "http://218.76.40.80:9000/hnxyfw/channel/searchJSComInfo_{page}.jspx?comName=&comType=&queryType=1"
#     resp = requests.get(url=nextUrl.format(page=j),
#                         headers=headers, cookies=cookies)
#     html = etree.HTML(resp.text)


        # path = os.path.abspath(os.path.dirname(__file__))
        # workbook = xlrd.open_workbook(path + "/company.xlsx")
        # sheet = workbook.sheet_by_name('Sheet1')
        # col = sheet.col_values(0)

        # # 工作簿
        # workbook = xlsxwriter.Workbook(path + '/result.xlsx')
        # # 工作表
        # worksheet = workbook.add_worksheet()
        # worksheet.write('A1', '项目名称')
        # worksheet.write('B1', '公路等级')
        # worksheet.write('C1', '合同段开始桩号')
        # worksheet.write('D1', '合同段结束桩号')
        # worksheet.write('E1', '工程量')
        # worksheet.write('F1', '合同金额')
        # worksheet.write('G1', '交工时间')
        # worksheet.write('H1', '单位')

        # n = 1
        # for j in range(0, len(col)):
        #     print(j)
        #     item = col[j]
        #     time.sleep(0.5)
        #     print(item)
        #     data["text"] = item
        #     response = requests.post(url=url, headers=headers,
        #                              cookies=cookies, data=data, verify=False)
        #     resJson = json.loads(response.text)
        #     ros = resJson["rows"]
        #     if not ros:
        #         continue
        #     id = ros[0]["id"]
        #     companyList = requests.post(url=company_achieve_list_url.format(
        #         companyId=id), headers=headers, cookies=cookies, data=company_data, verify=False)
        #     # with open("f.txt", "w", encoding="utf-8") as f:
        #     #     f.write(companyList.text)
        #     companyListJson = json.loads(companyList.text)

        #     dataRows = companyListJson["rows"]
        #     for i in range(1, len(dataRows)):
        #         com = dataRows[i]
        #         # 项目名称
        #         projectName = com["projectName"]
        #         # 公路等级
        #         technologyGrade = com["technologyGrade"]
        #         # 合同段开始桩号
        #         stakeEnd = com["stakeEnd"]
        #         # 合同段结束桩号
        #         stakeStart = com["stakeStart"]
        #         # 工程量
        #         remark = com["remark"]
        #         # 合同金额
        #         contractPrice = com["contractPrice"]
        #         # 交工时间
        #         handDate = com["handDate"]

        #         worksheet.write(n, 0, projectName)
        #         worksheet.write(n, 1, technologyGrade)
        #         worksheet.write(n, 2, stakeEnd)
        #         worksheet.write(n, 3, stakeStart)
        #         worksheet.write(n, 4, remark)
        #         worksheet.write(n, 5, contractPrice)
        #         worksheet.write(n, 6, handDate)
        #         worksheet.write(n, 7, item)
        #         n += 1
        #     # break
        # workbook.close()
        # new_workbook.save('result.xlsx')  # 保存工作簿
