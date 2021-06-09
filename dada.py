# -*- coding: utf-8 -*-
import requests
import json
import xlrd
import openpyxl
import xlsxwriter
import time
from xlutils.copy import copy
import os
import pymysql


class GetCompanyInfo:
    def __init__(self):

        self.conn = pymysql.connect(
            host='127.0.0.1',
            user='root',
            passwd='root',
            db='datacenter',
            port=3306,
            charset='utf8')
        self.cursor = self.conn.cursor()

        self.url = "https://glxy.mot.gov.cn/company/getCompanyAptitude.do"
        self.company_achieve_list_url = "https://glxy.mot.gov.cn/company/getCompanyAchieveList.do?companyId={companyId}&type=11"
        self.headers = {
            "Host": "glxy.mot.gov.cn",
            "Connection": "keep-alive",
            # "Content-Length": "175",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Origin": "https://glxy.mot.gov.cn",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Dest": "empty",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "zh-CN,zh;q=0.9,und;q=0.8,en;q=0.7"
        }
        self.cookies = {
            "JSESSIONID": "62324EB4A1F24BA135DABBECF55BE147"
        }
        self.data = {
            "page": "1",
            "rows": "15",
            "type": "0",
            "caname": "公路工程",
            "regProvinceCode": "",
            "catype": "总承包",
            "grade": "",
            "text": ""
        }
        self.company_data = {
            "page": "1",
            "rows": "500",
            "sourceInfo": "1",
            "projectname": "",
            "provinceSearch": "",
        }
        # 路径
        # self.path = os.path.abspath(os.path.dirname(__file__))
        # # workbook = xlrd.open_workbook(self.path + "/company.xlsx")
        # workbook = xlsxwriter.Workbook(self.path + '/newCompany.xlsx')
        # self.worksheet = workbook.add_worksheet()
        # self.worksheet.write('A1', 'companyName')
        # self.worksheet.write('B1', 'companyId')
        # self.rowNum = 0

    def _close_database_connection(self):
        print("Is closing connection...")
        self.cursor.close()
        self.conn.close()
        print("close over")

    def requestCompanyInfo(self):
        """请求公司列表"""
        for page in range(1, 727):
            print("页码：", str(page))
            self.data["page"] = str(page)
            self.getInfoProcess()
        # 关闭连接
        self._close_database_connection()

    def getInfoProcess(self):
        """
            获取公司相关信息，
        """
        # 请求列表
        response = requests.post(url=self.url, headers=self.headers, cookies=self.cookies,
                                 data=self.data, verify=False)
        resJson = json.loads(response.text)
        rows = resJson["rows"]
        for i in range(0, len(rows)):
            row = rows[i]
            id = row["id"]
            corpName = row["corpName"]
            print(corpName, id)
            self.rowNum += 1
            # self.saveData(corpName, id)

    def saveDataForExcel(self, companyName, companyId):
        """保存到Excel"""
        self.worksheet.write(self.rowNum, 0, companyName)
        self.worksheet.write(self.rowNum, 0, companyId)

    def saveDataForMySQL(self, companyName, companyId):
        """保存到数据"""

        sql = 'insert into dada_company(company_name, company_id) values(%s, %s)'
        values = (companyName, companyId)
        try:
            self.cursor.execute(sql, values)
            self.conn.commit()
        except Exception as e:
            self._close_database_connection()
            print(e)


if __name__ == '__main__':
    gc = GetCompanyInfo()
    gc.requestCompanyInfo()







# workbook = xlrd.open_workbook(path + "/company.xlsx")
# sheet = workbook.sheet_by_name('Sheet1')
# col = sheet.col_values(0)

# 工作簿

# 工作表
# worksheet = workbook.add_worksheet()
# # worksheet.write('A1', '项目名称')
# # worksheet.write('B1', '公路等级')
# # worksheet.write('C1', '合同段开始桩号')
# # worksheet.write('D1', '合同段结束桩号')
# # worksheet.write('E1', '工程量')
# # worksheet.write('F1', '合同金额')
# # worksheet.write('G1', '交工时间')
# # worksheet.write('H1', '单位')
# worksheet.write('A1', '公司名称')
# worksheet.write('B1', '资质等级')
# worksheet.write('C1', '项目名称')
# worksheet.write('D1', '合同段开始桩号')
# worksheet.write('E1', '合同段结束桩号')
# worksheet.write('F1', '金额(万元)')
# worksheet.write('G1', '公路等级')
# worksheet.write('H1', '工程内容')
# worksheet.write('I1', '开工时间')
# worksheet.write('J1', '交工时间')

# workbook = xlrd.open_workbook('result.xlsx')  # 打开工作簿
# sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
# worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
# rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
# new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
# new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格




#
#
#
#
# n = 1
# for j in range(0, len(col)):
#     print(j)
#     item = col[j]
#     time.sleep(0.5)
#     print(item)
#     data["text"] = item
#     response = requests.post(url=url, headers=headers, cookies=cookies, data=data, verify=False)
#     resJson = json.loads(response.text)
#     ros = resJson["rows"]
#     if not ros:
#         continue
#     id = ros[0]["id"]
#     companyList = requests.post(url=company_achieve_list_url.format(companyId=id), headers=headers, cookies=cookies, data=company_data, verify=False )
#     # with open("f.txt", "w", encoding="utf-8") as f:
#     #     f.write(companyList.text)
#     companyListJson = json.loads(companyList.text)
#
#     dataRows = companyListJson["rows"]
#     for i in range(0, len(dataRows)):
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
#
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