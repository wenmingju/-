import requests
import xlrd
from xlutils.copy import copy
url = 'http://ezcarry.tpddns.cn:8088/GetContainerEvent.ashx'
excel = xlrd.open_workbook('C:/Users/14620\Documents\Tencent Files/1462007678\FileRecv\港区数据测试\测试结果.xlsx')
sheet1 = excel.sheet_by_name('Sheet1')
cc = sheet1.col_values(1)
row = sheet1.nrows
wb = copy(excel)
ws = wb.get_sheet(0)  # 1代表是写到第几个工作表里，从0开始算是第一个。
for i in range(1,row):
    value = sheet1.cell(i,1).value
    company = value[:4]
    ne = value[4:]
    param = {'MawbPrefix': company, 'SMLM&ContainerNo': value, 'MawbSerial': ne,
                              'key': 'eft', 'MIC': '******'}
    re = requests.get(url=url, params=param)
    if '<?xml version="1.0"?>' in re.text:
        if '<OnBoardDate>' in re.text:
            gs = 1
            data = 1
        else:
            gs = 1
            data = 0
    else:
            gs = 0
            data = 0
    print(company, gs, data)
    ws.write(i, 2, gs)
    ws.write(i, 3, data)
wb.save('C:/Users/14620\Documents\Tencent Files/1462007678\FileRecv\港区数据测试\测试结果.xlsx')