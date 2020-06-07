# Analysis of geographical locations from processed data (分析基础)

import pandas as pd
from openpyxl import Workbook

isDebug = False

dfInput = pd.read_excel("200530-5-分析基础.xlsx")
wbOutput = Workbook()

wsProvince = wbOutput.active
wsProvince.title = "省份分布"
wsCity = wbOutput.create_sheet(title = "城市分布")

lstProvinceRaw = dfInput["省份"]
dictProvinceFreq = {}
for i in lstProvinceRaw:
    if i in dictProvinceFreq:
        dictProvinceFreq[i] += 1
    else:
        dictProvinceFreq[i] = 1

if isDebug == True:
    print("Province Frequency Dictionary", dictProvinceFreq)

wsProvince.cell(column = 1, row = 1, value = "省份")
wsProvince.cell(column = 2, row = 1, value = "人数")
tempRow = 2
for i in dictProvinceFreq:
    wsProvince.cell(column = 1, row = tempRow, value = i)
    tempRow += 1
tempRow = 2
for i in dictProvinceFreq:
    wsProvince.cell(column= 2, row = tempRow, value = dictProvinceFreq[i])
    tempRow += 1


lstCityRaw = dfInput["考点"]
dictCityFreq = {}
for i in lstCityRaw:
    if i in dictCityFreq:
        dictCityFreq[i] += 1
    else:
        dictCityFreq[i] = 1

if isDebug == True:
    print("City Frequency Dictionary", dictCityFreq)

wsCity.cell(column = 1, row = 1, value = "城市")
wsCity.cell(column = 2, row = 1, value = "人数")
tempRow = 2
for i in dictCityFreq:
    wsCity.cell(column = 1, row = tempRow, value = i)
    tempRow += 1
tempRow = 2
for i in dictCityFreq:
    wsCity.cell(column= 2, row = tempRow, value = dictCityFreq[i])
    tempRow += 1

wbOutput.save(filename = "200530-0-省份与地区测试文件.xlsx")



