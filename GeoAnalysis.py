# Analysis of geographical locations from processed data (分析基础)

import pandas as pd
from openpyxl import Workbook

# Debug switch
isDebug = True

# Set input file (single file only) and create output workbook
dfInput = pd.read_excel("200530-5-分析基础.xlsx", read_only=True)
wbOutput = Workbook()
wsProvince = wbOutput.active
wsProvince.title = "省份与城市信息"

# Create Province Distribution Dictionary from dfInput
lstProvinceRaw = dfInput["省份"]
dictProvinceFreq = {}
for i in lstProvinceRaw:
    if i in dictProvinceFreq:
        dictProvinceFreq[i] += 1
    else:
        dictProvinceFreq[i] = 1
if isDebug:
    print("Province Frequency Dictionary", dictProvinceFreq)

# Write Dictionary into Excel
# Write headers
wsProvince.cell(column=1, row=1, value="省份")
wsProvince.cell(column=2, row=1, value="人数")
# Initialize starting row
tempRow = 2
# Write in Dictionary Index
for i in dictProvinceFreq:
    wsProvince.cell(column=1, row=tempRow, value=i)
    tempRow += 1
# Reset starting row
tempRow = 2
# Write in Dictionary Value
for i in dictProvinceFreq:
    wsProvince.cell(column=2, row=tempRow, value=dictProvinceFreq[i])
    tempRow += 1

# Create City Distribution Dictionary from dfInput
lstCityRaw = dfInput["考点"]
dictCityFreq = {}
for i in lstCityRaw:
    if i in dictCityFreq:
        dictCityFreq[i] += 1
    else:
        dictCityFreq[i] = 1
if isDebug:
    print("City Frequency Dictionary", dictCityFreq)

# Write Dictionary into Excel
# Write Headers
wsProvince.cell(column=4, row=1, value="城市")
wsProvince.cell(column=5, row=1, value="人数")
# Reset starting row
tempRow = 2
# Write in Dictionary Index
for i in dictCityFreq:
    wsProvince.cell(column=4, row=tempRow, value=i)
    tempRow += 1
# Reset Starting Row
tempRow = 2
# Write in Dictionary Value
for i in dictCityFreq:
    wsProvince.cell(column=5, row=tempRow, value=dictCityFreq[i])
    tempRow += 1

# Save File
wbOutput.save(filename="200530-0-省份与地区测试文件.xlsx")
