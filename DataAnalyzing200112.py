import pandas as pd
import numpy as np
import math

# STEMA考试列表
# 第1次 2019 12 15
# 第2次 2020 01 12

# 设置考试日期字典
from numpy.core._multiarray_umath import ndarray
from pandas import DataFrame

dateTest = {1 : "2019-12-15", 2 : "2020-01-12"}

# 设置考试数据分析基础文件的文件名
fileAnalysisInput = {1 : "191215-5-分析基础.xlsx", 2 : "200112-5-分析基础.xlsx"}
dfAnalysisInput = list()
for i in range(len(fileAnalysisInput)) :
    print("Program Info ->", "考试序号", i, "考试日期", dateTest[i + 1], "文件名称", fileAnalysisInput[i + 1])
    dfAnalysisInput.append(pd.read_excel(fileAnalysisInput[i + 1], "分析基础", index_col="准考证号"))

# 设置考试数据分析基础文件的文件
fileAnalysisOutput = {1 : "191215-6-分析结果.xlsx", 2 : "200112-6-分析结果.xlsx"}
excelWriter = list()
for i in range(len(fileAnalysisOutput)) :
    excelWriter.append(pd.ExcelWriter(fileAnalysisOutput[i + 1]))

# 设置全局分析结果的文件
fileAnalysisGlobal = "7-分析结果.xlsx"
excelWriterGlobal = pd.ExcelWriter(fileAnalysisGlobal)

# 函数功能 - 统计每题目的正确率
def AnswerPercentage() :
    # 是否输出Debug信息
    isDebug = False

    print("Program Info ->", "进入AnswerPercentage()函数 开始每题目正确率", "考试次数：", len(fileAnalysisInput))
    for iTest in range(len(fileAnalysisInput)) :
        npSurvey = np.empty([0, 9])
        for iLevel in ['初级', '中级', '高级'] :
            if isDebug : print("Debug Info ->", "考试", dateTest[iTest + 1], "级别", iLevel)
            npTmp = np.empty([0, 9])
            dfLevelTmp = dfAnalysisInput[iTest][dfAnalysisInput[iTest]['级别']==iLevel]
            for i in range(1, 73) :
                lCurrentAnswer = dfLevelTmp["答案"+str(i)].unique()
                if isDebug : print("Debug Info ->", "答案"+str(i), lCurrentAnswer)
                pA = pB = pC = pD = pE = pF = pN = pZ = 0
                if isDebug : print("Debug Info ->", dfLevelTmp["答案" + str(i)].value_counts(dropna=False))
                while len(lCurrentAnswer) > 0 :
                    sCurrentAnswer = str(lCurrentAnswer[0])
                    lCurrentAnswer = np.delete(lCurrentAnswer, 0)
                    if sCurrentAnswer.strip() == "A" :
                        pA += dfLevelTmp["答案"+str(i)].value_counts()[sCurrentAnswer] / len(dfLevelTmp) * 100
                    elif sCurrentAnswer.strip() == "B" :
                        pB += dfLevelTmp["答案" + str(i)].value_counts()[sCurrentAnswer] / len(dfLevelTmp) * 100
                    elif sCurrentAnswer.strip() == "C" :
                        pC += dfLevelTmp["答案" + str(i)].value_counts()[sCurrentAnswer] / len(dfLevelTmp) * 100
                    elif sCurrentAnswer.strip() == "D":
                        pD += dfLevelTmp["答案" + str(i)].value_counts()[sCurrentAnswer] / len(dfLevelTmp) * 100
                    elif sCurrentAnswer.strip() == "E":
                        pE += dfLevelTmp["答案" + str(i)].value_counts()[sCurrentAnswer] / len(dfLevelTmp) * 100
                    elif sCurrentAnswer.strip() == "F":
                        pD += dfLevelTmp["答案" + str(i)].value_counts()[sCurrentAnswer] / len(dfLevelTmp) * 100
                    elif sCurrentAnswer.strip() == "nan" or sCurrentAnswer.strip() == "":
                        pN += dfLevelTmp["答案" + str(i)].isnull().sum() / len(dfLevelTmp) * 100
                    else :
                        pZ += dfLevelTmp["答案" + str(i)].value_counts()[sCurrentAnswer] / len(dfLevelTmp) * 100
                npTmp = np.append(npTmp, [["答案" + str(i), round(pA, 1), round(pB, 1), round(pC, 1), round(pD, 1), round(pE, 1), round(pF, 1),
                     round(pN, 1), round(pZ, 1)]], axis=0)
            if isDebug : print("Debug Info ->", npTmp)
            dfTmp = pd.DataFrame(npTmp, columns=['题目', '选A', '选B', '选C', '选D', '选E', '选F', '选空', '其它'])
            dfTmp = dfTmp.set_index('题目')
            # 将DataFrame列表的当前项写入分析结果文件
            if isDebug : print("Debug Info ->", "保存分析结果文件 考试", dateTest[iTest], iLevel)
            dfTmp.to_excel(excelWriter[iTest], sheet_name=iLevel+'正确率')

    return 0

# 函数功能 - 统计共有多少城市参加，并给出省份列表
def CityUnique() :
    # 是否输出Debug信息
    isDebug = False
    # 城市与所属省份的对应关系
    dictCity = {'北京':'北京'}

    print("Program Info ->", "进入CityUnique()函数 开始统计参赛城市数据", "考试次数：", len(fileAnalysisInput))

    # 定义列表，存储每次考试数据的DataFrame
    dfTmp = []
    # 处理每次考试的数据
    for iTest in range(len(fileAnalysisInput)) :
        # 本次考试省份列表
        arChengShi = dfAnalysisInput[iTest]['考点'].unique()
        # 构建序列，存储本次考试情况
        npTmp = np.empty([0, 3])
        for i in range(len(arChengShi)) :
            # 填写城市与所属省份的对应关系字典
            dictCity[arChengShi[i]] = dfAnalysisInput[iTest][dfAnalysisInput[iTest]['考点']==arChengShi[i]].iloc[0]['省份']
            # 将省份名称、城市名称、人数加入序列
            npTmp = np.append(npTmp, [[dictCity[arChengShi[i]], arChengShi[i], int(dfAnalysisInput[iTest]['考点'].value_counts()[arChengShi[i]])]], axis=0)
        if isDebug :
            print("Debug Info ->", "考试", iTest + 1, "日期", dateTest[iTest + 1])
            print(npTmp)
        # 将序列存入DataFrame列表
        dfTmp.append(pd.DataFrame(npTmp, columns=['省份', '城市', '人数']))
        dfTmp[iTest]['人数'] = dfTmp[iTest]['人数'].astype('int')
        dfTmp[iTest] = dfTmp[iTest].set_index('城市')
        dfTmp[iTest] = dfTmp[iTest].sort_values(by="省份")
        # 将DataFrame列表的当前项写入分析结果文件
        dfTmp[iTest].to_excel(excelWriter[iTest], sheet_name='参加城市')

    # 合并所有考试数据
    # 将第一次考试的数据装入汇总表中
    dfTotal = dfTmp[0]
    # 将后继考试的数据装入汇总表中
    for iTest in range(1, len(fileAnalysisInput)) :
        dfTotal = pd.concat([dfTotal, dfTmp[iTest]], axis=1, sort=True)
    # 将各次考试数据相加求和
    dfTotal= pd.DataFrame(dfTotal.sum(axis=1))
    dfTotal.index.name = '城市'
    npTmp = np.empty([0, 1])
    for i in range(len(dfTotal)) :
        npTmp = np.append(npTmp, [dictCity[dfTotal.index[i]]])
    dfTotal.insert(0, '省份', npTmp)
    dfTotal.columns = ['省份', '人数']
    dfTotal['人数'] = dfTotal['人数'].astype('int')
    dfTotal = dfTotal.sort_values(by='省份')
    if isDebug :
        print("Debug Info ->", "合并后的数据")
        print(dfTotal)

    # 写入分析结果文件
    dfTotal.to_excel(excelWriterGlobal, sheet_name='参加城市')

    return 0

# 函数功能 - 统计共有多少省份参加，并给出省份列表
def ProvinceUnique() :
    # 是否输出Debug信息
    isDebug = False

    print("Program Info ->", "进入ProvinceUnique()函数 开始统计参赛省份数据", "考试次数：", len(fileAnalysisInput))

    # 定义列表，存储每次考试数据的DataFrame
    dfTmp = []
    # 处理每次考试的数据
    for iTest in range(len(fileAnalysisInput)) :
        # 本次考试省份列表
        arShengFen = dfAnalysisInput[iTest]['省份'].unique()
        # 构建序列，存储本次考试情况
        npTmp = np.empty([0, 2])
        for i in range(len(arShengFen)) :
            # 将省份名称、人数加入序列
            npTmp = np.append(npTmp, [[arShengFen[i], int(dfAnalysisInput[iTest]['省份'].value_counts()[arShengFen[i]])]],
                              axis=0)
        if isDebug :
            print("Debug Info ->", "考试", iTest + 1, "日期", dateTest[iTest + 1])
            print(npTmp)
        # 将序列存入DataFrame列表
        dfTmp.append(pd.DataFrame(npTmp, columns=['省份', '人数']))
        dfTmp[iTest]['人数'] = dfTmp[iTest]['人数'].astype('int')
        dfTmp[iTest] = dfTmp[iTest].set_index('省份')
        # 将DataFrame列表的当前项写入分析结果文件
        dfTmp[iTest].to_excel(excelWriter[iTest], sheet_name='参加省份')

    # 合并所有考试数据
    # 将第一次考试的数据装入汇总表中
    dfTotal = dfTmp[0]
    # 将后继考试的数据装入汇总表中
    for iTest in range(1, len(fileAnalysisInput)) :
        dfTotal = pd.concat([dfTotal, dfTmp[iTest]], axis=1, sort=True)
    # 将各次考试数据相加求和
    dfTotal = pd.DataFrame(dfTotal.sum(axis=1))
    dfTotal.index.name = '省份'
    dfTotal.columns = ['人数']
    dfTotal['人数'] = dfTotal['人数'].astype('int')
    if isDebug :
        print("Debug Info ->", "合并后的数据")
        print(dfTotal)

    # 写入分析结果文件
    dfTotal.to_excel(excelWriterGlobal, sheet_name='参加省份')

    return 0

# 主程序开始-------------------------------------------------------------------------------------------------------------

# 统计参加省份数据
ProvinceUnique()

# 统计参加城市的数据
CityUnique()

# 统计每题目的正确率
AnswerPercentage()

# 关闭需要写入的excel文件
print("Program Info ->", "写入分析结果文件...")
for iTest in range(len(fileAnalysisInput)) :
    excelWriter[iTest].save()
    excelWriter[iTest].close()
excelWriterGlobal.save()
excelWriterGlobal.close()

# 主程序结束-------------------------------------------------------------------------------------------------------------
