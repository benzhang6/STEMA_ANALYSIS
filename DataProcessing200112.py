# 本程序是2020年1月12日考试的阅卷程序

import pandas as pd
import numpy as np
import math
from scipy import stats

# 本次考试是第几次STEMA考试
# 第1次 2019 12 15
# 第2次 2020 01 12
paraNoOfTest = 2

# 设置全局成绩文件名
if paraNoOfTest == 1 :
    file1 = "191215-1-收到汇总.xlsx"
    file2 = "191215-2-整理去除项.xlsx"
    file3 = "191215-3-分数评判.xlsx"
    file4 = "191215-4-发布成绩.xlsx"
    file5 = "191215-5-分析基础.xlsx"
elif paraNoOfTest == 2 :
    file1 = "200112-1-收到汇总.xlsx"
    file2 = "200112-2-整理去除项.xlsx"
    file3 = "200112-3-分数评判.xlsx"
    file4 = "200112-4-发布成绩.xlsx"
    file5 = "200112-5-分析基础.xlsx"

# 时间偏移设置
paraTime = (paraNoOfTest - 1) * 2

# 预留给未来按难度系数调整
paraDifficulty = 0

# 正态分布最高1%的分数
paraHighScore = 350 + paraTime + paraDifficulty
# 正态分布最低1%的分数
paraLowScore = 150
# 最终分数的步长，即间隔
paraScoreStep = 5

# 正态分布平均值
paraMu: int = (paraHighScore + paraLowScore) / 2
# 正态分布方差
paraSigma: int = math.ceil(0.841*(paraHighScore-paraLowScore)/4)

# 设置成绩处理数据结构
print("Program Info ->", "读取 汇总答案...")
dfFinal = pd.read_excel(file2, "汇总答案", index_col="准考证号")
nTotal = len(dfFinal)
print("Program Info ->", "核算 答案总人数：", nTotal)

# 本函数功能 - 检查数据完整性
def CheckData() :
    # 编程题总分不会超过128分
    # 初级组编程没有第五题
    # 级别不能有 初级 中级 高级 以外
    # 组别不能有 Python Scratch 以外
    # 答案有ABCDE之外的字符
    # 警告 选择题全空
    # 警告 编程题全空

    return 0

# 本函数功能 - 选择题判卷
def MarkingChoice() :
    global dfFinal
    # 是否打开调试输出
    isDebug = False

    # 读取答案表
    dfChujiDaan = pd.read_excel(file2, sheet_name="初级组选择答案")
    dfZhongjiDaan = pd.read_excel(file2, sheet_name="中高级组选择答案")

    # 定义新判分数，原分数是否正确两个列表
    npNewScore = np.zeros(dfFinal.shape[0])
    npIsCorrect = np.zeros(dfFinal.shape[0])

    # 定义需要详细审查（即输出Debug信息）的学生序号
    lStudent = []

    print("Program Info ->", "进入MarkingChoice()函数 开始选择题目判分", "答案总人数：", nTotal)

    # 选择题判分
    for j in range(0, nTotal) :
        nZongfen = 0
        # 初级组选择题判分
        if dfFinal.iloc[j]['级别'] == '初级' :
            for i in range(9, 73) :
                if str(dfFinal.iloc[j]['答案' + str(i)]).strip() == str(dfChujiDaan.iloc[0]['答案' + str(i)]).strip() :
                    nZongfen += 2
                elif str(dfFinal.iloc[j]['答案' + str(i)]) == "nan" or str(dfFinal.iloc[j]['答案' + str(i)]).strip() == "" :
                    pass
                else :
                    nZongfen -= 1
                if j in lStudent :
                    if isDebug : print("Debug Info ->", "题目", i, "学生答案", dfFinal.iloc[j]['答案' + str(i)], "正确答案", dfChujiDaan.iloc[0]['答案' + str(i)], "原分数", dfFinal.iloc[j]['选择' + str(i)], "现总分", nZongfen)
        #中高级组选择题判分
        else :
            for i in range(9, 73) :
                if str(dfFinal.iloc[j]['答案' + str(i)]).strip() == str(dfZhongjiDaan.iloc[0]['答案' + str(i)]).strip() :
                    nZongfen += 2
                elif str(dfFinal.iloc[j]['答案' + str(i)]) == "nan" or str(dfFinal.iloc[j]['答案' + str(i)]).strip() == "" :
                    pass
                else :
                    nZongfen -= 1
                if j in lStudent :
                    if isDebug : print("Debug Info ->", "题目", i, "学生答案", dfFinal.iloc[j]['答案' + str(i)], "正确答案", dfZhongjiDaan.iloc[0]['答案' + str(i)], "原分数", dfFinal.iloc[j]['选择' + str(i)], "现总分", nZongfen)
        #判分结果与原有Excel表中判分结果是否相同
        if nZongfen != int(dfFinal.iloc[j]['选择总分']) :
            npIsCorrect[j] = False
        else :
            npIsCorrect[j] = True
        #记录新计算选择题分数
        npNewScore[j] = nZongfen
        #显示新计算成绩与原成绩不一致的项目

        if isDebug :
            print("Debug Info ->", "正在处理", j, "正确" if npIsCorrect[j] else "错误", "原成绩：", dfFinal.iloc[j]['选择总分'], "现成绩：", nZongfen)
        else :
            if not npIsCorrect[j] : print("Program Info ->", "正在处理", j, "正确" if npIsCorrect[j] else "错误", "原成绩：", dfFinal.iloc[j]['选择总分'], "现成绩：", nZongfen)

    dfFinal.insert(0, '新计算选择总成绩', npNewScore)
    dfFinal.insert(0, '正确与否', npIsCorrect)

    return 0

def CurveChoice(sJiBie) : # 本函数功能 - 选择题核算发布分数，曲率计算
    global dfFinal
    # 是否打开调试输出
    isDebug = False

    # 从总体数据中取出当前级别的成绩数据
    dfJiBie = dfFinal[dfFinal["级别"] == sJiBie]
    # 核算当前级别总人数
    nStudentCount = len(dfJiBie)
    print("Program Info ->", "进入CurveChoice()函数 开始选择曲率判分", sJiBie, "总人数：", nStudentCount)
    # 按当前级别选择题成绩排序
    dfJiBie = dfJiBie.sort_values(by='新计算选择总成绩')

    # 定义当前级别最终分数序列
    npCurveScore = np.zeros(dfJiBie.shape[0])
    # i 为学生序号当前值，j 为最终分数当前值
    i = 0
    j = 0
    for j in range(paraLowScore, paraHighScore, paraScoreStep) :
        # tmpPercentage 为当前分数j在正态后所处的百分比
        tmpPercentage = stats.norm.cdf(j, paraMu, paraSigma)
        if isDebug : print("Debug Info ->", "进入新分数段", "序号", i, "成绩", j, round(stats.norm.cdf(j, paraMu, paraSigma) * 100, 1), "%")
        # 如果当前学生排名位置小于此百分比，意即学生应该得此分数
        while ((i+1)/nStudentCount) < tmpPercentage :
            if isDebug : print("Debug Info ->", "填充分数", "序号", i, "成绩", j)
            npCurveScore[i] = j
            i += 1
        # 除第一个学生外，超过此百分比，但和前一名学生分数一样的情况，也赋予同样分数
        if i > 0 :
            while (dfJiBie.iloc[i]['新计算选择总成绩'] == dfJiBie.iloc[i-1]['新计算选择总成绩']) :
                if isDebug : print("Debug Info ->", "成绩相同并列填充", "序号", i, "成绩", j)
                npCurveScore[i] = j
                i += 1
    # 如果上述计算结束，仍有部分学生没有分数，均赋予最高分
    if isDebug : print("Debug Info ->", "目前序号", i, "正态分布99%外获得最高分的学生人数：", nStudentCount-i, "分数：",j)
    while i < nStudentCount :
        if isDebug : print("Debug Info ->", "正态99%外填充", "序号", i, "成绩", j)
        npCurveScore[i] = j
        i += 1

    # 定义当前级别的全国%成绩
    npPercentageScore = np.zeros(dfJiBie.shape[0])
    # 分数最高的学生，%成绩永远是99%
    npPercentageScore[nStudentCount-1] = 0.99
    for i in range(nStudentCount-2, -1, -1) :
        if npCurveScore[i] == npCurveScore[i+1] :
            npPercentageScore[i] = npPercentageScore[i+1]
        else :
            npPercentageScore[i] = math.floor((i+1)/nStudentCount*100)/100
            if npPercentageScore[i]==0.00 : npPercentageScore[i] = 0.01

    # 计算需要填写的两列在表格中的列序号
    index1 = list(dfFinal.columns).index('第一部分成绩')
    index2 = list(dfFinal.columns).index('第一部分全国%')
    # 成绩表排序
    dfFinal = dfFinal.sort_values(by=['级别', '新计算选择总成绩'])
    # 组合选择题成绩进入原成绩数据
    i = 0
    while dfFinal.iloc[i]['级别'] != sJiBie :
        i += 1
    j = i
    while dfFinal.iloc[i]['级别'] == sJiBie :
        if isDebug : print("Debug Info ->", "正在填充全局变量dfFinal", "当前序列号", i, "起始序列号", j, dfFinal.iloc[i]['级别'])
        dfFinal.iloc[i, index1] = npCurveScore[i-j]
        dfFinal.iloc[i, index2] = format(npPercentageScore[i-j], ".0%")
        i += 1
        # 避免i超出学生总数，在上面while语句判断时产生下标益处
        if i-j>=nStudentCount: break

    return 0

# 本函数功能 - 编程题核算发布分数，曲率计算
def CurveProgram(sJiBie, sZuBie) :
    global dfFinal
    # 是否打开调试输出
    isDebug = False

    # 从总体数据中取出当前级别的成绩数据
    dfJiBieZuBie = dfFinal[dfFinal["级别"] == sJiBie]
    # 进一步选出当前组别的成绩数据
    dfJiBieZuBie = dfJiBieZuBie[dfJiBieZuBie["组别"] == sZuBie]
    # 核算当前级别组别总人数
    nStudentCount = len(dfJiBieZuBie)
    print("Program Info ->", "进入CurveProgram()函数 开始编程曲率判分", sJiBie, sZuBie, "总人数：", nStudentCount)
    # 按当前级别编程题成绩排序
    dfJiBieZuBie = dfJiBieZuBie.sort_values(by='编程总分')

    # 定义当前级别最终分数序列
    npCurveScore = np.zeros(dfJiBieZuBie.shape[0])
    # i 为学生序号当前值，j 为最终分数当前值
    i = 0
    j = 0
    for j in range(paraLowScore, paraHighScore, paraScoreStep) :
        # tmpPercentage 为当前分数j在正态后所处的百分比
        tmpPercentage = stats.norm.cdf(j, paraMu, paraSigma)
        if isDebug : print("Debug Info ->", "进入新分数段", "序号", i, "成绩", j, round(stats.norm.cdf(j, paraMu, paraSigma) * 100, 1), "%")
        # 如果当前学生排名位置小于此百分比，意即学生应该得此分数
        while ((i+1)/nStudentCount) < tmpPercentage :
            if isDebug : print("Debug Info ->", "填充分数", "序号", i, "成绩", j)
            npCurveScore[i] = j
            i += 1
        # 除第一个学生外，超过此百分比，但和前一名学生分数一样的情况，也赋予同样分数
        if i > 0 :
            while (dfJiBieZuBie.iloc[i]['编程总分'] == dfJiBieZuBie.iloc[i-1]['编程总分']) :
                if isDebug : print("Debug Info ->", "成绩相同并列填充", "序号", i, "成绩", j)
                npCurveScore[i] = j
                i += 1
                # break的目的是避免i超出学生总数，在上面while语句判断时产生下标溢出
                if i>=nStudentCount: break
        if i>=nStudentCount : break
    # 如果上述计算结束，仍有部分学生没有分数，均赋予并列最高分
    if isDebug : print("Debug Info ->", i, "正态分布99%外获得最高分的学生人数：", nStudentCount-i, "分数：",j)
    while i < nStudentCount :
        if isDebug : print("Debug Info ->", "正态99%外填充", "序号", i, "成绩", j)
        npCurveScore[i] = j
        i += 1

    # 定义当前级别组别的全国%成绩
    npPercentageScore = np.zeros(dfJiBieZuBie.shape[0])
    # 分数最高的学生，%成绩永远是99%，不论本组有多少人
    npPercentageScore[nStudentCount-1] = 0.99
    for i in range(nStudentCount-2, -1, -1) :
        if npCurveScore[i] == npCurveScore[i+1] :
            npPercentageScore[i] = npPercentageScore[i+1]
        else :
            npPercentageScore[i] = math.floor((i+1)/nStudentCount*100)/100
            if npPercentageScore[i]==0.00 : npPercentageScore[i] = 0.01

    # 计算需要填写的两列在表格中的列序号
    index1 = list(dfFinal.columns).index('第二部分成绩')
    index2 = list(dfFinal.columns).index('第二部分全国%')
    # 成绩表排序
    dfFinal = dfFinal.sort_values(by=['级别', '组别', '编程总分'])
    # 组合选择题成绩进入原成绩数据
    i = 0
    # 跳过所有非本级别
    while dfFinal.iloc[i]['级别'] != sJiBie :
        i += 1
    # 跳过所有非本组别
    while dfFinal.iloc[i]['组别'] != sZuBie :
        i += 1
    j = i
    # 填写本级别本组别第二部分成绩数据
    while dfFinal.iloc[i]['级别'] == sJiBie and dfFinal.iloc[i]['组别'] == sZuBie:
        # print("Debug Info ->", i, j, dfFinal.iloc[i]['级别'])
        dfFinal.iloc[i, index1] = npCurveScore[i-j]
        dfFinal.iloc[i, index2] = format(npPercentageScore[i-j], ".0%")
        i += 1
        if i-j>=nStudentCount: break
    return 0

# 本函数功能 - 计算所有人总分，并写入全局变量dfFinal
def TotalScore() :
    global dfFinal
    global nTotal

    print("Program Info ->", "进入TotalScore()函数，计算总成绩", "答案总人数:", nTotal)

    # 计算需要填写的列在表格中的列序号
    index1 = list(dfFinal.columns).index('总成绩')

    for i in range(nTotal) :
        dfFinal.iloc[i, index1] = int(dfFinal.iloc[i]['第一部分成绩']) + int(dfFinal.iloc[i]['第二部分成绩'])

    return 0

# 本函数功能 - 计算本级别本组别的全国百分比成绩
def TotalPercentage(sJiBie, sZuBie) :
    global dfFinal
    # 是否打开调试输出
    isDebug = False

    # 从总体数据中取出当前级别的成绩数据
    dfJiBieZuBie = dfFinal[dfFinal["级别"] == sJiBie]
    # 进一步选出当前组别的成绩数据
    dfJiBieZuBie = dfJiBieZuBie[dfJiBieZuBie["组别"] == sZuBie]
    # 核算当前级别组别总人数
    nStudentCount = len(dfJiBieZuBie)
    print("Program Info ->", "进入TotalPercentage()函数，计算全国百分比成绩", sJiBie, sZuBie, "总人数：", nStudentCount)

    # 按当前级别总成绩排序
    dfJiBieZuBie = dfJiBieZuBie.sort_values(by='总成绩')

    # 定义当前级别组别的全国%成绩
    npPercentageScore = np.zeros(dfJiBieZuBie.shape[0])
    # 分数最高的学生，%成绩永远是99%，不论本组有多少人
    npPercentageScore[nStudentCount-1] = 0.99
    for i in range(nStudentCount-2, -1, -1) :
        if dfJiBieZuBie.iloc[i]['总成绩'] == dfJiBieZuBie.iloc[i+1]['总成绩'] :
            npPercentageScore[i] = npPercentageScore[i+1]
            if isDebug : print("Debug Info ->", "分数并列，与上一人同样百分比", "序列号", i, "百分比", round(npPercentageScore[i]*100), "%")
        else :
            npPercentageScore[i] = math.floor((i+1)/nStudentCount*100)/100
            # 做边界处理，避免向下取整之后，小于1%的成绩被写为0%
            if npPercentageScore[i]==0.00 : npPercentageScore[i] = 0.01
            if isDebug : print("Debug Info ->", "分数不同，计算新百分比", "序列号", i, "百分比", round(npPercentageScore[i]*100), "%")
    # 计算需要填写的列在表格中的列序号
    index1 = list(dfFinal.columns).index('总成绩全国%')

    # 成绩表排序
    dfFinal = dfFinal.sort_values(by=['级别', '组别', '总成绩'])
    # 组合选择题成绩进入原成绩数据
    i = 0
    # 跳过所有非本级别
    while dfFinal.iloc[i]['级别'] != sJiBie :
        i += 1
    # 跳过所有非本组别
    while dfFinal.iloc[i]['组别'] != sZuBie :
        i += 1
    j = i
    # 填写本级别本组别全国百分比成绩数据
    while dfFinal.iloc[i]['级别'] == sJiBie and dfFinal.iloc[i]['组别'] == sZuBie:
        if isDebug : print("Debug Info ->", "填写本级别本组别百分比成绩", i, j, dfFinal.iloc[i]['级别'])
        dfFinal.iloc[i, index1] = format(npPercentageScore[i-j], ".0%")
        i += 1
        if i-j>=nStudentCount: break

    return 0

# 本函数功能 - 计算本省本级别本组别的省内百分比成绩
def ProvincePercentage(sJiBie, sZuBie) :
    global dfFinal
    # 是否打开调试输出
    isDebug = False

    # 从总体数据中取出当前级别的成绩数据
    dfJiBieZuBie = dfFinal[dfFinal["级别"] == sJiBie]
    # 进一步选出当前组别的成绩数据
    dfJiBieZuBie = dfJiBieZuBie[dfJiBieZuBie["组别"] == sZuBie]
    # 核算当前级别组别总人数
    nStudentCount = len(dfJiBieZuBie)
    print("Program Info ->", "进入ProvincePercentage()函数，计算省内百分比成绩", sJiBie, sZuBie, "总人数：", nStudentCount)

    # 按当前级别省份、总成绩排序
    dfJiBieZuBie = dfJiBieZuBie.sort_values(by=['省份', '总成绩'])

    # 定义当前级别组别的省内%成绩
    npPercentageScore = np.zeros(dfJiBieZuBie.shape[0])
    # 获得省份列表与计数序列
    seriesShengFen = dfJiBieZuBie['省份'].value_counts()
    if isDebug :
        print("Debug Info ->", "本级别组别包括省份", seriesShengFen.size, "个")
        print(seriesShengFen)
    # i为当前级别组别省内第一名学生序号
    nShengNeiFirst = 0
    # 每个i循环处理一个省份
    for i in range(seriesShengFen.size) :
        # 当前i循环所处理的省份
        sShengFen = dfJiBieZuBie.iloc[nShengNeiFirst]['省份']
        if isDebug : print("Debug Info ->", "当前处理省份：", sShengFen)
        # 当前省份学生数量
        nShengNeiCount = seriesShengFen[sShengFen]
        npPercentageScore[nShengNeiCount+nShengNeiFirst-1] = 0.99
        if isDebug : print("Debug Info ->", "当前省份第一人，成绩为99%", "序列号", nShengNeiCount+nShengNeiFirst-1, "百分比", "99%")
        for j in range(nShengNeiCount+nShengNeiFirst-2, nShengNeiFirst-1, -1) :
            # 每个j循环处理一个学生
            if dfJiBieZuBie.iloc[j]['总成绩'] == dfJiBieZuBie.iloc[j+1]['总成绩'] :
                # 分数与上一人并列
                npPercentageScore[j] = npPercentageScore[j+1]
                if isDebug : print("Debug Info ->", "分数并列，与上一人同样百分比", "序列号", j, "百分比", round(npPercentageScore[j] * 100), "%")
            else :
                # 分数与上一人不同
                npPercentageScore[j] = math.floor((j-nShengNeiFirst+1) / nShengNeiCount * 100) / 100
                # 做边界处理，避免向下取整之后，小于1%的成绩被写为0%
                if npPercentageScore[j] == 0.00 : npPercentageScore[j] = 0.01
                if isDebug : print("Debug Info ->", "分数不同，计算新百分比", "序列号", j, "百分比", round(npPercentageScore[j] * 100), "%")
        nShengNeiFirst += nShengNeiCount

    # 计算需要填写的列在表格中的列序号
    index1 = list(dfFinal.columns).index('总成绩省内%')

    # 成绩表排序
    dfFinal = dfFinal.sort_values(by=['级别', '组别', '省份', '总成绩'])
    # 组合选择题成绩进入原成绩数据
    i = 0
    # 跳过所有非本级别
    while dfFinal.iloc[i]['级别'] != sJiBie :
        i += 1
    # 跳过所有非本组别
    while dfFinal.iloc[i]['组别'] != sZuBie :
        i += 1
    j = i
    while dfFinal.iloc[i]['级别'] == sJiBie and dfFinal.iloc[i]['组别'] == sZuBie :
        if isDebug : print("Debug Info ->", "填写本级别本组别省内百分比成绩", "总序号", i, "组内起始", j, "成绩", format(npPercentageScore[i - j], ".0%"))
        dfFinal.iloc[i, index1] = format(npPercentageScore[i - j], ".0%")
        i += 1
        if i - j >= nStudentCount : break

    return 0

# 填写相应的蓝桥杯竞赛级别
def LanQiaoAward(top1:float, provincial1:float, provincial2:float, provincial3:float, provincial4:float) :
    global dfFinal
    global nTotal

    print("Program Info ->", "进入LanQiaoAward()函数，计算蓝桥对应奖项", "答案总人数：", nTotal)

    # 计算需要填写的列在表格中的列序号
    index1 = list(dfFinal.columns).index('蓝桥杯推荐')

    for i in range(nTotal) :
        if float(dfFinal.iloc[i]['总成绩省内%'].strip("%"))/100 >= 1-provincial1 :
            dfFinal.iloc[i, index1] = "省赛一等奖，推荐参加国赛"
        elif float(dfFinal.iloc[i]['总成绩省内%'].strip("%"))/100 >= 1-provincial2 :
            dfFinal.iloc[i, index1] = "地区选拔赛二等奖，推荐参加省赛"
        elif float(dfFinal.iloc[i]['总成绩省内%'].strip("%"))/100 >= 1-provincial3 :
            dfFinal.iloc[i, index1] = "地区选拔赛三等奖，推荐参加省赛"
        elif float(dfFinal.iloc[i]['总成绩省内%'].strip("%"))/100 >= 1-provincial4 :
            dfFinal.iloc[i, index1] = "地区选拔赛优秀奖，推荐参加省赛"
        if float(dfFinal.iloc[i]['总成绩全国%'].strip("%"))/100 >= 1-top1 :
            dfFinal.iloc[i, index1] = "TOP1%，省赛一等奖，推荐参加国赛"

    return 0

# 主程序起点----------------------------------------------------------------------

# 检查数据完整性
CheckData()
# 调用选择题判卷函数
MarkingChoice()

# 计算选择题最终发布成绩
dfFinal["第一部分成绩"] = 0
dfFinal["第一部分全国%"] = 0
CurveChoice("初级")
CurveChoice("中级")
CurveChoice("高级")

# 计算编程题最终发布成绩
dfFinal["第二部分成绩"] = 0
dfFinal["第二部分全国%"] = 0
CurveProgram("初级", "Python")
CurveProgram("初级", "Scratch")
CurveProgram("中级", "Python")
CurveProgram("中级", "Scratch")
CurveProgram("高级", "Python")
CurveProgram("高级", "Scratch")

# 计算总分最终发布成绩
dfFinal["总成绩"] = 0
dfFinal["总成绩全国%"] = 0
TotalScore()
TotalPercentage("初级", "Python")
TotalPercentage("初级", "Scratch")
TotalPercentage("中级", "Python")
TotalPercentage("中级", "Scratch")
TotalPercentage("高级", "Python")
TotalPercentage("高级", "Scratch")

# 计算分省百分比成绩
dfFinal['总成绩省内%'] = 0
ProvincePercentage("初级", "Python")
ProvincePercentage("初级", "Scratch")
ProvincePercentage("中级", "Python")
ProvincePercentage("中级", "Scratch")
ProvincePercentage("高级", "Python")
ProvincePercentage("高级", "Scratch")

# 计算等同蓝桥杯奖项
dfFinal['蓝桥杯推荐'] = ""
LanQiaoAward(0.01, 0.15, 0.30, 0.60, 0.80)

# 将判卷结果写入中间文件file3
print("写入分数评判文件...", file3)
dfFinal.to_excel(file3, sheet_name='阅卷结果')

# 将发布成绩写入file4
print("写入发布成绩文件...", file4)
excelWriter = pd.ExcelWriter(file4)
dfPublish = dfFinal[['姓名', '第一部分成绩', '第一部分全国%', '第二部分成绩', '第二部分全国%', '总成绩', '总成绩全国%', '总成绩省内%', '蓝桥杯推荐', '省份', '考点', '级别', '组别' ]]
dfPublish = dfPublish.sort_values(by=['级别', '组别', '省份', '总成绩'], ascending=False)
dfPublish.to_excel(excelWriter, sheet_name='发布成绩')
dfQuChu = pd.read_excel(file2, "去除项", index_col='准考证号')
dfQuChu.to_excel(excelWriter, sheet_name='去除项')
excelWriter.save()
excelWriter.close()

# 写入分析基础数据文件，供未来分析使用
print("写入发布成绩文件...", file5)
# 选取需要的列
lColumns = ['姓名', '省份', '考点', '级别', '组别', '第一部分成绩', '第一部分全国%', '第二部分成绩', '第二部分全国%', '总成绩', '总成绩全国%', '总成绩省内%']
for i in range (1, 73) :
    lColumns.append("答案"+str(i))
for i in range (1, 6) :
    lColumns.append("编程"+str(i))
dfPublish = dfFinal[lColumns]
# 写入文件
excelWriter = pd.ExcelWriter(file5)
dfPublish.to_excel(excelWriter, sheet_name='分析基础')
dfChuJiDaAn = pd.read_excel(file2, "初级组选择答案")
dfZhongGaoiDaAn = pd.read_excel(file2, "中高级组选择答案")
dfChuJiDaAn.to_excel(excelWriter, sheet_name='初级组选择答案')
dfZhongGaoiDaAn.to_excel(excelWriter, sheet_name='中高级组选择答案')
excelWriter.save()
excelWriter.close()

# 主程序终点----------------------------------------------------------------------