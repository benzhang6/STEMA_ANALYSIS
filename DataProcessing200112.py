# 本程序是2020年1月12日考试的阅卷程序

import pandas as pd
import numpy as np
import math
from scipy import stats

# 本次考试是第几次STEMA考试
# 第1次 2019 12 15
# 第2次 2020 01 12
paraNoOfTest = 2

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
paraSigma: int = 43

# 设置全局成绩文件名
file2 = "200112-2-整理去除项.xlsx"
file3 = "200112-3-分数评判.xlsx"
file4 = "200112-4-发布成绩.xlsx"

# 设置成绩处理数据结构
print("Debug Info ->", "读取 汇总答案...")
dfFinal = pd.read_excel(file2, "汇总答案", index_col="准考证号")
nTotal = len(dfFinal)
print("Debug Info ->", "核算 答案总人数：", nTotal)

# 选择题阅卷
def MarkingChoice() : # 本函数功能 - 选择题判卷
    global dfFinal
    # 读取答案表
    dfChujiDaan = pd.read_excel(file2, sheet_name="初级组选择答案")
    dfZhongjiDaan = pd.read_excel(file2, sheet_name="中高级组选择答案")

    # 定义新判分数，原分数是否正确两个列表
    npNewScore = np.zeros(dfFinal.shape[0])
    npIsCorrect = np.zeros(dfFinal.shape[0])

    #定义需要详细审查（即输出Debug信息）的学生序号
    lStudent = []

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
                    print("Debug Info ->", "题目", i, "学生答案", dfFinal.iloc[j]['答案' + str(i)], "正确答案", dfChujiDaan.iloc[0]['答案' + str(i)], "原分数", dfFinal.iloc[j]['选择' + str(i)], "现总分", nZongfen)
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
                    print("Debug Info ->", "题目", i, "学生答案", dfFinal.iloc[j]['答案' + str(i)], "正确答案", dfZhongjiDaan.iloc[0]['答案' + str(i)], "原分数", dfFinal.iloc[j]['选择' + str(i)], "现总分", nZongfen)
        #判分结果与原有Excel表中判分结果是否相同
        if nZongfen != int(dfFinal.iloc[j]['选择总分']) :
            npIsCorrect[j] = False
        else :
            npIsCorrect[j] = True
        #记录新计算选择题分数
        npNewScore[j] = nZongfen
        #显示新计算成绩与原成绩不一致的项目
        if not npIsCorrect[j] :
            print("Debug Info ->", "正在处理", j, "正确" if npIsCorrect[j] else "错误", "原成绩：", dfFinal.iloc[j]['选择总分'], "现成绩：", nZongfen)

    dfFinal.insert(0, '新计算选择总成绩', npNewScore)
    dfFinal.insert(0, '正确与否', npIsCorrect)

    return 0

def CurveChoice(sJiBie) : # 本函数功能 - 选择题核算发布分数，曲率计算
    global dfFinal
    # 是否打开调试输出
    isDebug = True
    # 从总体数据中取出当前级别的成绩数据
    dfJiBie = dfFinal[dfFinal["级别"] == sJiBie]
    # 核算当前级别总人数
    nStudentCount = len(dfJiBie)
    if isDebug : print("Debug Info ->", "进入CurveChoice()函数", sJiBie, "总人数：", nStudentCount)
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
            npPercentageScore[i] = math.ceil((i+1)/nStudentCount*100)/100

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

def CurveProgram(sJiBie, sZuBie) : # 本函数功能 - 编程题核算发布分数，曲率计算

    global dfFinal
    # 从总体数据中取出当前级别的成绩数据
    dfJiBieZuBie = dfFinal[dfFinal["级别"] == sJiBie]
    # 进一步选出当前组别的成绩数据
    dfJiBieZuBie = dfJiBieZuBie[dfJiBieZuBie["组别"] == sZuBie]
    # 核算当前级别组别总人数
    nStudentCount = len(dfJiBieZuBie)
    print("Debug Info ->", sJiBie, sZuBie, "总人数：", nStudentCount)
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
        print("Debug Info ->", "进入新分数段", "序号", i, "成绩", j, round(stats.norm.cdf(j, paraMu, paraSigma) * 100, 1), "%")
        # 如果当前学生排名位置小于此百分比，意即学生应该得此分数
        while ((i+1)/nStudentCount) < tmpPercentage :
            print("Debug Info ->", "填充分数", "序号", i, "成绩", j)
            npCurveScore[i] = j
            i += 1
        # 除第一个学生外，超过此百分比，但和前一名学生分数一样的情况，也赋予同样分数
        if i > 0 :
            while (dfJiBieZuBie.iloc[i]['编程总分'] == dfJiBieZuBie.iloc[i-1]['编程总分']) :
                print("Debug Info ->", "成绩相同并列填充", "序号", i, "成绩", j)
                npCurveScore[i] = j
                i += 1
                # break的目的是避免i超出学生总数，在上面while语句判断时产生下标溢出
                if i>=nStudentCount: break
        if i>=nStudentCount : break
    # 如果上述计算结束，仍有部分学生没有分数，均赋予并列最高分
    print("Debug Info ->", i, "正态分布99%外获得最高分的学生人数：", nStudentCount-i, "分数：",j)
    while i < nStudentCount :
        print("Debug Info ->", "正态99%外填充", "序号", i, "成绩", j)
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
            npPercentageScore[i] = math.ceil((i+1)/nStudentCount*100)/100

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

def TotalScore() : # 本函数功能 - 计算所有人总分，并写入全局变量dfFinal
    global dfFinal
    global nTotal
    # 计算需要填写的列在表格中的列序号
    index1 = list(dfFinal.columns).index('总成绩')

    for i in range(nTotal) :
        dfFinal.iloc[i, index1] = int(dfFinal.iloc[i]['第一部分成绩']) + int(dfFinal.iloc[i]['第二部分成绩'])

    return 0

def TotalPercentage(sJiBie, sZuBie) : # 本函数功能 - 计算本级别本组别的全国百分比成绩
    global dfFinal
    # 是否打开调试输出
    isDebug = True
    # 从总体数据中取出当前级别的成绩数据
    dfJiBieZuBie = dfFinal[dfFinal["级别"] == sJiBie]
    # 进一步选出当前组别的成绩数据
    dfJiBieZuBie = dfJiBieZuBie[dfJiBieZuBie["组别"] == sZuBie]
    # 核算当前级别组别总人数
    nStudentCount = len(dfJiBieZuBie)
    if isDebug : print("Debug Info ->", "计算全国百分比成绩", sJiBie, sZuBie, "总人数：", nStudentCount)

    # 按当前级别总成绩排序
    dfJiBieZuBie = dfJiBieZuBie.sort_values(by='总成绩')

    # 定义当前级别组别的全国%成绩
    npPercentageScore = np.zeros(dfJiBieZuBie.shape[0])
    # 分数最高的学生，%成绩永远是99%，不论本组有多少人
    npPercentageScore[nStudentCount-1] = 0.99
    for i in range(nStudentCount-2, -1, -1) :
        if dfJiBieZuBie.iloc[i]['总成绩'] == dfJiBieZuBie.iloc[i+1]['总成绩'] :
            npPercentageScore[i] = npPercentageScore[i+1]
            if isDebug : print("Debug Info ->", "分数并列，与上一人同样百分比", "序列号", i, "百分比", npPercentageScore[i], "%")
        else :
            npPercentageScore[i] = math.ceil((i+1)/nStudentCount*100)/100
            # 做边界处理，避免向上取整之后，大于99%的成绩被写为100%
            if npPercentageScore[i]==1.00 : npPercentageScore[i] = 0.99
            if npPercentageScore[i]==0.00 : npPercentageScore[i] = 0.01
            if isDebug : print("Debug Info ->", "分数不同，计算新百分比", "序列号", i, "百分比", npPercentageScore[i], "%")
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

# 主程序起点----------------------------------------------------------------------
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

# 将判卷结果写入中间文件file3
print("writing excel file...")
dfFinal.to_excel(file3, sheet_name='Result')
# 主程序终点----------------------------------------------------------------------

