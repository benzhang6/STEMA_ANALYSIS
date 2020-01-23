import pandas as pd
import numpy as np
from scipy import stats

# 设置输入输出文件名
fileZongFen = "191215-总分.xlsx"
fileOutput1 = "tmp_output1.xlsx"
fileOutput2 = "tmp_output2.xlsx"

def curveChuji191215(fileINPUT, fileOUTPUT="") :
    # 本次考试是第几次STEMA考试
    timeoftest = 1
    paraTime = (timeoftest - 1) * 2
    paraDifficulty = 0
    # 最高1%的分数
    paraHighScore = 350 + paraTime + paraDifficulty
    # 最低1%的分数
    paraLowScore = 150
    # 平均分
    paraMu: int = (paraHighScore + paraLowScore) / 2
    paraSigma: int = 43

    print("Debug Info ->", "读取 dfChengji...")
    dfChengji = pd.read_excel(fileINPUT, sheet_name=0, index_col="准考证号")
    dfChuji = dfChengji[dfChengji["级别"] == "初级"]
    # 总人数
    paraStudentCount = len(dfChuji)
    print("Debug Info ->", "总人数：", paraStudentCount)
    dfChuji.sort_values(by='新计算选择总成绩')
    npCurveScore = np.zeros(dfChuji.shape[0])
    i = 0 # 人数增量
    for j in range(150, 370, 10) :
        tmpPercentage = stats.norm.cdf(j, paraMu, paraSigma)
        print("Debug Info ->", j, round(stats.norm.cdf(j, paraMu, paraSigma) * 100, 1), "%")
        while ((i+1)/paraStudentCount) < tmpPercentage :
            print("Debug Info ->", i, j)
            npCurveScore[i] = j
            i += 1
        if i > 0 :
            while (dfChuji.iloc[i]['新计算选择总成绩'] == dfChuji.iloc[i-1]['新计算选择总成绩']) :
                npCurveScore[i] = j
                i += 1
    while i < paraStudentCount :
        npCurveScore[i] = j
        i += 1

    dfChuji.insert(0, '选择发布成绩', npCurveScore)
    dfYuanShiFaBu = pd.read_excel("总分-4发布成绩.xlsx", sheet_name=0, index_col="准考证号")
    dfChuji = pd.merge(dfChuji, dfYuanShiFaBu, how="left", on="准考证号")

    for i in range(0, paraStudentCount) :
        print("Debug Info ->", i, dfChuji.iloc[i]['选择发布成绩'], dfChuji.iloc[i]['第一部分成绩'])

    print("writing excel file...")
    dfChuji.to_excel(fileOUTPUT, sheet_name='Result')

    return 0


def preprocessing191215(fileINPUT, fileOUTPUT) :
    # 读取发布总分表并与原始数据合并，以此刨除"去除项"
    print("Debug Info ->", "读取 dfZongFen...")
    dfZongfen = pd.read_excel(fileINPUT, sheet_name=0, index_col="准考证号", usecols='A,B')
    print("Debug Info ->", "读取 dfYuanShi...")
    dfYuanShi = pd.read_excel(fileINPUT, "原始成绩录入核算", index_col="准考证号")
    print("Debug Info ->", "合并 dfFinal...")
    dfFinal = pd.merge(dfZongfen, dfYuanShi, how="left", on="准考证号")

    # print("Debug Info ->", "合并之后表格长度：", len(dfFinal))

    # 读取答案表
    dfChujiDaan = pd.read_excel(fileINPUT, sheet_name="初级组正确答案")
    dfZhongjiDaan = pd.read_excel(fileINPUT, sheet_name="中高级组正确答案")

    # 定义新判分数，原分数是否正确两个列表
    npNewScore = np.zeros(dfFinal.shape[0])
    npIsCorrect = np.zeros(dfFinal.shape[0])

    # 开始判分
    for j in range(0, 1063) :
        nZongfen = 0
        if dfFinal.iloc[j]['级别'] == '初级' :
            for i in range(9, 73) :
                if dfFinal.iloc[j]['答案' + str(i)] == dfChujiDaan.iloc[0]['答案' + str(i)] :
                    nZongfen += 2
                elif str(dfFinal.iloc[j]['答案' + str(i)]) == "nan" or str(dfFinal.iloc[j]['答案' + str(i)]).strip() == "" :
                    pass
                else :
                    nZongfen -= 1
                # print("Debug Info ->", "题目", i, "学生答案", dfFinal.iloc[j]['答案' + str(i)], "正确答案", dfChujiDaan.iloc[0]['答案' + str(i)], "原分数", dfFinal.iloc[j]['成绩' + str(i)], "现总分", nZongfen)
        else :
            for i in range(9, 73) :
                if dfFinal.iloc[j]['答案' + str(i)] == dfZhongjiDaan.iloc[0]['答案' + str(i)] :
                    nZongfen += 2
                elif str(dfFinal.iloc[j]['答案' + str(i)]) == "nan" or str(dfFinal.iloc[j]['答案' + str(i)]).strip() == "" :
                    pass
                else :
                    nZongfen -= 1
                # print("Debug Info ->", "题目", i, "学生答案", dfFinal.iloc[j]['答案' + str(i)], "正确答案", dfZhongjiDaan.iloc[0]['答案' + str(i)], "原分数", dfFinal.iloc[j]['成绩' + str(i)], "现总分", nZongfen)

        if nZongfen != int(dfFinal.iloc[j]['选择总成绩']) :
            npIsCorrect[j] = False
        else :
            npIsCorrect[j] = True
        npNewScore[j] = nZongfen
        if not npIsCorrect[j] :
            print("Debug Info ->", "正在处理", j, "正确" if npIsCorrect[j] else "错误", "原成绩：", dfFinal.iloc[j]['选择总成绩'],
                  "现成绩：",
                  nZongfen)

    dfFinal.insert(0, '新计算选择总成绩', npNewScore)
    dfFinal.insert(0, '正确与否', npIsCorrect)

    print("writing excel file...")
    dfFinal.to_excel(fileOUTPUT, sheet_name='Result')
    return 0


curveChuji191215(fileOutput1, fileOutput2)
