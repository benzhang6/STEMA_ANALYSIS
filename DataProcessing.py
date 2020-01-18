import pandas as pd
import numpy as np

fileZongFen = "191215-总分.xlsx"
fileOutput = "tmp_output.xlsx"


print("读取 dfZongFen...")
dfZongFen = pd.read_excel(fileZongFen, sheet_name=0, index_col="准考证号", usecols="A,B")
print("读取 dfYuanShi...")
dfYuanShi = pd.read_excel(fileZongFen, "原始成绩录入核算", index_col="准考证号")
print("合并 dfFinal...")
dfFinal = pd.merge(dfZongFen, dfYuanShi, how="left", on="准考证号")

print("合并之后表格长度：", len(dfFinal))

#判卷
dfChujiDaan = pd.read_excel(fileZongFen, sheet_name="初级组正确答案")
dfZhongjiDaan = pd.read_excel(fileZongFen, sheet_name="中高级组正确答案")

npNewScore = np.zeros(dfFinal.shape[0])
npIsCorrect = np.zeros(dfFinal.shape[0])

for j in range(1059, 1063):
    nZongfen = 0
    for i in range(9, 73):
        if dfFinal.iloc[j]['级别'] == '初级':
            # print("DEBUG", "初级", nZongfen)
            if dfFinal.iloc[j]['答案'+str(i)] == dfChujiDaan.iloc[0]['答案'+str(i)] :
                nZongfen += 2
            elif dfFinal.iloc[j]['答案' + str(i)] == None or dfFinal.iloc[j]['答案' + str(i)] == "" :
                pass
            else:
                nZongfen -= 1
        else :
            #print("DEBUG", "中高级", nZongfen)
            if dfFinal.iloc[j]['答案' + str(i)] == dfZhongjiDaan.iloc[0]['答案' + str(i)]:
                nZongfen += 2
            elif dfFinal.iloc[j]['答案'+str(i)] == None or dfFinal.iloc[j]['答案'+str(i)] == "" :
               pass
            else :
               nZongfen -= 1
        print("DEBUG", dfFinal.iloc[j]['答案'+str(i)],dfZhongjiDaan.iloc[0]['答案'+str(i)],dfFinal.iloc[j]['成绩'+str(i)])

    if nZongfen != int(dfFinal.iloc[j]['选择总成绩']):
        npIsCorrect[j] = False
    else :
        npIsCorrect[j] = True
    npNewScore[j] = nZongfen
    print("正在处理", j, "正确" if npIsCorrect[j] else "错误", "原成绩：", dfFinal.iloc[j]['选择总成绩'], "现成绩：", nZongfen)

dfFinal.insert(0, "新计算选择总成绩", npNewScore)
dfFinal.insert(0, "正确与否", npIsCorrect)

print("writing excel file...")
dfFinal.to_excel(fileOutput, sheet_name='Result')

#dfTemp = pd.DataFrame()
#for i in range(0, len(dfFinal)):
#    #print(dfFinal.iloc[i]['姓名_x'], dfFinal.iloc[i]['姓名_y'])
#    if dfFinal.iloc[i]['组别_x'] != dfFinal.iloc[i]['组别_y']:
#        print(dfFinal.iloc[i])
#        print(dfFinal.iloc[i]['组别_x'], dfFinal.iloc[i]['组别_y'])
#        dfTemp = dfTemp.append(dfFinal.iloc[i])

#rb = xlrd.open_workbook(file)
#sheet = rb.sheet_by_index(0) #表示Excel的第一个Sheet
#nrows = sheet.nrows
#print(nrows)

#overall = pd.read_excel("191215-总分.xlsx", sheet_name="总分表", index_col="准考证号")
#print(overall)
# detail = pd.read_excel("191215-总分.xlsx", sheet_name="原始成绩录入核算", index_col="准考证号")

#print("The result is :" + overall.iloc[0].size)

"""
print(df.shape)
print(df.dtypes)
print(df.columns)
print(df.head(1))
print(df.tail(1))
print(df["成绩"].unique())
print(df.isnull())
# test comment
print(df.sort_values(by=["成绩"], ascending=False))
"""





