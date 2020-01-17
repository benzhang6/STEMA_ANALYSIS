import pandas as pd

fileZongFen = "191215-总分.xlsx"
fileOutput = "tmp_output.xlsx"

print("读取 dfZongFen...")
dfZongFen = pd.read_excel(fileZongFen, sheet_name=0, index_col="准考证号", usecols="A,B")
print("读取 dfYuanShi...")
dfYuanShi = pd.read_excel(fileZongFen, "原始成绩录入核算", index_col="准考证号")
print("合并 dfFinal...")
dfFinal = pd.merge(dfZongFen, dfYuanShi, how="left", on="准考证号")

print("合并之后表格长度：", len(dfFinal))

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





