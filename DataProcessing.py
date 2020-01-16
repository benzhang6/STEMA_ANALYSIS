import pandas as pd
import xlrd

file = "191215-总分.xlsx"
rb = xlrd.open_workbook(file)
sheet = rb.sheet_by_index(0) #表示Excel的第一个Sheet
nrows = sheet.nrows
print(nrows)

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





