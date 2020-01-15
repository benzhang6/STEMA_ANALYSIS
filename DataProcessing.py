import pandas as pd

df = pd.read_excel("成绩单.xlsx")
print(df)
print(df.shape)
print(df.dtypes)
print(df.columns)
print(df.head(1))
print(df.tail(1))
print(df["成绩"].unique())
print(df.isnull())
print(df.sort_values(by=["成绩"], ascending=False))





