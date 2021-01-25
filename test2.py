import pandas

df = pandas.read_excel(r"C:\Users\TLiss\Desktop\test_data.xlsx")

dispdf = df.loc[df["col_2"] == "data 2 2"]
print(dispdf)
dispdf = dispdf.drop(columns=["col_2"])
print(dispdf)