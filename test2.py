import pandas

df = pandas.read_excel(r"C:\Users\TLiss\Desktop\test_data.xlsx")

dispdf = df.drop(columns=["col_3"])
print(dispdf)