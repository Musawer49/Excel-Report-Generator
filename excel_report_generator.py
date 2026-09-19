import pandas as pd
data = pd.read_excel("sales_data.xlsx")
print(data)
data["Total"] = data["Quantity"] = data["Price"]
total_sales = data["Total"].sum()
print("Total Sales:", total_sales)