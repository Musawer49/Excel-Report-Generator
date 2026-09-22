import pandas as pd
data = pd.read_excel("sales_data.xlsx")
print(data)
data["Total"] = data["Quantity"] * data["Price"]
total_sales = data["Total"].sum()
#print("Total Sales:", total_sales)
category_sales = data.groupby("Category")["Total"].sum()
# print("\nSales by Category: ", category_sales)
product_sales = data.groupby("Product")["Total"].sum()
top_product = product_sales.idxmax()
# print("\nTop-Selling Product: ", top_product)
report = data.copy()
report.to_excel("sales_report.xlsx", index=False)
print("\nReport generated successfully!")