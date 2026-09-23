print("NEW VERSION RUNNING")
import pandas as pd
data = pd.read_excel("sales_data.xlsx")
print(data)
data["Total"] = data["Quantity"] * data["Price"]
total_sales = data["Total"].sum()

category_sales = data.groupby("Category")["Total"].sum()

product_sales = data.groupby("Product")["Total"].sum()
top_product = product_sales.idxmax()

total_quantity_sales = data["Quantity"].sum()
print(total_quantity_sales)

report = pd.DataFrame(
    [
        ["Total Quantity Sold", total_quantity_sales],
        ["Total Sales", total_sales],
        ["Top Productt", top_product],
    ],
    columns=["Sales", "Values"]
)
report.to_excel("sales_report.xlsx", index=False)
