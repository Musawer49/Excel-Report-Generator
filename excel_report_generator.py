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

report_data = [
    ["Total Quantity Sold", total_quantity_sales],
    ["Total Sales", total_sales],
    ["Top Product", top_product],
]

for category, sales in category_sales.items():
    report_data.append([category + " Sales", sales])

report = pd.DataFrame(
    report_data,
    columns=["Sales", "Values"]
)

print(report)
with pd.ExcelWriter("sales_report.xlsx") as writer:
    report.to_excel(writer, sheet_name="Summary", index=False)
