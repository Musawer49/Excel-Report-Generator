import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import DataBarRule
import tkinter as tk
from tkinter import filedialog, messagebox

# Read the input Excel file
root = tk.Tk()
root.withdraw()

input_file = filedialog.askopenfilename(
    title="Select Sales Excel File",
    filetypes=[
        ("Excel files", "*.xlsx *.xls"),
        ("All files", "*.*")
    ]
)

if not input_file:
    messagebox.showinfo("Cancelled", "No file was selected.")
    exit()

try:
    data = pd.read_excel(input_file)
    data.columns = data.columns.str.strip()

    required_columns = [
        "Data",
        "Product",
        "Category",
        "Quantity",
        "Price"
    ]

    missing_columns = [
        column for column in required_columns
        if column not in data.columns
    ]

    if missing_columns:
        messagebox.showerror(
            "Invalid Excel File",
            f"Missing columns: {', '.join(missing_columns)}"
        )
        exit()

except Exception as error:
    messagebox.showerror(
        "Error",
        f"Could not read the Excel file.\n\n{error}"
    )
    exit()

data["Data"] = pd.to_datetime(
    data["Data"],
    unit="D",
    origin="1899-12-30"
)

data.rename(columns={"Data": "Date"}, inplace=True)

# Calculate total sales for each row
data["Total"] = data["Quantity"] * data["Price"]

# Calculate summary values
total_sales = data["Total"].sum()
total_quantity_sales = data["Quantity"].sum()

# Calculate sales by category
category_sales = data.groupby("Category")["Total"].sum()
category_report = category_sales.reset_index(name="Total Sales")

# Calculate monthly sales
monthly_sales = (
    data.groupby(data["Date"].dt.to_period("M"))["Total"]
    .sum()
    .reset_index()
)

monthly_sales["Date"] = monthly_sales["Date"].astype(str)
monthly_sales.rename(
    columns={"Date": "Month"},
    inplace=True
)

# Calculate sales by product
product_sales = data.groupby("Product")["Total"].sum()
product_report = product_sales.reset_index(name="Total Sales")
product_report = product_report.sort_values("Total Sales", ascending=False)

# Find the top-selling product
top_product = product_sales.idxmax()

# Create the Summary DataFrame
report = pd.DataFrame(
    [
        ["Total Quantity Sold", total_quantity_sales],
        ["Total Sales", total_sales],
        ["Top Product", top_product],
    ],
    columns=["Sales", "Values"]
)

# Create the Out-put location and filename
output_file = filedialog.asksaveasfilename(
    title="Save Sales Report",
    defaultextension=".xlsx",
    filetypes=[
        ("Excel files", "*.xlsx"),
        ("All files", "*.*")
    ],
    initialfile="sales_report.xlsx"
)

if not output_file:
    messagebox.showinfo("Cancelled", "Report was not saved.")
    exit()

# Create the Excel report
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    
    # Write all sheets
    report.to_excel(writer, sheet_name="Summary", index=False)
    category_report.to_excel(writer, sheet_name="Category Sales", index=False)
    product_report.to_excel(writer, sheet_name="Product Sales", index=False)
    data.to_excel(writer, sheet_name="Raw Data", index=False)
    monthly_sales.to_excel(writer, sheet_name="Monthly Sales", index=False)

    # Common formatting
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # =========================
    # SUMMARY SHEET
    # =========================
    summary = writer.sheets["Summary"]
    summary.column_dimensions["A"].width = 25
    summary.column_dimensions["B"].width = 20

    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    for row in summary.iter_rows():
        for cell in row:
            cell.border = thin_border

    summary["B2"].number_format = "#,##0"
    summary["B3"].number_format = "#,##0"

    # =========================
    # CATEGORY SALES SHEET
    # =========================
    category = writer.sheets["Category Sales"]
    category.column_dimensions["A"].width = 20
    category.column_dimensions["B"].width = 20

    for cell in category[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    for row in category.iter_rows():
        for cell in row:
            cell.border = thin_border

    for row in category.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "#,##0"

    # Category Sales chart
    chart = BarChart()
    chart.title = "Sales by Category"
    chart.y_axis.title = "Total Sales"
    chart.x_axis.title = "Category"

    data_ref = Reference(
        category,
        min_col=2,
        min_row=1,
        max_row=category.max_row
    )
    categories_ref = Reference(
        category,
        min_col=1,
        min_row=2,
        max_row=category.max_row
    )

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(categories_ref)
    category.add_chart(chart, "D2")

    # =========================
    # PRODUCT SALES SHEET
    # =========================
    product = writer.sheets["Product Sales"]
    product.column_dimensions["A"].width = 20
    product.column_dimensions["B"].width = 20

    for cell in product[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    for row in product.iter_rows():
        for cell in row:
            cell.border = thin_border

    for row in product.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "#,##0"

    # Product Sales chart
    product_chart = BarChart()
    product_chart.title = "Sales by Product"
    product_chart.y_axis.title = "Total Sales"
    product_chart.x_axis.title = "Product"

    product_data = Reference(
        product,
        min_col=2,
        min_row=1,
        max_row=product.max_row
    )
    product_categories = Reference(
        product,
        min_col=1,
        min_row=2,
        max_row=product.max_row
    )

    product_chart.add_data(product_data, titles_from_data=True)
    product_chart.set_categories(product_categories)
    product.add_chart(product_chart, "D2")

    # Data bars for product sales
    product.conditional_formatting.add(
        f"B2:B{product.max_row}",
        DataBarRule(
            start_type="min",
            end_type="max",
            color="5B9BD5"
        )
    )

    # =========================
    # MONTHLY SALES SHEET
    # =========================
    monthly = writer.sheets["Monthly Sales"]

    monthly.column_dimensions["A"].width = 20
    monthly.column_dimensions["B"].width = 20

    for cell in monthly[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    for row in monthly.iter_rows():
        for cell in row:
            cell.border = thin_border

    for row in monthly.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "#,##0"

    # Monthly Sales line chart
    monthly_chart = LineChart()
    monthly_chart.title = "Monthly Sales"
    monthly_chart.y_axis.title = "Total Sales"
    monthly_chart.x_axis.title = "Month"

    monthly_data = Reference(
        monthly,
        min_col=2,
        min_row=1,
        max_row=monthly.max_row
    )

    monthly_categories = Reference(
        monthly,
        min_col=1,
        min_row=2,
        max_row=monthly.max_row
    )

    monthly_chart.add_data(
    monthly_data,
    titles_from_data=True
    )

    monthly_chart.set_categories(monthly_categories)
    monthly.add_chart(monthly_chart, "D2")

    # =========================
    # RAW DATA SHEET
    # =========================
    raw = writer.sheets["Raw Data"]

    for cell in raw["A"][1:]:
        cell.number_format = "dd-mm-yyyy"

    for cell in raw[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    for row in raw.iter_rows():
        for cell in row:
            cell.border = thin_border

    raw.freeze_panes = "A2"

    raw.column_dimensions["A"].width = 15
    raw.column_dimensions["B"].width = 18
    raw.column_dimensions["C"].width = 18
    raw.column_dimensions["D"].width = 12
    raw.column_dimensions["E"].width = 15
    raw.column_dimensions["F"].width = 15

    # Convert Raw Data into an Excel table
    table_ref = f"A1:F{raw.max_row}"

    table = Table(
        displayName="RawDataTable",
        ref=table_ref
    )

    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    table.tableStyleInfo = style
    raw.add_table(table)
