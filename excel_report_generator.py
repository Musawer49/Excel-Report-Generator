import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import DataBarRule
from tkinter import ttk

def create_report(input_file, output_file):
    # Read the selected Excel file
    data = pd.read_excel(input_file)
    data.columns = data.columns.str.strip()

    # Check required columns
    required_columns = ["Data", "Product", "Category", "Quantity", "Price"]
    missing_columns = [column for column in required_columns if column not in data.columns]

    if missing_columns:
        raise ValueError(
            f"Missing columns: {', '.join(missing_columns)}"
        )

    # Convert the Data column into real dates
    if pd.api.types.is_numeric_dtype(data["Data"]):
        data["Date"] = pd.to_datetime(
            data["Data"],
            unit="D",
            origin="1899-12-30"
        )
    else:
        data["Date"] = pd.to_datetime(
            data["Data"],
            errors="coerce"
        )

    if data["Date"].isna().any():
        raise ValueError("Some dates in the Excel file could not be converted.")

    data.drop(columns=["Data"], inplace=True)

    # Calculate total sales for each row
    data["Total"] = data["Quantity"] * data["Price"]

    # Calculate summary values
    total_sales = data["Total"].sum()
    total_quantity_sales = data["Quantity"].sum()

    # Calculate sales by category
    category_sales = data.groupby("Category")["Total"].sum()
    category_report = category_sales.reset_index(name="Total Sales")

    # Calculate sales by product
    product_sales = data.groupby("Product")["Total"].sum()
    product_report = product_sales.reset_index(name="Total Sales")
    product_report = product_report.sort_values(
        "Total Sales",
        ascending=False
    )

    # Find top-selling product
    top_product = product_sales.idxmax()

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

    # Create summary DataFrame
    report = pd.DataFrame(
        [
            ["Total Quantity Sold", total_quantity_sales],
            ["Total Sales", total_sales],
            ["Top Product", top_product],
        ],
        columns=["Sales", "Values"]
    )

    # Create Excel workbook
    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:
        
        # Write all sheets
        report.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )
        category_report.to_excel(
            writer,
            sheet_name="Category Sales",
            index=False
        )
        product_report.to_excel(
            writer,
            sheet_name="Product Sales",
            index=False
        )
        monthly_sales.to_excel(
            writer,
            sheet_name="Monthly Sales",
            index=False
        )
        data.to_excel(
            writer,
            sheet_name="Raw Data",
            index=False
        )

        # Common formatting
        header_fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )
        header_font = Font(
            color="FFFFFF",
            bold=True
        )
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        # SUMMARY SHEET
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

        # CATEGORY SALES SHEET
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

        for row in category.iter_rows(
            min_row=2,
            min_col=2,
            max_col=2
        ):
            row[0].number_format = "#,##0"

        # Category sales chart
        category_chart = BarChart()
        category_chart.title = "Sales by Category"
        category_chart.y_axis.title = "Total Sales"
        category_chart.x_axis.title = "Category"

        category_data = Reference(
            category,
            min_col=2,
            min_row=1,
            max_row=category.max_row
        )
        category_names = Reference(
            category,
            min_col=1,
            min_row=2,
            max_row=category.max_row
        )

        category_chart.add_data(
            category_data,
            titles_from_data=True
        )
        category_chart.set_categories(category_names)
        category.add_chart(category_chart, "D2")

        # PRODUCT SALES SHEET
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

        for row in product.iter_rows(
            min_row=2,
            min_col=2,
            max_col=2
        ):
            row[0].number_format = "#,##0"

        # Product sales chart
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
        product_names = Reference(
            product,
            min_col=1,
            min_row=2,
            max_row=product.max_row
        )

        product_chart.add_data(
            product_data,
            titles_from_data=True
        )
        product_chart.set_categories(product_names)
        product.add_chart(product_chart, "D2")

        # Product sales data bars
        product.conditional_formatting.add(
            f"B2:B{product.max_row}",
            DataBarRule(
                start_type="min",
                end_type="max",
                color="5B9BD5"
            )
        )

        # MONTHLY SALES SHEET
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

        for row in monthly.iter_rows(
            min_row=2,
            min_col=2,
            max_col=2
        ):
            row[0].number_format = "#,##0"

        # Monthly sales line chart
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
        monthly_names = Reference(
            monthly,
            min_col=1,
            min_row=2,
            max_row=monthly.max_row
        )

        monthly_chart.add_data(
            monthly_data,
            titles_from_data=True
        )
        monthly_chart.set_categories(monthly_names)
        monthly.add_chart(monthly_chart, "D2")

        # RAW DATA SHEET
        raw = writer.sheets["Raw Data"]

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

        # Date formatting
        for cell in raw["A"][1:]:
            cell.number_format = "dd-mmm-yyyy"

        # Convert Raw Data to Excel table
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

def select_input_file():
    file_path = filedialog.askopenfilename(
        title="Select Sales Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if file_path:
        input_file_var.set(file_path)
        input_label.config(text=file_path, fg=text_color)

        output_file_var.set(
            os.path.join(
                os.path.dirname(file_path),
                "sales_report.xlsx"
            )
        )

        update_generate_button()


def update_generate_button():
    if input_file_var.get():
        generate_button.config(state="normal")
    else:
        generate_button.config(state="disabled")

def generate_report():
    generate_button.config(state="disabled")

    input_file = input_file_var.get()
    output_file = output_file_var.get()

    status_label.config(
        text="Generating report...",
        fg=primary_color
    )

    root.update_idletasks()

    try:
        create_report(input_file, output_file)

        status_label.config(
            text="Report generated successfully!",
            fg="#16803C"
        )
        generate_button.config(state="normal")
        messagebox.showinfo(
            "Success",
            f"Report generated successfully!\nSaved as: {output_file}"
        )

    except PermissionError:
        status_label.config(
            text="Please close the report in Excel and try again.",
            fg="#C62828"
        )
        generate_button.config(state="normal")
        messagebox.showerror(
            "Permission Error",
            "Please close sales_report.xlsx if it is currently open."
        )

    except Exception as error:
        status_label.config(
            text="An error occurred.",
            fg="#C62828"
        )
        generate_button.config(state="normal")
        messagebox.showerror(
            "Error",
            str(error)
        )

# GUI
root = tk.Tk()
root.title("Excel Report Generator")
root.geometry("760x400")
root.resizable(False, False)
root.configure(bg="#F4F7FB")

input_file_var = tk.StringVar()
output_file_var = tk.StringVar()

# Colors
card_color = "#FFFFFF"
primary_color = "#1F4E78"
primary_hover = "#0B2D4D"
text_color = "#1E293B"
secondary_text = "#64748B"
border_color = "#D9E2EC"

# Main Card 
card = tk.Frame(
    root,
    bg=card_color,
    highlightbackground=border_color,
    highlightthickness=1
)
card.place(relx=0.5, rely=0.5, anchor="center", width=650, height=340)

# Title 
title_label = tk.Label(
    card,
    text="Excel Report Generator",
    font=("Segoe UI", 24, "bold"),
    bg=card_color,
    fg=primary_color
)
title_label.pack(pady=(25, 5))

subtitle_label = tk.Label(
    card,
    text="Generate a professional Excel sales report automatically",
    font=("Segoe UI", 11),
    bg=card_color,
    fg=secondary_text
)
subtitle_label.pack(pady=(0, 20))

# Input File 
input_title = tk.Label(
    card,
    text="Input Excel File",
    font=("Segoe UI", 10, "bold"),
    bg=card_color,
    fg=text_color
)
input_title.pack(anchor="w", padx=55)

input_frame = tk.Frame(
    card,
    bg="#F8FAFC",
    highlightbackground=border_color,
    highlightthickness=1
)
input_frame.pack(fill="x", padx=55, pady=(5, 25))

input_label = tk.Label(
    input_frame,
    text="No file selected",
    font=("Segoe UI", 9),
    bg="#F8FAFC",
    fg=secondary_text,
    anchor="w"
)
input_label.pack(side="left", fill="x", expand=True, padx=12, pady=10)

input_button = tk.Button(
    input_frame,
    text="Browse",
    font=("Segoe UI", 9, "bold"),
    bg=primary_color,
    fg="white",
    activebackground=primary_hover,
    activeforeground="white",
    relief="flat",
    bd=0,
    cursor="hand2",
    padx=15,
    pady=7,
    command=select_input_file
)
input_button.pack(side="right", padx=5, pady=5)

# Generate Button 
generate_button = tk.Button(
    card,
    text="Generate Report",
    command=generate_report,
    bg=primary_color,
    fg="white",
    font=("Segoe UI", 11, "bold"),
    relief="flat",
    cursor="hand2",
    padx=20,
    pady=10
)
generate_button.pack(pady=(0, 15))

def generate_hover(event):
    generate_button.config(bg=primary_hover)

def generate_leave(event):
    generate_button.config(bg=primary_color)

generate_button.bind("<Enter>", generate_hover)
generate_button.bind("<Leave>", generate_leave)

status_label = tk.Label(
    card,
    text="Select an Excel file to generate your report.",
    font=("Segoe UI", 9),
    bg=card_color,
    fg=secondary_text
)
status_label.pack()
root.mainloop()
