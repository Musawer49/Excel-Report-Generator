# Excel Report Generator

A Python-based desktop application that automatically converts raw sales data from an Excel file into a professional, formatted sales report.

The project is built with **Python, Pandas, OpenPyXL, and Tkinter** and is designed to reduce repetitive Excel reporting work.

## Features

* Select an Excel sales file through a simple GUI
* Automatically clean and process the input data
* Calculate total sales
* Calculate total quantity sold
* Identify the top-selling product
* Generate category-wise sales
* Generate product-wise sales
* Generate monthly sales
* Create professional Excel charts
* Apply formatting, borders, column widths, and number formats
* Add Excel data bars for product sales
* Convert raw data into an Excel table
* Freeze headers in the Raw Data sheet
* Automatically save the generated report as `sales_report.xlsx`
* Handle missing columns and invalid files with user-friendly error messages
* Detect permission errors when the generated report is open

## Generated Report

The generated Excel workbook contains:

### 1. Summary

Provides key sales metrics such as:

* Total Quantity Sold
* Total Sales
* Top Product

### 2. Category Sales

Shows total sales for each category with a bar chart.

### 3. Product Sales

Shows sales by product, sorted by sales performance, with:

* Bar chart
* Conditional data bars

### 4. Monthly Sales

Shows sales performance by month with a line chart.

### 5. Raw Data

Contains the processed source data with:

* Proper formatting
* Date formatting
* Excel table
* Borders
* Frozen header row

## Input File Requirements

The current version is designed specifically for **sales Excel files** containing these columns:

```text
Data
Product
Category
Quantity
Price
```

The actual products and categories can be different for each user.

For example, one user may have:

```text
Electronics
Clothing
Food
```

while another may have:

```text
Laptops
Furniture
Stationery
```

The program processes the categories found in the user's data automatically.

## How It Works

```text
Excel Sales Data
       ↓
Select File
       ↓
Data Cleaning & Processing
       ↓
Sales Calculations
       ↓
Summary & Analysis
       ↓
Charts & Formatting
       ↓
Professional Excel Report
```

## Technologies Used

* **Python**
* **Pandas** — Data processing and analysis
* **OpenPyXL** — Excel workbook creation and formatting
* **Tkinter** — Desktop GUI
* **Git & GitHub** — Version control

## How to Run

Make sure Python 3.13 and the required libraries are installed.
Note: This project currently uses Python 3.13 because the Pandas environment 
encountered a compatibility/Application Control issue with Python 3.14 on the 
development system. Therefore, Python 3.13 is recommended for running 
the current version

Run the application with:

```bash
py -3.13 excel_report_generator.py
```

Then:

1. Click **Browse**
2. Select your sales Excel file
3. Click **Generate Report**
4. The generated report will be saved as `sales_report.xlsx` beside the selected input file.

## Current Limitations

The current version has some limitations:

* It does **not support every type of Excel file**.
* The input file must follow the required sales-data column structure.
* Column names must currently correspond to:
  `Data`, `Product`, `Category`, `Quantity`, and `Price`.
* Users cannot currently customize or map their own column names through the GUI.
* The application is focused on **sales reporting** rather than being a general-purpose Excel reporting tool.
* The report structure and analysis are currently predefined.

These limitations are intentional for the current version and provide opportunities for future improvements.

## Future Improvements

Possible future versions could include:

* Custom column mapping
* Support for different Excel data structures
* User-selectable report types
* Custom date ranges
* Additional financial metrics
* More chart types
* Custom report templates
* Export to PDF
* Improved validation
* More customizable dashboards
* Standalone `.exe` application

## Project Purpose

This project was created to practice and demonstrate practical skills in:

* Python automation
* Data processing with Pandas
* Excel automation with OpenPyXL
* GUI development with Tkinter
* Data analysis and visualization
* Error handling
* File handling
* Git and GitHub workflow

The main goal is to automate repetitive Excel reporting tasks and turn raw sales data into a structured and professional report with minimal user effort.

## Author

**Musawer Ahmed**

Computer Science Student | Python & Automation Enthusiast

---

⭐ If you find this project useful, feel free to explore the repository and follow the development of future Python automation projects.
