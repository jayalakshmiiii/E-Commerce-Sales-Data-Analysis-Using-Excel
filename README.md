# E-Commerce Sales Data Analysis Using Excel

## 1. Project Overview

This project focuses on cleaning, transforming, analysing and visualising an e-commerce sales dataset using Microsoft Excel. The objective is to transform raw data into a structured dataset and generate meaningful business insights through statistical analysis, PivotTables and an interactive dashboard.

## 2. Tools Used

- Microsoft Excel
- Power Query
- XLOOKUP and VLOOKUP
- Analysis ToolPak
- PivotTables and PivotCharts
- Excel Slicers

## 3. Dataset Overview

The workbook contains four related datasets:

| Dataset | Records |
|---|---:|
| Customer_Dim | 500 |
| Product_Dim | 100 |
| Store_Dim | 20 |
| Sales_Fact | 2,000 |
| **Total** | **2,620** |

## 4. Data Cleaning and Transformation

**Tool Used: Microsoft Excel Power Query**

All four datasets were cleaned and transformed using Power Query before being loaded into Excel worksheets.

### 4.1 General Text Cleaning

Applied the following transformations to text columns across the datasets:

- **Trim:** Removed leading and trailing spaces.
- **Clean:** Removed unwanted non-printable characters.
- **Capitalize Each Word:** Standardised text capitalization for consistent formatting.

### 4.2 Customer_Dim

- Removed blank rows and duplicate records.
- Corrected inconsistent Customer_ID values using Replace Values.
- Applied Trim, Clean and Capitalize Each Word to text columns.
- Standardised the Country column by correcting inconsistent capitalization and replacing inconsistent values.
- Replaced missing Loyalty_Level values with "Unknown".

### 4.3 Product_Dim

- Corrected inconsistent Product_ID values, including P-3 and P-16.
- Removed the unnecessary Sub_Category column.
- Replaced missing Stock values using median imputation.
- Handled missing Product_Cost values using category-wise median imputation:
  1. Used Group By to calculate the median Product_Cost for each Product_Category.
  2. Created a custom column to replace null Product_Cost values with the corresponding category median.
  3. Preserved existing non-null values.

This approach ensured that missing product costs were replaced using the median of their respective categories rather than a single overall median.

### 4.4 Sales_Fact

Missing numerical values were handled using median imputation:

- **Discount:** Replaced missing values with the median Discount.
- **Unit_Price:** Replaced missing values with the median Unit_Price.
- **Quantity:** Replaced missing values with the median Quantity.

**Total_Amount Calculation**

Missing Total_Amount values were recalculated using a conditional custom column in Power Query:

```powerquery
if [Total_Amount] = null
then [Quantity] * [Unit_Price] * (1 - [Discount])
else [Total_Amount]
```

This ensured that only missing Total_Amount values were calculated using the sales formula, while existing values remained unchanged.

## 5. Table Combination

Used XLOOKUP and VLOOKUP to combine customer, product and store information with sales records.

Created a consolidated worksheet named `Final_Lookup`, containing:

- Customer_Name
- Product_Name
- Product_Category
- Brand
- Product_Cost
- Stock
- Store_Name
- Store_Region
- Store_City
- Store_Type

The original sales transaction details were retained.

## 6. Descriptive Statistical Analysis

Used Excel Analysis ToolPak to generate descriptive statistics for the `Total_Amount` column. The output was stored in a separate worksheet named `Descriptive Statistics`.

## 7. PivotTable Analysis

Created PivotTables to analyse:

1. Monthly and Yearly Revenue
2. Product Category-wise Revenue
3. Sales Revenue by Store Region
4. Top 10 Best-Selling Products by Quantity
5. Transaction Count by Payment Method
6. Average Discount by Product Category
7. Product Demand vs Average Selling Price

## 8. Data Visualization and Dashboard

Created an interactive Excel dashboard titled **E-Commerce Store Sales Dashboard** featuring:

- Seven PivotCharts using line, pie, column, bar, doughnut and combo chart types.
- Five KPI cards displaying Total Revenue, Average Selling Price, Total Quantity Sold, Total Transactions and Average Discount.
- Interactive slicers for Product Category, Store Region and Order Date (Year).

## 9. Project Outcome

Transformed raw e-commerce datasets into a cleaned and consolidated analytical workbook. The interactive dashboard provides insights into revenue trends, product demand, pricing, regional performance, discounts and payment preferences.

## 10. Workbook Structure

- Customer_Cleaned
- Product_Cleaned
- Store_Cleaned
- Final_Lookup
- Descriptive Statistics
- Pivot Tables
- Dashboard

---

**Project:** Data Cleaning, Analysis & Visualization with Excel

**Focus:** Data Cleaning | Power Query | Missing Value Imputation | Data Transformation | Statistical Analysis | Business Intelligence | Data Visualization
