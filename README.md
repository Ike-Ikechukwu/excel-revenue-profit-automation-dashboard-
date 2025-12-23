# Excel Revenue & Profit Automation Dashboard
excel automation dashboard for revenue and profit analysis using Power Query, Pivot Tables, and light VBA.


## Overview
This project is an **Excel-based automated dashboard** for analyzing **revenue and profit** from the *Nigeria Retail and E-commerce Transactional Data*.
It automates data cleaning, consolidation, and reporting, and allows users to toggle between **Revenue** and **Profit** views using light VBA.

## Problem Statement
Monthly sales data often comes as seperate CSV files that require repeated manual cleaning, consolidation and analysis to track revenue and profit.
This slows reporting and increases the chance of errors, making it difficult for decision-makers to get timely insights.

The challenge is to create an **Excel-based solution** that **automates monthly revenue and profit analysis**, refreshes insight with minimal effort and presents results through a clear, decision-ready dashboard.

## Solution
An Excel-based automation dashboard was developed to streamline monthly revenue and profit analysis. The solution automatically consolidates transactional data, standardizes it for consistent reporting, and presents key performance metrics through an interactive dashboard.
  
  With a single refresh, the latest available data is automatically included in the analysis, updating revenue, profit, and margin insights end-to-end without repeated manual work. The design priortizes minimal manual effort, repeatable monthly analysis, and clear business insights, while using light interactivity to allow users to switch between revenue and profit views as needed.

## How it Works
1.  Drop new monthly files into the source folder
2.  Open the Excel dashboard file
3.  Click **Refresh All**
4.  Dashboard updates automatically
5.  Use the toggle buttons to switch between Revenue and Profit Views

## Dataset Overview
**Nigeria Retail and E-commerce Monthly Transactional Data**
  
  Columns :
  1.   Transaction id- Unique identifier of each sales transaction.
  2.   Transaction date- Date the transaction occured
  3.   Product Category- High level classification of the product
  4.   Product name- Specific product sold
  5.   Purchase City- Location where the transaction took place
  6.   Unit price- Selling Price per unit
  7.   Unit cost- Cost incurred per unit
  8.   Total price- Total revenue generated per tansaction
  9.   Total cost- Total cost associated with the transaction
  10.  Total profit- Net profit earned from the transaction

## Methodology


### Data Cleaning & Preparation
1.  **Data Collection:**
  -  Monthly transactional CSV files are gathered and placed in a source folder for consolidation.

2.  **Data Cleaning and Preparation:**
  -  Standardized column names and data types.
  -  Corrected errors in dates, product names, and numeric fields.
  -  Handled duplicates and missing values.
  -  Calculated Total Price, Total Cost, Profit, and Profit Margin.
3.  **Data Consolidation:**
   -  Combined multiple monthly files into a single dataset using Power Query.
   -  Built a dynamic data model for seamless inclusion of new data.
4.  **Analysis & KPI Calculation:**
   -  Defined KPIs: Total Transactions, Total Units Sold, Total Revenue, Total Cost, Total Profit, Average Revenue per Transaction, Profit Margin (%).
   -  Metrics prepared for both Revenue and Profit views.
5.  **Dashboard Design & Automation:**
   -  Created interactive dashboards with Pivot Tables and Pivot Charts.
   -  Added light VBA for toggling between Revenue and Profit views.
   -  Visualizations include monthly trends, MoM change, city and category performance, top/bottom products, and weekday/week type analysis.
6.  **Automation & Refresh:**
   -  Designed for repeatable monthly updates.
   -  New files in the source folder can update the entire dashboard with one refresh, ensuring all KPIs and visuals are current.

## Analysis & Visuals
The dashboard provides both Revenue and Profit views, allowing users to switch between perspectives depending on business needs.

### Dashboard Charts
-  **Monthly Trends & MoM Change:** Shows monthly performance based on revenue or profit, highlights the best and worst performing months, and displays the percentage increase or decrease to track growth trends over time.
-  **Revenue by Week Type:** Compares weekday versus weekend performance to understand buying behavior.
-  **Revenue by Weekday:** Breaks down revenue across days of the week to identify peak and low-performing days.
-  **City Performance:**  Analyzes revenue and profit contribution by city to identify strong and underperforming locations. Also Top five cities with the highest contribution where emphasized, with thier percentage contribution to overall performance displayed for easy comparison.
-  **Product Category Performance:** Evaluates category-level performance to support product mix and inventory decisions and top selling product category is highlighted to quickly identify the strongest revenue or profit driver.
-  **Top & Bottom 10 Products:** Identifies the highest and lowest performing products based on revenue or profit.

### Key KPIs
-  **Total Transactions:** Total number of completed sales transactions within the selected period.
-  **Total Cost:** Total cost incurred to generate the recorded sales.
-  **Total Units Sold:** Sum of all units sold across all transaction.
-  **Total Sales Revenue:** Total revenue generated from all sales transactions.
-  **Average Revenue per Transaction:** Average sales value generated per transaction.
-  **Total Profit:** Net profit generated after deducting total cost from total revenue.
-  **Profit Margin (%):** Percentage of revenue retained as profit after costs.  

## Tools
-  **Microsoft Excel**: Excel Serves as the core plaatform for analysis, automation, and dashboard delivery. it enables a solution that samll and medium businesses can adopt easily.
-  **Power Query**: Power Query is used to automate data consolidation and standardization. This allows the latest monthly data to be included in the analysis with a single refresh.
-  **Pivot Tables & Charts**: Pivot tales and chart power all KPIs and visuals in the dashboard. They ensure metrics update dynamically as the underlying data changes.
-  **VBA**: VBA is used only to enchanse usability by toggling the dashboards between Revenue and Profit views.

## Scalability & Use Cases
This solution is ideal for small and medium businesses with monthly datasets of 1000 to 50,000 rows, where Excel remains efficient and easy to maintain.
  
  A full refresh typically takes 30 seconds to 2 minutes, depending on file size and system resources. This time is worth it, as the refresh consolidates files, cleans data, recalculates KPIs, and updates the dashboard. Delivering a complete revenue and profit analysis in one step.
  
  The same workflow can be applied quarterly or yearly, depending on business needs.
  
  For larger datasets, a BI tool is recommended. A Power BI version of this automation is available here:

👉 [link]

## Skills Demonstrated
-  Excel automation with Power Query
-  Revenue and Profit analysis
-  KPI-drdashiven dashboard design
-  Pratical VBA for interactivity
-  Business-focused reporting and decision support

## Author
**Ikechukwu Ike**
Data Analyst| Excel•SQL•PowerBI

