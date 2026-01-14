# Excel Revenue & Profit Automation Dashboard
excel automation dashboard for revenue and profit analysis using Power Query, Pivot Tables, and light VBA.

## Table of Contents
- [Project Overview](#project-overview)
- [Problem Statement](#problem-statement)
- [Solution Overview](#solution)
- [How It Works](#how-it-works)
- [Dataset Overview](#dataset-overview)
- [Methodology](#methodology)
- [Tools Used](#tools)
- [Data Cleaning and Preparation](#data-cleaning-and-preparation)
- [Analysis and Dashboard Visuals](#analysis-and-dashboard-visuals)
- [Dashboard Insights](#dashboard-insights)
- [Key KPIs](#key-kpis)
- [Revenue Dashboard View](#revenue-dashboard-view)
- [Profit Dashboard View](#profit-dashboard-view)
- [Scalability and Use Cases](#scalability-and-use-cases)
- [Skills Demonstrated](#skills-demonstrated)
- [Author](#author)


## Project Overview
This project is an **Excel-based automated dashboard** for analyzing **revenue and profit** from the *Nigeria Retail and E-commerce Transactional Data*.
It automates data cleaning, consolidation, and reporting, and allows users to toggle between **Revenue** and **Profit** views using light VBA.


## Problem Statement
Organizations operating in retail and e-commerce environments receive monthly sales data. These files require repeated manual cleaning, consolidation, and reconciliation before meaningful revenue and profit analysis can be conducted. This manual process slows down reporting, introduces inconsistencies, and increases the risk of errors, making it difficult for stakeholders to access timely and reliable financial insights. 

As the volume of data grows over time, maintaining consistent, accurate, and scalable analysis becomes increasingly inefficient. Business leaders struggle to track month-over-month performance, evaluate city and product-level contributions, identify underperforming segments, and understand profitability drivers in a timely manner. The lack of an automated reporting framework limits the organization’s ability to make fast, data-driven decisions around pricing, cost control, and operational optimization. 

The business need is for a structured, Excel-based analytical solution that automates monthly data consolidation, revenue and profit calculations, and performance analysis across time, geography, and product categories. The solution must refresh with minimal manual intervention and present insights through a clear, interactive dashboard that supports financial monitoring, performance reviews, and strategic decision-making.


## Solution
An Excel-based automation dashboard was developed to streamline monthly revenue and profit analysis. The solution automatically consolidates transactional data, standardizes it for consistent reporting, and presents key performance metrics through an interactive dashboard.
  
With a single refresh, the latest available data is automatically included in the analysis, updating revenue, profit, and margin insights end-to-end without repeated manual work. The design priortizes minimal manual effort, repeatable monthly analysis, and clear business insights, while using light interactivity to allow users to switch between revenue and profit views as needed.


## How it Works
1.  Drop new monthly files into the source folder
2.  Open the Excel dashboard file
3.  Click **Refresh All**
4.  Dashboard updates automatically with the latest data
5.  Use the toggle buttons to switch between Revenue and Profit Views

[Watch the full demo on Linkedin](https://www.linkedin.com/posts/ikechukwu-ike-2745a4206_i-approached-this-project-by-starting-with-activity-7411784598705291264-gR52?utm_source=share&utm_medium=member_android&rcm=ACoAADRpongBUZAuCIAWn_JgaBcKLKwnzC_UpOE)

## Dataset Overview
The dataset used is **Nigeria Retail and E-commerce Monthly Transactional Data**, containing transactional level sales records used for revenue and profit analysis

[click to view all monthly datasets](data/raw)  
 
  **Columns :**
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

### Tools
-  **Microsoft Excel**: Excel Serves as the core plaatform for analysis, automation, and dashboard delivery. it enables a solution that samll and medium businesses can adopt easily.
-  **Power Query**: Power Query is used to automate data consolidation and standardization. This allows the latest monthly data to be included in the analysis with a single refresh.
-  **Pivot Tables & Charts**: Pivot tales and chart power all KPIs and visuals in the dashboard. They ensure metrics update dynamically as the underlying data changes.
-  **VBA**: VBA is used only to enchanse usability by toggling the dashboards between Revenue and Profit views.

### Data Cleaning and Preparation
1.  **Data Collection:**
  -  Monthly transactional CSV files are gathered and placed in a source folder for consolidation.
2.  **Data Cleaning and Preparation:** Before analysis, the data was cleaned to ensure accurate and consistent results.
  -  Data types were checked and corrected.
  -  Text fields were standardized to fix inconsistent letter casing
  -  Unnecessary spaces were removed.
  -  Missing values were reviewed and handled
  -  Duplicate records were removed to avoid double counting. 
3.  **Data Consolidation:**
   -  Multiple monthly files are combined into a single structured dataset using Power Query, enabling automated updates as new data is added  
4.  **Dashboard Design & Automation:**
   -  Created interactive dashboards with Pivot Tables and Pivot Charts.
   -  Added light VBA for toggling between Revenue and Profit views to improve usability.
5.  **Automation & Refresh:**
   -  Designed for repeatable monthly updates.
   -  With one refresh, the entire dashboard gets updated with New files in the source folder ensuring all KPIs and visuals are current.


## Analysis and Dashboard Visuals
The dashboard provides insights across Revenue and Profit views, allowing users to switch between perspectives depending on business needs.
### Dashboard Insights
-  **Monthly Trends & MoM Change:** Shows monthly performance based on revenue or profit, highlights the best and worst performing months, and displays the percentage increase or decrease to track growth trends over time.
-  **Revenue by Week Type:** Compares weekday versus weekend performance to understand buying behavior.
-  **Revenue by Weekday:** Breaks down revenue across days of the week to identify peak and low-performing days.
-  **City Performance:**  Analyzes revenue and profit contribution by city to identify strong and underperforming locations. Also Top five cities with the highest contribution where emphasized, with thier percentage contribution to overall performance displayed for easy comparison.
-  **Product Category Performance:** Evaluates category-level performance to support product mix and inventory decisions and top selling product category is highlighted to quickly identify the strongest revenue or profit driver.
-  **Top & Bottom 10 Products:** Identifies the highest and lowest performing products based on revenue or profit.

### Key KPIs
-  **Total Transactions:** Number of completed sales transactions.
-  **Total Cost:** Total cost incurred.
-  **Total Units Sold:** Total quantity of products sold.
-  **Total Sales Revenue:** Total revenue generated from all sales transactions.
-  **Average Revenue per Transaction:** Average sales value generated per transaction.
-  **Total Profit:** Net profit generated after deducting total cost from total revenue.
-  **Profit Margin (%):** Percentage of revenue retained as profit after costs.  

### Revenue Dashboard View
<img width="1826" height="890" alt="Screenshot 2025-12-23 122524" src="https://github.com/user-attachments/assets/1482c7b7-a5e6-4603-8a71-969a3b3053d5" />


### Profit Dashboard View
<img width="1836" height="892" alt="Screenshot 2025-12-23 122457" src="https://github.com/user-attachments/assets/a91b9bd2-d50d-481f-86af-d08f8b8ecbb5" />



## Scalability and Use Cases
This solution is ideal for small and medium businesses with monthly datasets of 1000 to 50,000 rows, where Excel remains efficient and easy to maintain.
  
  A full refresh typically takes 30 seconds to 2 minutes, depending on file size and system resources. This time is worth it, as the refresh consolidates files, cleans data, recalculates KPIs, and updates the dashboard. Delivering a complete revenue and profit analysis in one step.
  
  The same workflow can be applied quarterly or yearly, depending on business needs.
  
  For larger datasets, a BI tool is recommended.


## Skills Demonstrated
-  Excel automation with Power Query
-  Revenue and Profit analysis
-  KPI-drdashiven dashboard design
-  Pratical VBA for interactivity
-  Business-focused reporting and decision support


## Author
**Ikechukwu Ike**
Data Analyst| Excel•SQL•PowerBI

