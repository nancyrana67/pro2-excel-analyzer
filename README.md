Sales Data Analysis & Executive Dashboard Project

Welcome to the Sales Data Analysis project repository! This document outlines the structured workflow, formulas, and Excel techniques used to analyze a 200-row sales dataset (pro2(excel).xlsx).

🚀 Project Overview & Objectives

The goal of this project is to transform raw transaction data into actionable business intelligence through structured data analysis tasks:

High-Value Customer Lookup: Identifying top spenders using filtering, rankings, and INDEX/MATCH.

Multi-Dimensional Pivot Analysis: Analyzing total sales cross-tabulated by region and product category.

Data Visualization (KPI Charts): Building line, column, and pie charts to visualize trends and category shares.

Executive Dashboard: Summarizing key performance indicators (KPIs) and visual elements into a single canvas.

Growth Analytics: Adding dynamic up/down arrows (▲ / ▼) to track month-over-month sales growth.

📋 Task Breakdown & Implementation Guide

Task 1: Identify High-Value Customers

Filters: Enabled data filtering (Ctrl + Shift + L) on the main dataset and applied Number Filters > Top 10 on the Sales column to isolate top revenue-generating rows.

INDEX & MATCH Lookup: Created a ranked summary list of customers paired with a dynamic search cell (P2) to pull customer names instantly based on their rank:

=INDEX(L2:L6, MATCH(P2, K2:K6, 0))


Task 2: Pivot Table (Sales by Region & Product)

Configuration: Generated a Pivot Table on a dedicated worksheet mapping:

Rows: Region

Columns: Product_Category

Values: Sum of Sales

Task 3: Visualizing KPIs with Charts

Line Chart: Tracks monthly sales trajectory and seasonal performance over time.

Column Chart (Clustered Column): Compares performance metrics side-by-side across regions.

Pie Chart: Displays percentage contribution shares across the 5 product categories (Books, Clothing, Electronics, Furniture, Office Supplies).

Task 4: Executive Dashboard

Layout: Designed a dedicated Dashboard tab with gridlines turned off for a clean presentation canvas.

KPI Summary Cards: High-level top-line aggregators for Total Revenue, Total Profit, and Average Order Value.

Visual Integration: Embedded Pivot Charts and interactive Slicers (Region, Product_Category) for dynamic filtering.

Task 5: Monthly Sales Growth with Up/Down Arrows

Formula: Calculated month-over-month growth percentage in Column C:

=(B3-B2)/B2


Custom Number Formatting: Applied conditional custom cell styling to automatically toggle green up-arrows or red down-arrows:

[Color10]▲ 0.0%;[Red]▼ 0.0%;- 0.0%


🛠️ Tools & Technologies Used

Microsoft Excel: Pivot Tables, Formulas (INDEX, MATCH, SUM, AVERAGE, RANK), Conditional Formatting, Slicers, and Charting tools.
