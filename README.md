# Store-Sales-Dashboard
An end-to-end Excel Business Intelligence project that transforms raw customer, order, and product data into an interactive sales dashboard.
The project covers data cleaning using **Power Query**, data modeling using **Power Pivot**, and data analysis through **DAX measures, PivotTables, and PivotCharts**.

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Tools Used](#tools-used)
3. [Workflow](#workflow)
4. [Stage 1: Data Cleaning in Power Query](#stage-1-data-cleaning-in-power-query)
5. [Stage 2: Data Modeling in Power Pivot](#stage-2-data-modeling-in-power-pivot)
6. [Stage 3: PivotTables](#stage-3-pivottables)
7. [Stage 4: Final Dashboard](#stage-4-final-dashboard)
8. [Key Insights](#key-insights)
9. [Repository Structure](#repository-structure)
10. [How to Use](#how-to-use)

---

## 1. Project Overview

This project uses three datasets: **Customers, Orders, and Products** to analyze store sales performance over one year.

### Project Goals

- Clean and standardize raw data using Power Query.
- Build a Star Schema in Power Pivot with a dedicated Calendar table.
- Create DAX measures for key business metrics.
- Summarize data using PivotTables and PivotCharts.
- Build an interactive dashboard with slicers and a timeline.

### Headline KPIs

| Metric | Value |
|---|---:|
| Total Sales | $101.1M |
| Total Customers | 300 |
| Total Quantity | 108,120 |
| Total Orders | 19,594 |

### Dashboard Preview

<img width="1458" height="595" alt="Store Sales Dashboard" src="https://github.com/user-attachments/assets/2415d7f9-37d3-4518-9450-2bcfcc5079f0" />


---

## 2. Tools Used

| Tool | Purpose |
|---|---|
| Excel Power Query | Data cleaning and transformation |
| Excel Power Pivot | Data modeling, relationships, and DAX measures |
| PivotTables / PivotCharts | Data aggregation and visualization |
| Slicers / Timeline | Dashboard interactivity |

---

## 3. Workflow

**Raw Data → Power Query → Power Pivot (Star Schema) → PivotTables → Interactive Dashboard**

---

## 4. Stage 1: Data Cleaning in Power Query

Three queries were created in Power Query and loaded into the Data Model:

- **Customers:** Imported from a single CSV file.
- **Products:** Imported from a single CSV file.
- **Orders:** Combined from multiple monthly files stored in a folder.



### Cleaning Summary

| Table | Issue | Action Taken |
|---|---|---|
| Customers | Inconsistent name casing and extra spaces | Applied proper case and trimmed text |
| Customers | Full name stored in one column | Split CustomerName into First Name and Last Name |
| Customers | Inconsistent Region formatting | Removed spaces and applied proper case |
| Products | Price contained the text "Two Hundred" | Replaced it with 200 and converted the column to Currency |
| Orders | Monthly files stored separately | Combined files using Folder.Files and a custom Transform File function |
| Orders | Blank or missing rows | Filtered out rows with missing ProductID or Quantity |
| Orders | Order IDs not unique across files | Replaced original OrderID with a sequential index |
| Orders | Mixed date formats | Standardized date components and created a unified New Date column |

The complete Power Query M code is available in:

[PowerQuery Applied Steps.docx](https://github.com/user-attachments/files/35157799/PowerQuery.Applied.Steps.docx)


---

## 5. Stage 2: Data Modeling in Power Pivot

The cleaned queries were loaded into the Excel Data Model and connected using a **Star Schema**.

Orders serves as the fact table, while Products, Customers, and Calendar serve as dimension tables.

### Data Model

<img width="991" height="583" alt="Star Schema" src="https://github.com/user-attachments/assets/be0fbb8d-ef68-4abc-8184-eac72af11867" />


### Tables

| Table | Type | Main Columns |
|---|---|---|
| Orders | Fact | OrderID, CustomerID, ProductID, Quantity, New Date |
| Products | Dimension | ProductID, ProductName, Category, Price |
| Customers | Dimension | CustomerID, First Name, Last Name, Region |
| Calendar | Dimension | Date, Year, Month Number, Month, MMM-YYYY, Day Of Week Number, Day Of Week, Date Hierarchy |

### Relationships

| From (One Side) | To (Many Side) | Relationship |
|---|---|---|
| Products[ProductID] | Orders[ProductID] | One-to-Many |
| Customers[CustomerID] | Orders[CustomerID] | One-to-Many |
| Calendar[Date] | Orders[New Date] | One-to-Many |

### DAX Measures

The following measures were created in the Orders table:

| Measure | Purpose |
|---|---|
| Total Sales | Calculates total revenue |
| Total Quantity | Calculates total units sold |
| Total Orders | Counts distinct orders |
| Total Customers | Counts distinct customers |

**DAX Formulas**

```dax
Total Sales :=
SUMX(
    Orders,
    Orders[Quantity] * RELATED(Products[Price])
)

Total Quantity :=
SUM(Orders[Quantity])

Total Orders :=
DISTINCTCOUNT(Orders[OrderID])

Total Customers :=
DISTINCTCOUNT(Orders[CustomerID])
```

`SUMX` iterates through the Orders table, multiplying the quantity in each order by the corresponding product price retrieved through `RELATED`. This allows sales to be calculated without adding the product price directly to the Orders table.

---

## 6. Stage 3: PivotTables

PivotTables were created directly from the Data Model to summarize the data and provide the foundation for the dashboard visuals.

### PivotTables Overview

<img width="1339" height="479" alt="Pivot Tables" src="https://github.com/user-attachments/assets/cd6e5baf-96b1-4f44-876f-e619c328646e" />


| PivotTable | Rows | Values | Dashboard Visual |
|---|---|---|---|
| KPI Cards | None | Total Sales, Total Customers, Total Quantity, Total Orders | KPI Cards |
| Sales by Category | Category | Total Sales (% of Grand Total) | Pie Chart |
| Orders by Region | Region | Total Orders (% of Grand Total) | Donut Chart |
| Orders by Weekday | Day Of Week | Total Orders | Column Chart |
| Orders by Month | Month | Total Orders | Line Chart |
| Top 5 Products | ProductName (Top 5 filter) | Total Sales | Horizontal Bar Chart |
| Bottom 5 Products | ProductName (Bottom 5 filter) | Total Sales | Horizontal Bar Chart |

---

## 7. Stage 4: Final Dashboard

PivotCharts, KPI cards, slicers, and a timeline were arranged on a dedicated dashboard sheet.

The slicers and timeline were connected to all relevant PivotTables, allowing the dashboard visuals to update together when filters are applied.

### Dashboard Components

- **KPI Cards:** Total Sales, Total Customers, Total Quantity, and Total Orders.
- **Orders Per Region:** Donut chart showing the percentage distribution of orders across the five regions.
- **Sales Per Product Category:** Pie chart showing the sales distribution across categories.
- **Orders Per Month:** Line chart showing monthly order trends.
- **Weekly Order Distribution:** Column chart comparing orders across weekdays.
- **Top 5 Best Selling Products:** Horizontal bar chart ranked by sales.
- **Bottom 5 Least Selling Products:** Horizontal bar chart ranked by sales.

### Dashboard Filters

- **Region:** Filter performance by geographical region.
- **Category:** Filter performance by product category.
- **Order Date:** Timeline filter for selecting the reporting period.

---

## 8. Key Insights

The dashboard highlights several findings about store performance:

- **Category Performance:** Clothing (28.8%) and Toys (25.1%) contribute more than half of total sales.
- **Regional Distribution:** Orders are relatively evenly distributed across the five regions, ranging from 18.7% to 20.7%.
- **Product Performance:** Doll, Shirt, and Pants are the top-selling products, while Sofa and Headphones generate the lowest sales.
- **Weekly Trends:** Orders peak on Thursday (2,901) and Sunday (2,882), while Wednesday records the lowest number of orders (2,696).
- **Monthly Trends:** Monthly orders remain relatively stable throughout the year, ranging from approximately 1,620 to 1,645 orders.

  ## 9. Repository Structure

```text
Sales-Insights-Excel-Dashboard/
│
├── README.md
│
├── data/
│   └── (raw data files)
│
├── power_query/
│   └── power_query_steps.txt
│
├── images/
│   ├── dashboard.png
│   ├── power_query.png
│   ├── pivot_table.png
│   └── star_schema.png
│
└── Sales_Dashboard.xlsx
```

---

## 10. How to Use

1. Clone or download this repository.
2. Open `Sales_Dashboard.xlsx` using Microsoft Excel.
3. Explore the dashboard using the Region and Category slicers and the Order Date timeline.
4. Review the Power Query transformations in `power_query/power_query_steps.txt`.
5. Explore the Data Model, relationships, and DAX measures through Power Pivot.


---

