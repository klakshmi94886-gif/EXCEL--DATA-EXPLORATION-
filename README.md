# EXCEL--DATA-EXPLORATION-
# 📊 Excel Data Exploration & Analysis Project

## 🔍 Project Overview
This project focuses on data cleaning, text parsing, and exploratory data analysis (EDA) using **Microsoft Excel**. Working with the file **Excel Assignment 1 - Data Exploration LAKSHMI K.xlsx**, the objective was to transform raw transactional records into a structured, business-ready format to extract key operational insights.

The analysis answers critical business questions regarding catalog depth, inventory valuation, and category-specific performance metrics by leveraging text manipulation, logical functions, and conditional aggregations.

---

## 🛠️ Technical Skills & Excel Functions Used

* **Text Manipulation & Substring Extraction:** Applied string functions (`LEFT`, `MID`, `RIGHT`) to dissect complex, hyphenated alphanumeric `Product ID` codes into clean, standalone columns for structural analysis.
* **Logical Data Segmentation:** Used the conditional `IF` function to dynamically classify inventory based on pricing tiers.
* **Conditional Aggregations:** Developed summary statistics using criteria-based functions (`SUMIF`, `COUNTIF`) to isolate performance within specific industry verticals.
* **Descriptive Statistics:** Generated baseline data metrics using `SUM`, `COUNT`, `AVERAGE`, `MIN`, and `MAX`.

---

## 💾 Repository Structure & Dataset Architecture

The repository contains the primary workbook:
* **`Excel Assignment 1 - Data Exploration LAKSHMI K.xlsx`**
  * `Instructions` Sheet: Outlines the project criteria, business logic guidelines, and problem prompts.
  * `Dataset` Sheet: The primary analytical workspace containing 60 product records.

### Data Dictionary & Derived Logic

| Column Name | Data Type | Formula / Logic Applied | Description |
| :--- | :--- | :--- | :--- |
| **Product ID** | Text | *Raw Data* | Unique alphanumeric identifier (e.g., `28-JAN-US`). |
| **LEFT** | Text/Int | `=LEFT(A2, 2)` | Extracts the day component from the Product ID. |
| **MID** | Text | `=MID(A2, 4, 3)` | Extracts the month component from the Product ID. |
| **RIGHT** | Text | `=RIGHT(A2, 2)` | Extracts the geographical region code from the Product ID. |
| **Product Name**| Text | *Raw Data* | Generic product classification (e.g., Laptop, Sneakers). |
| **Brand Name**  | Text | *Raw Data* | Manufacturing brand identity. |
| **Price ($)**   | Currency | *Raw Data* | Retail unit price of the item. |
| **Price Range** | Category | `=IF(G2>=500, "High Price", "Standard Price")` | Dynamically tags items priced at or above $500 as 'High Price'. |
| **Quantity**    | Integer  | *Raw Data* | Current stock levels / units sold. |
| **Category**    | Category | *Raw Data* | Broad market category (e.g., Electronics, Fashion). |

---

## 🎯 Business Problems Solved

The workbook explicitly calculates and displays answers to the following core operational questions:

### 1. Catalog Diagnostics
* **Total Catalog Value:** Cumulative cost profile of all managed products (`SUM`).
* **Total Product Count:** Total volume of unique stock items tracked (`COUNT`).
* **Pricing Benchmarks:** Average catalog unit price (`AVERAGE`), alongside absolute pricing floors (`MIN`) and ceilings (`MAX`).

### 2. Segmented Portfolio Metrics
* **Inventory Tiering:** Classification of premium inventory vs. standard items to analyze revenue weight.
* **Targeted Category Performance:** Focused financial extraction for high-priority sectors (e.g., calculating total pricing volume for the **Electronics** category using `SUMIF`).

---

## 🚀 Getting Started

1. **Clone the Repository:**
```bash
   git clone [https://github.com/your-username/your-repository-name.git](https://github.com/your-username/your-repository-name.git)
