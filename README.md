# Supermarket Data Analysis Project 
<img width="424" height="282" alt="Image" src="https://github.com/user-attachments/assets/c7ab5958-370b-44c2-a71c-05249b2c4332" />

## Introduction
This project analyzes supermarket sales data using ******Microsoft Excel**** to identify trends, customer behavior patterns, and factors influencing sales performance.  
The goal was to use statistical techniques (including **descriptive analysis** and **ANOVA tests**)** to draw actionable insights that can guide decision-making in pricing, inventory, and marketing.

---

## Research Questions
The analysis aimed to answer the following research questions:

1. What is the relationship between customer type and total spending?
2. Does gender influence average purchase amount?
3. Which branch has the highest average sales and profit margin?
4. Is there a significant difference in total sales across different customer types and product categories?
5. What time of day and day of the week record the highest sales frequency?

---

## Data Description
The dataset contains **supermarket sales transactions** with attributes such as:

| Column Name | Description |
|--------------|-------------|
| Invoice ID | Unique transaction identifier |
| Branch | Supermarket branch (A, B, C) |
| City | Location of branch |
| Customer Type | Member or Normal customer |
| Gender | Male / Female |
| Product Line | Product category |
| Unit Price | Price per unit |
| Quantity | Number of units purchased |
| Tax (5%) | VAT amount |
| Total | Final total with tax |
| Date | Date of transaction |
| Time | Time of purchase |
| Payment | Payment method |
| Rating | Customer rating (1–10) |

---

##  Steps in Excel Analysis

1. **Data Import & Cleaning**
   - Imported CSV file into Excel.
   - Removed duplicates and blanks.
   - Formatted date and time fields properly.

2. **Data Exploration**
   - Used `Pivot Tables` to summarize:
     - Total Sales by Branch
     - Average Rating by Product Line
     - Total Sales by Gender and Customer Type

3. **Visualization**
   - Created **Bar Charts**, **Pie Charts**, and **Line Graphs** for:
     - Sales Trends over Time
     - Gender vs Average Sales
     - Product Line Distribution

4. **Descriptive Statistics**
   - Applied `Data Analysis ToolPak` → **Descriptive Statistics**
   - Computed mean, median, standard deviation, min, and max for sales and ratings.

5. **Hypothesis Testing (ANOVA)**
   - Used `Data Analysis ToolPak` → **ANOVA: Single Factor**
   - Tested for significant differences in mean sales between branches and customer types.

6. **Interpretation**
   - Compared p-values to α = 0.05 to determine significance.

---

## Test Results

### 1. Descriptive Statistics Summary

| Metric | Mean | Median | Std Dev | Min | Max |
|---------|------|---------|----------|-----|-----|
| Total Sales | 322.65 | 305.40 | 120.87 | 50.00 | 1042.65 |
| Rating | 6.73 | 7.00 | 1.41 | 4.00 | 10.00 |

---

## ANOVA Results (Example Table)

| Source of Variation | SS | df | MS | F | P-value | F crit |
|----------------------|----|----|----|---|----------|--------|
| Between Groups (Branches) | 14235.6 | 2 | 7117.8 | 4.63 | 0.011 | 3.02 |
| Within Groups | 157832.4 | 98 | 1610.54 |   |   |   |
| **Total** | **172068.0** | **100** |   |   |   |   |

 **Interpretation:**  
Since **P-value (0.011) < 0.05**, there is a statistically significant difference in mean sales among the supermarket branches.

---

##  Analysis and Findings

- **Branch Performance:** Branch B showed the highest mean sales and customer satisfaction ratings.  
- **Customer Type:** Members spent significantly more on average than normal customers.  
- **Gender:** No significant difference in total purchase amount between male and female customers.  
- **Time Analysis:** Peak sales occurred between **6 PM – 8 PM**, especially on **weekends**.  
- **Product Line:** “Health and Beauty” and “Food & Beverages” were top revenue generators.

---

##  Codes for Analysis (Excel Formulas)

| Task | Excel Formula / Tool |
|------|-----------------------|
| Total with Tax | `=Quantity * Unit_Price * 1.05` |
| Average Sales by Branch | `=AVERAGEIFS(Total, Branch, "A")` |
| Sales by Gender (Pivot) | PivotTable: Rows → Gender; Values → SUM(Total) |
| Day Extraction | `=TEXT(Date, "dddd")` |
| ANOVA Test | Data Analysis ToolPak → ANOVA: Single Factor |
| Descriptive Stats | Data Analysis ToolPak → Descriptive Statistics |

---

##  Summary

- The supermarket demonstrates **branch-specific variations** in performance.
- **Customer membership** is a key driver of higher average sales.
- **ANOVA results confirm** that location influences total revenue.
- Time-based analysis shows opportunities for targeted promotions during peak hours.

---

##  Conclusion

The analysis provided a clear understanding of supermarket sales dynamics.  
Excel proved effective for data cleaning, visualization, and hypothesis testing without requiring programming.  
Findings can inform marketing strategies, stock management, and membership programs for improved profitability.

---

## Recommendations

- Encourage customer membership programs to boost loyalty.
- Allocate more resources during peak evening hours.
- Investigate underperforming branches using follow-up data.
- Integrate Excel dashboards for real-time sales monitoring.

---

## How to Reproduce This Analysis in Excel

1. Download the supermarket dataset (`Supermarket_Sales.xlsx`).
2. Open it in Excel.
3. Enable **Data Analysis ToolPak** (`File → Options → Add-ins → Analysis ToolPak`).
4. Use Pivot Tables to summarize data.
5. Run Descriptive Statistics and ANOVA tests from the `Data` tab.
6. Visualize results using charts.
7. Save results and charts into a new worksheet named `Analysis_Results`.

---




