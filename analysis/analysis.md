# Sales Performance Analysis — MIS Dashboard (Clean Dataset)

This document explains the analysis performed on the cleaned sales dataset used in the MIS Sales Performance Dashboard.  
The dataset required no major cleaning operations, allowing direct focus on KPI generation, pivot modeling, and dashboard automation.

---

## 1. Dataset Status — Already Clean

The dataset was provided in a clean and analysis‑ready format.  
This means:

- No missing values
- No duplicate rows
- No inconsistent date formats
- No incorrect data types
- No negative or invalid sales/profit values
- No formatting issues

Because of this, the dataset could be used **directly** for pivot modeling and dashboard creation.

---

## 2. Validation Checks Performed

Even though the dataset was already clean, the following checks were performed to confirm data quality:

### ✔ Data Type Verification

- Dates were in proper Excel date format
- Sales, Profit, Quantity were numeric
- Region, Category, Sub‑Category were text fields

### ✔ Duplicate Check

- No duplicate Order IDs or rows were found

### ✔ Missing Value Check

- No blank or null values in key fields

### ✔ Outlier Scan

- No unrealistic values (e.g., negative sales, extreme discounts)

### ✔ Consistency Check

- Region names were consistent
- Category/Sub‑Category labels were standardized

These checks confirmed that the dataset was **clean and reliable**.

---

## 3. KPIs Generated

The dashboard uses the following KPIs:

- **Total Sales**
- **Total Profit**
- **Total Quantity Sold**
- **Average Discount**
- **Profit Margin %**

These KPIs help evaluate performance across regions, categories, and time periods.

---

## 4. Insights Derived

### **Regional Insights**

- Identify top‑performing and low‑performing regions
- Compare sales and profit distribution

### **Category Insights**

- Understand which product categories drive revenue
- Spot low‑margin or low‑demand categories

### **Time‑Series Trends**

- Month‑on‑Month (MoM) changes
- Quarter‑on‑Quarter (QoQ) patterns
- Year‑on‑Year (YoY) growth

---

## 5. Dashboard Interactivity

The dashboard includes:

- Slicers (Region, Category, Sub‑Category, Year, Quarter)
- Automated slicer connections using VBA
- Checkbox‑based control for PivotTables
- Dynamic filtering and drill‑down

This makes the dashboard highly interactive and user‑friendly.

---

## 6. Business Interpretation Summary

- Sales are concentrated in specific regions
- Certain categories consistently outperform others
- Discounts impact profit margins
- Seasonal trends influence sales volume
- High‑margin categories drive profitability

---

## Author

**Vishal Thakare**  
MIS & Data Analytics Dashboard Developer
