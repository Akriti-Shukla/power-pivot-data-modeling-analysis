## Advanced Excel Project – Data Modeling & Analysis Using Power Pivot

This project demonstrates how to build a scalable **data model in Excel** using **Power Query**, **Power Pivot**, and **DAX**. The dataset includes one **fact table** (`Orders`) and two **dimension tables** (`Customers` and `Products`), enabling relational modeling and advanced analysis — all within Microsoft Excel.

---
![Sales trend of coffee](https://github.com/user-attachments/assets/47478bf3-81cf-4deb-b513-da0ec3dfa7e9)


![Profit Heatmap](https://github.com/user-attachments/assets/5634fe64-43cb-4398-9a22-eac6f90943d4)


### Project Structure

### 🗂 Tables Used:
- `Orders` (Fact Table): Includes Date, Customer ID, Product ID, Quantity
- `Customers` (Dimension Table): Includes Customer ID, Name, Segment, etc.
- `Products` (Dimension Table): Includes Product ID, Category, Sub-Category, Price, Profit

---

### 🔧 Step-by-Step Workflow

### 1. **Data Cleaning using Power Query**
- Loaded all datasets using Power Query
- Performed cleaning (e.g., renaming headers, removing nulls, changing data types)
- Loaded cleaned data into the Data Model

---

### 2. **Data Modeling with Relationships**
- Used **Diagram View in Power Pivot** to create relationships:
  - `Orders[Customer ID]` → `Customers[Customer ID]`
  - `Orders[Product ID]` → `Products[Product ID]`
- This forms the foundation of a **star schema** model

---

### 3. **Calculated Columns with DAX**

#### Calculated Profit per Order
```dax
=RELATED(Products[Profit]) * Orders[Quantity]
```

This pulls profit per unit from the Products table and multiplies it by the quantity in each order.

### Identifying Returning Customers

```dax
=CALCULATE(COUNTROWS(Orders), ALLEXCEPT(Orders, Orders[Customer ID])) > 1
```

### Dashboard and Visual Insights

Created various PivotTables and PivotCharts using measures from the Power Pivot model:

- Heatmap
Coffee Type vs. Yearly Sales

Highlights which products perform best each year

- Line Trend
Sales by Category over Time

Trend analysis across categories like Espresso, Cold Brew, etc.

- Returning Customer Analysis (In tabular form)
  
Used calculated column IsReturningCustomer to:
Count returning vs. new customers
Analyze average profit and quantity ordered

#### Created additional DAX Measures using the Power Pivot “New Measure” feature:

```dax
Profit per Order = [Total Profit] / [Order Count]
```

### Key Excel Features Used

- Power Query for data transformation and loading

- Power Pivot for relationships and star schema modeling

- DAX calculated columns and measures

- PivotTables and PivotCharts for reporting

- Date Table integration for time intelligence


![Returning Customer Data Analysis](https://github.com/user-attachments/assets/7ef42d3c-af61-4a11-9d32-cf23f50bd454)


### Conclusion and Outcome

This project helped me:

- Understand data modeling in Excel using Power Pivot

- Apply DAX to create meaningful insights like returning customers and profit metrics

- Design a star schema for scalable reporting

- Use Power Query and Power Pivot together for end-to-end data analysis in Excel
#### Project files and snapshots are attached for reference
