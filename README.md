# Amazon Sales Performance Analysis
## Advanced Excel – Data Analyst Project

### 1. Introduction

E-commerce businesses generate large amounts of order data every day. Analysing this data helps sellers understand sales performance, customer purchasing behaviour, product demand, and order fulfilment.

In this project, I analysed Amazon sales transaction data using Microsoft Excel. The project covered data understanding, data cleaning, validation, preprocessing, PivotTable analysis, PivotCharts, and interactive dashboard development.

The main purpose was to convert raw order-level data into meaningful business information and help the seller understand which products, categories, and customer locations contribute most to sales.

### 2. Business Problem Statement

An Amazon seller provided three  months of order-level transactional data and wanted to understand overall business performance.

The seller faced the following business questions:

- Which products generate the highest sales?
- Which product categories contribute most to revenue?
- Which customer locations generate the highest and lowest sales?
- Which payment methods are commonly used by customers?
- How many orders are completed, pending, or cancelled?
- How does sales performance change over time?
- Where should the seller focus to improve business performance?

The challenge was to organize the raw data, analyse different business dimensions, and present the results in an interactive dashboard.

### 3. Project Objectives

The main objectives were:

1. Understand the structure and contents of the Amazon sales dataset.
2. Identify and handle missing, duplicate, and inconsistent records.
3. Standardize data formats and improve data quality.
4. Create calculated fields for time-based analysis.
5. Use PivotTables to summarize sales information.
6. Compare product, category, location, payment, and order-status performance.
7. Develop PivotCharts to present the analysis visually.
8. Build an interactive Excel dashboard using slicers and timelines.
9. Identify business insights and provide practical recommendations.

### 4. Dataset Overview

**Dataset:** Amazon Sales Dataset

**Source:** Amazon order-level sales dataset provided for the case study.

The dataset contains 250 order records and 12 columns.

| Column | Description |
|---|---|
| Order ID | Unique identification number for each order |
| Date | Date when the order was placed |
| Product | Name of the purchased product |
| Category | Product category |
| Price | Price per unit |
| Quantity | Number of units ordered |
| Total Sales | Total value of the order |
| Customer Name | Name of the customer |
| Customer Location | City of the customer |
| Payment Method | Method used for payment |
| Order Status | Completed, Pending, or Cancelled |
| Month | Month extracted from the order date |

**Dataset summary**

- Total records: 250
- Total columns: 12
- Total sales: ₹2,43,845
- Total quantity: 714 units
- Unique orders: 250
- Product categories: 5
- Unique products: 10
- Customer locations: 10
- Payment methods: 5
- Order statuses: 3

### 5. Data Cleaning

Data cleaning was an important step before performing the analysis. The objective was to make the dataset consistent and suitable for PivotTable reporting.

The documented data-cleaning workflow included:

**1. Duplicate checking**
- Checked for duplicate order records.
- Checked the uniqueness of Order IDs.
- Ensured duplicate records would not affect sales calculations.

**2. Missing value checking**
- Checked all columns for blank or missing values.
- Ensured important fields such as Order ID, Product, Quantity, and Total Sales were available.

**3. Data format standardization**
- Converted Date into a proper date format.
- Formatted Price, Quantity, and Total Sales as numeric values.
- Standardized text fields.

**4. Text cleaning**
- Used TRIM to remove unnecessary spaces.
- Used PROPER where appropriate to standardize text presentation.
- Checked category, product, and customer location names for consistency.

**5. Category and location standardization**
- Checked for inconsistent spellings.
- Ensured similar categories and locations were represented consistently.

**6. Data accuracy**
- Checked whether Total Sales matched Price multiplied by Quantity.

### 6. Data Validation

Data validation was performed to check whether the records were reliable for analysis.

The checks on the supplied dataset produced the following results:

| Validation check | Result |
|---|---:|
| Total records | 250 |
| Duplicate records | 0 |
| Duplicate Order IDs | 0 |
| Missing values | 0 |
| Sales calculation | Valid |
| Month and date consistency | Valid |

**Sales validation formula:**

`=Price*Quantity`

The calculated value was compared with the existing Total Sales column.

All 250 records matched the Price × Quantity calculation.

The month values were also consistent with the corresponding order dates.

**Result:** The supplied dataset passed these basic validation checks and was suitable for the analysis.

### 7. Data Preprocessing

After understanding and validating the dataset, additional fields were considered for better analysis.

The dataset already contained Total Sales, so there was no need to create another revenue column.

The following fields were used or considered:

**1. Month**

Used to identify the month of each transaction.

Example Excel formula:

`=TEXT(B2,"MMM")`

**2. Day of Week**

Can be calculated to identify which days receive more orders.

`=TEXT(B2,"DDDD")`

**3. Order Status Grouping**

The existing order status field was used to classify orders into:

- Completed
- Pending
- Cancelled

No profit or profit-margin calculation was performed because the dataset did not contain cost price or expense information.

### 8. Methodology

The project followed a step-by-step Excel analytics workflow.

1. **Data Loading:** Imported the Amazon sales dataset into Excel.
2. **Data Understanding:** Examined the columns, data types, categories, and business information.
3. **Data Cleaning:** Checked duplicates, missing values, inconsistent text, and formatting.
4. **Data Validation:** Verified order uniqueness, sales calculations, and date-month consistency.
5. **Data Preprocessing:** Used the Month field and other relevant calculated fields for analysis.
6. **PivotTable Analysis:** Summarized sales across products, categories, locations, payment methods, customers, and order statuses.
7. **PivotChart Development:** Converted the PivotTable summaries into visual charts.
8. **Dashboard Development:** Arranged KPI cards, charts, slicers, and date timelines into one interactive dashboard.
9. **Business Analysis:** Interpreted the results and prepared recommendations.

### 9. PivotTable Analysis

PivotTables were used to summarize the transactional data and answer the seller's business questions.

**1. Category-wise Sales Analysis**

- Rows: Category
- Values: Sum of Total Sales

Purpose: To identify which product categories contribute most to total sales.

**2. Product-wise Sales Analysis**

- Rows: Product
- Values: Sum of Total Sales

Purpose: To identify high-performing and low-performing products.

**3. Location-wise Sales Analysis**

- Rows: Customer Location
- Values: Sum of Total Sales

Purpose: To compare sales performance across different customer cities.

**4. Order Status Analysis**

- Rows: Order Status
- Values: Count of Order ID

Purpose: To understand the number of completed, pending, and cancelled orders.

**5. Payment Method Analysis**

- Rows: Payment Method
- Values: Count of Order ID

Purpose: To understand customer payment preferences.

**6. Customer-wise Sales Analysis**

- Rows: Customer Name
- Values: Sum of Total Sales

Purpose: To identify customers contributing more to total sales.

**7. Monthly Sales Analysis**

- Rows: Month
- Values: Sum of Total Sales and Count of Order ID

Purpose: To compare monthly sales and order activity.

### 10. PivotCharts and Dashboard Information

PivotCharts were created from the summarized PivotTable results to make the analysis easier to understand.

The dashboard contains the following visualizations:

**1. Total Sales by Category**

Displays sales performance across product categories and supports category-level comparison.

**2. Total Sales by Product**

A column chart comparing the sales generated by individual products.

**3. Total Sales by Location**

A bar chart showing sales contributions from different customer cities.

**4. Order Status Breakdown**

A doughnut chart showing the distribution of completed, pending, and cancelled orders.

**5. Payment Method Distribution**

A pie chart comparing the usage of different payment methods.

**6. KPI Cards**

The dashboard displays:

- Total Sales
- Total Orders
- Total Quantity
- Completed Orders
- Pending Orders
- Cancelled Orders

**7. Interactive Slicers**

The dashboard includes slicers for:

- Category
- Customer Location
- Order Status
- Month

**8. Date Timelines**

Date-based filters are provided for:

- Year
- Quarter
- Month
- Day

These controls help users filter the report and examine specific periods or business segments.

### 11. Time-Based Analysis

Time-based analysis was performed using the order Date and Month fields.

The objective was to understand how sales and order volumes changed over the available period.

The analysis included:

- Monthly total sales.
- Monthly order counts.
- Monthly quantity sold.
- Year-level filtering.
- Quarter-level filtering.
- Month-level filtering.
- Day-level filtering.

The available data covers February, March, and April 2025.

Monthly sales results:

| Month | Orders | Total Quantity | Total Sales |
|---|---:|---:|---:|
| February | 113 | 326 | ₹1,22,695 |
| March | 131 | 367 | ₹1,17,730 |
| April | 6 | 21 | ₹3,420 |
| **Total** | **250** | **714** | **₹2,43,845** |

March recorded the highest number of orders and quantity sold.

February generated the highest total sales.

April contains only six orders, so it should not be treated as a complete-month performance comparison.

### 12. Analysis Results and Key Findings

The following results were calculated from the 250 supplied records.

**A. Overall Sales Performance**

- Total sales: ₹2,43,845
- Total orders: 250
- Total quantity: 714 units
- Average sales per order: ₹975.38

**B. Category-wise Sales Performance**

| Category | Total Sales |
|---|---:|
| Electronics | ₹1,29,950 |
| Home Appliances | ₹1,05,000 |
| Footwear | ₹4,320 |
| Clothing | ₹3,540 |
| Books | ₹1,035 |

Electronics generated the highest sales, contributing approximately 53.3% of total sales.

Home Appliances was the second-highest category, contributing approximately 43.1%.

Together, these two categories contributed approximately 96.4% of total sales.

**C. Product-wise Sales Performance**

| Product | Total Sales |
|---|---:|
| Refrigerator | ₹78,000 |
| Laptop | ₹58,400 |
| Smartphone | ₹48,500 |
| Washing Machine | ₹27,000 |
| Smartwatch | ₹15,750 |
| Headphones | ₹7,300 |
| Running Shoes | ₹4,320 |
| Jeans | ₹2,480 |
| T-Shirt | ₹1,060 |
| Book | ₹1,035 |

Key observation:

Refrigerators generated the highest sales, followed by Laptops and Smartphones.

Books generated the lowest sales.

**D. Location-wise Sales Performance**

| Customer Location | Total Sales |
|---|---:|
| Miami | ₹31,700 |
| Denver | ₹29,785 |
| Houston | ₹28,390 |
| Dallas | ₹27,145 |
| Seattle | ₹26,890 |
| Boston | ₹26,170 |
| Chicago | ₹20,810 |
| New York | ₹18,940 |
| Los Angeles | ₹17,820 |
| San Francisco | ₹16,195 |

Key observation:

Miami generated the highest sales, while San Francisco recorded the lowest sales among the listed locations.

**E. Order Status Analysis**

| Order Status | Orders | Percentage |
|---|---:|---:|
| Completed | 88 | 35.2% |
| Pending | 85 | 34.0% |
| Cancelled | 77 | 30.8% |
| **Total** | **250** | **100%** |

Key observation:

Completed orders represent 35.2% of all orders.

Pending orders represent 34%, while cancelled orders account for 30.8%.

The high number of pending and cancelled orders requires attention from the seller.

**F. Payment Method Analysis**

| Payment Method | Orders | Percentage |
|---|---:|---:|
| PayPal | 60 | 24.0% |
| Credit Card | 54 | 21.6% |
| Debit Card | 53 | 21.2% |
| Gift Card | 42 | 16.8% |
| Amazon Pay | 41 | 16.4% |
| **Total** | **250** | **100%** |

Key observation:

PayPal was the most frequently used payment method, followed by Credit Card and Debit Card.

Amazon Pay recorded the lowest number of orders.

### 13. What Problems Did We Identify?

The analysis helped identify the following business concerns:

1. Sales are highly concentrated in Electronics and Home Appliances.
2. Some products generate significantly lower sales than others.
3. Sales contributions vary across customer locations.
4. Pending orders account for 34% of all orders.
5. Cancelled orders account for 30.8% of all orders.
6. Payment preferences are not equally distributed.
7. Monthly sales values and order volumes show different patterns.

These observations help the seller understand where business attention may be required.

### 14. What Did We Observe and What Did We Do?

| What we observed | What we did |
|---|---|
| Raw transactional records were difficult to analyse directly | Organized and prepared the data in Excel |
| Different products contributed different sales amounts | Created product-wise PivotTables and charts |
| Sales varied across categories | Compared category-wise sales |
| Customer locations showed different sales performance | Created location-wise sales summaries |
| Orders had different statuses | Analysed completed, pending, and cancelled orders |
| Customers used different payment methods | Created payment method summaries |
| Sales changed across months | Performed time-based analysis |
| Users needed to explore different business segments | Added slicers and date timelines to the dashboard |

### 15. Business Insights

**Insight 1: Electronics is the leading category**

Electronics contributes approximately 53.3% of total sales.

Business meaning: The seller receives a major share of sales from electronic products.

**Insight 2: Home Appliances is another major contributor**

Home Appliances contributes approximately 43.1% of total sales.

Business meaning: Products such as Refrigerators and Washing Machines are important contributors to overall sales.

**Insight 3: Refrigerators are the highest-selling product by sales value**

Refrigerators generated ₹78,000, the highest sales among the ten products.

Business meaning: The seller should pay attention to refrigerator stock availability and product demand.

**Insight 4: Miami is the strongest customer location**

Miami generated ₹31,700 in sales.

Business meaning: The seller can examine customer demand and promotional opportunities in high-performing locations.

**Insight 5: Pending and cancelled orders require attention**

There are 85 pending orders and 77 cancelled orders.

Business meaning: Order processing, customer communication, and cancellation reasons should be reviewed.

**Insight 6: PayPal is the most frequently used payment method**

PayPal accounts for 24% of orders.

Business meaning: The seller should monitor payment preferences and ensure a smooth checkout experience.

**Insight 7: February generated the highest sales**

February generated ₹1,22,695, while March recorded the highest order count.

Business meaning: Higher order volume does not always result in higher total sales.

### 16. Business Recommendations

Based on the findings, the following actions can be recommended:

1. **Focus on high-performing products:** Maintain sufficient stock of Refrigerators, Laptops, and Smartphones to support customer demand.

2. **Improve category planning:** Continue monitoring Electronics and Home Appliances because they contribute most of the sales.

3. **Improve low-performing product sales:** Review pricing, promotions, product visibility, and customer demand for Books, T-Shirts, and Jeans.

4. **Focus on location-based marketing:** Study successful locations such as Miami and Denver and identify opportunities to improve sales in San Francisco and Los Angeles.

5. **Reduce pending orders:** Review order processing delays and improve communication with customers.

6. **Understand cancellation reasons:** Investigate why customers cancel orders and take steps to reduce avoidable cancellations.

7. **Support popular payment methods:** Maintain a smooth payment experience for PayPal, Credit Card, and Debit Card users.

8. **Monitor monthly performance:** Compare sales, order counts, and quantities regularly to understand changes in business performance.

### 17. Conclusion

The Amazon Sales Performance Analysis project demonstrates how Microsoft Excel can be used to convert raw e-commerce transaction data into useful business information.

Through data cleaning, validation, preprocessing, PivotTables, PivotCharts, and interactive dashboard development, the project provides a structured view of product performance, category sales, customer locations, payment preferences, order statuses, and monthly sales trends.

The analysis identified Electronics and Home Appliances as the major sales contributors, Refrigerators as the highest-selling product by sales value, Miami as the leading customer location, and PayPal as the most frequently used payment method.

It also highlighted the need to monitor pending and cancelled orders.

Overall, the project helped develop practical skills in Excel-based data analysis, reporting, dashboard development, business understanding, and data-driven decision-making.

The final dashboard allows users to explore sales performance through interactive slicers and date timelines, making the report easier to understand and use for business decisions.

### 18. Tools and Skills Used

**Tool:** Microsoft Excel

**Excel concepts and techniques:**
- Data Cleaning and Formatting
- Data Validation
- Excel Formulas
- Calculated Fields
- PivotTables
- PivotCharts
- KPI Reporting
- Slicers
- Date Timelines
- Interactive Dashboard Development
- Business Insights and Recommendations

