 Healthcare-Data-Analysis
# 1.Problem Statement
This project focuses on analyzing healthcare data to identify patient health trends, medical histories, hospitalization patterns, and healthcare costs. It involves data cleaning, transformation, exploration, and visualization to generate meaningful insights that support informed decision-making and improve healthcare management.
## ## 2. Dataset Description

The healthcare dataset contains **three related tables with 2,335 customer records**, covering customer information, medical examinations, and hospitalisation details. The tables are linked through the common **Customer ID** column, enabling comprehensive analysis of patient health conditions, medical history, hospitalisation patterns, and healthcare expenses.

### 1.2 Dataset Structure

| Worksheet | Rows | Columns | Description |
|---|---:|---:|---|
| Customer Names | 2,335 | 2 | Customer identification and names |
| Medical Examinations | 2,335 | 8 | Health conditions, examination results, and medical history |
| Hospitalisation Details | 2,343 | 9 | Hospitalisation records, hospital tiers, and healthcare charges |
| **Total** | **7,013** | **19** | **Combined healthcare dataset** |

### Key Data Attributes

- **Customer Information:** Customer IDs and names.
- **Medical Information:** BMI, HbA1c, heart issues, transplant history, cancer history, major surgeries, and smoking status.
- **Hospitalisation Information:** Admission dates, number of children, healthcare charges, hospital tiers, city tiers, and state IDs.

### Data Overview

- All 2,335 Customer IDs in the Customer Names and Medical Examinations tables are unique.
- The Medical Examinations table contains information on smoking habits, medical conditions, and surgery history.
- The Hospitalisation Details table contains 2,343 records, indicating that some customers may have multiple hospitalisation records.
- Hospital Tier 2 has the highest number of records (1,339).
- Healthcare charges range from ₹563.84 to ₹63,770.43.

### Analysis Objectives

The dataset supports analysis of:
- Patient health conditions and medical history.
- Smoking habits and their relationship with cancer history.
- BMI, HbA1c, and age-related health trends.
- Hospitalisation patterns across hospital and city tiers.
- Healthcare expenses and variations across patient groups.

### 1.3 Data Preprocessing
a) Check for the number of missing values marked with '?' in each column of the “Medical 
Examinations” Table and "Hospitalization Details" Table.  


The `COUNTIF` function was used to identify and count the missing values represented by `?` in each column of the **Medical Examinations** and **Hospitalisation Details** tables.

**Formula used:**
```excel
=COUNTIF(A2:A2344,"~?")
```
The tilde (`~`) was used as an escape character to treat `?` as a literal character rather than a wildcard.

**Results:**

- **Hospitalisation Details:** Missing values were identified in the respective columns, with counts of `6, 2, 3, 0, 0, 0, 1, 1, 2`.
- **Medical Examinations:** Missing values were identified in the respective columns, with counts of `0, 0, 0, 0, 0, 0, 0, 2`.

These results indicate that the Hospitalisation Details table contains missing values across multiple columns, while the Medical Examinations table contains missing values in only one column (2 occurrences).


b)Fill in the missing values of ‘month’ with Sep and ‘year’ with its average rounded to the 
nearest integer.  

Steps:
for month
1. Insert a temporary column next to Month.
2. Enter the formula in the first data row.
3. Drag the formula down to all records.
4. Copy the resulting values.
5. Use Paste Special → Values to replace the original Month column.

for year
1.Calculate the average of the available values in the Year column using the AVERAGE() function.
2.Apply the ROUND() function to round the calculated average to the nearest integer.
3.Replace missing values in the Year column with the rounded average.


c)Determine the most frequently occurring values in the ‘smoker’, 'Hospital tier' and 'City 
tier' columns, and fill in the missing values accordingly. 

**Objective:** Identify the most frequently occurring value in the `Smoker`, `Hospital Tier`, and `City Tier` columns and replace missing values with the respective most frequent value.

**steps:**

1. Use the `COUNTIF` function to calculate the frequency of each category in the respective columns.

   **Smoker:**
  =COUNTIF(H2:H2336,"yes")

   **Hospital Tier:**
  =COUNTIF(W2:W2344,"tier - 3")
  =COUNTIF(W2:W2344,"tier - 2")
  =COUNTIF(W2:W2344,"tier - 1")

   **City Tier:**
  =COUNTIF(X2:X2344,"tier - 3")
  =COUNTIF(X2:X2344,"tier - 1")
  =COUNTIF(X2:X2344,"tier - 2")

2. Compare the counts for each category and identify the value with the highest frequency (Mode).
3. Use the `IF` function to replace missing values with the identified Mode while retaining existing values.
4. Replace `"tier - 3/tier - 2/tier - 1"`Most Frequent Value in each formula with the actual Mode identified using `COUNTIF`.
5. Drag the formulas down to apply them to all records.
6. Copy the resulting columns and use **Paste Special → Values** to update the original data if required.

**Result:** Missing values in the three columns are replaced with their respective most frequently occurring values, while existing values remain unchanged.
      


