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
      
d) If any 'State ID' values are missing, consider filling them with 'Unknown' or using another 
appropriate strategy.  

1. Used the `COUNTIF` function to calculate the frequency of each State ID value in the State ID column.
   
   Example:
   - `=COUNTIF(Y2:Y2344,"R1013")`
   - `=COUNTIF(Y2:Y2344,"R1012")`
   - `=COUNTIF(Y2:Y2344,"R1011")`
   - `=COUNTIF(Y2:Y2344,"R1014")`

2. Compared the occurrence counts of the State ID values to identify the most frequently occurring value.
3. Checked for missing values in the State ID column.
4. Replaced missing State ID values with **"Unknown"** to avoid assigning an incorrect State ID to those records.
5. Used a helper column with the following formula to identify missing values:
`=IF(Y2="","Unknown",Y2)`
6. Filled the formula down for all records, then copied the results and pasted them as values.


3.Data Transformation
1) Split the ‘names’ column in the “Customer Names” Table into 3 meaningful columns: 
‘Title’, ‘First Name’, and ‘Last Name’. 

**Expected Name Format:**

The original Names column follows this format:

`Last Name, Title. First Name`

Example: `Hawks, Mr. John`

**Steps:**

1. In the **Customer Names** table, created three new columns: **Title**, **First Name**, and **Last Name**.

2. Extracted the title (Mr, Mrs, Ms, etc.) using the following formula:

   `=LEFT(TRIM(MID(B2,FIND(",",B2)+1,LEN(B2))),FIND(".",TRIM(MID(B2,FIND(",",B2)+1,LEN(B2))))-1)`

3. Extracted the first name using the following formula:

  '=TRIM(LEFT(TRIM(RIGHT(B2,LEN(B2)-FIND(".",B2))),FIND(" ",TRIM(RIGHT(B2,LEN(B2)-FIND(".",B2)))&" ")-1))'

4. Extracted the last name using the following formula:

   `=LEFT(B2,FIND(",",B2)-1)`

5. Applied each formula to the remaining records by dragging the fill handle down.

6. Verified the extracted values to ensure that the Title, First Name, and Last Name were placed in their respective columns.

**Expected Output:**

| Original Name | Title | First Name | Last Name |
|---|---|---|---|
| Hawks, Mr. John | Mr | John | Hawks |

2)Convert the "NumberOfMajorSurgeries" column in the “Medical Examinations” Table to numerical data by replacing non-numeric characters with meaningful numerical values. 

**Steps:**

1. Insert a new column named `NumberOfMajorSurgeries_Cleaned` next to the original column.
2. Enter the following IF formula in the first data row:

   ```excel
   =IF(G2="No major surgery",0,VALUE(G2))
   ```
3. Press Enter to apply the formula.
4. Drag the fill handle down to apply the formula to all remaining rows.
5. Verify that:
   - `No major surgery` is converted to `0`.
   - Existing numerical values (1.00, 2.00, 3.00) are retained as numbers.
6. Copy the cleaned column.
7. Use **Paste Special → Values** to replace the formulas with their calculated values.

**Result:** The `NumberOfMajorSurgeries` column is converted into numerical data, with `No major surgery` replaced by `0` using the IF function.

3)Check for inconsistencies in the 'Heart Issues' and 'smoker' columns and propose 
corrective actions if necessary.
steps

1. Used the `PROPER` function to standardize the text formatting in the **Heart Issues** and **Smoker** columns.
2. Applied the following formula to the Heart Issues column:`=PROPER(A2)`
3. Applied the following formula to the Smoker column:`=PROPER(B2)`
4. Dragged the formulas down to apply them to all records.
5. Reviewed the standardized values to identify inconsistencies in capitalization, such as `YES`, `yes`, and `Yes`.
6. Replaced the original values with the corrected, consistently formatted values using Copy and Paste Values.
7. Verified the columns to ensure consistent text formatting.

4) Create a new column named “Weight Status” that categorizes BMI into different 
categories as below:  
BMI 
Weight Status 
Below 18.5 -Underweight 
18.5 – 24.9 - Normal Weight 
25.0 – 29.9  - Overweight
30.0 and Above Obesity 

steps

1. Created a new column named **Weight Status** in the Medical Examinations table.
2. Used nested `IF` conditions to categorize BMI values into four weight-status categories based on the specified BMI ranges.
3. Entered the following formula, assuming the BMI value is in cell `M3`:
 `=IF(M3<18.5,"Underweight",IF(M3<25,"Normal Weight",IF(M3<30,"Overweight","Obesity")))`
4. Applied the formula to all records by dragging the fill handle down.
5. Verified the categorized results against the specified BMI ranges.

**BMI Classification:**

| BMI Range      | Weight Status |
| -------------- | ------------- |
| Below 18.5     | Underweight   |
| 18.5 – 24.9    | Normal Weight |
| 25.0 – 29.9    | Overweight    |
| 30.0 and above | Obesity       |

**Result:** The new `Weight Status` column categorizes each customer's BMI into the appropriate weight-status category.

5)Create a new column named “Diabetes Status” and fill it as per the information given 
below:  
HbA1C 
Diabetes Status 
Below 5.7 Normal
5.7 – 6.4 Prediabetes 
6.5 and Above Diabetes 

steps

1. Created a new column named **Diabetes Status** in the Medical Examinations table.
2. Used nested `IF` conditions to categorize HbA1C values into three categories based on the specified ranges.
3. Entered the following formula, assuming the HbA1C value is in cell `O3`:`=IF(O3="","",IF(O3<5.7,"Normal",IF(O3<6.5,"Prediabetes","Diabetes")))`
4. Applied the formula to all records by dragging the fill handle down.
5. Verified the results against the specified HbA1C ranges.
6. Kept the Diabetes Status cell blank when the corresponding HbA1C value was missing.
**HbA1C Classification:**

| HbA1C Range | Diabetes Status |
|---|---|
| Below 5.7 | Normal |
| 5.7 – Below 6.5 | Prediabetes |
| 6.5 and Above | Diabetes |

6).Merge ‘year’, ‘month’ and ‘date’ columns in the “Hospitalization Details” Table into one 
column named ‘Date of Birth’ and format it in ‘DD-MMM-YYYY’ custom format. 
###  Combining Year, Month, and Date into Date of Birth

1. In the **Hospitalization Details** table, created a new column named **Date of Birth**.
2. Combined the existing `year`, `month`, and `date` columns into a single date value.
3. Since the month column contained text values (e.g., `Sep`), used the `DATE`, `DATEVALUE`, and `MONTH` functions to convert the month name into its corresponding numerical value.
4. Entered the following formula, assuming:
   - `Z3` contains the year.
   - `AA3` contains the month.
   - `AB3` contains the date.
`=DATE(VALUE(Z3),MONTH(DATEVALUE("1 "&AA3&" 2000")),VALUE(AB3))`

5. Applied the formula to all records by dragging the fill handle down.
6. Selected the **Date of Birth** column and applied the custom date format `DD-MMM-YYYY`.
7. Verified the resulting dates to ensure that the year, month, and date were combined correctly.

**Result:** The separate year, month, and date values were merged into a single Date of Birth column and displayed in the `DD-MMM-YYYY` format.

 7).Calculate the ‘Age’ of each customer based on their ‘Date of Birth’ and the date of 
collection of the dataset, which is 8thJune 2023. 
### . Calculating Customer Age

1. Created a new column named **Age** in the Hospitalization Details table.
2. Calculated the age of each customer based on their **Date of Birth**, using the dataset collection date, **8th June 2023**, as the reference date.
3. Entered the following formula, assuming `AH3` contains the customer's Date of Birth:
`=2023-YEAR(AH3)`
4. Applied the formula to all records by dragging the fill handle down.
5. Verified the calculated ages to ensure that the results were based on the year of birth and the dataset collection year.

**Result:** The Age column contains the calculated age of each customer based on the dataset collection year (2023).

8)Format ‘charges’ column as currency ($). 
### . Formatting the Charges Column as Currency

1. Selected the **Charges** column in the Hospitalization Details worksheet.
2. Navigated to the **Home** tab and opened the **Number Format** dropdown in the Number group.
3. Changed the existing **General** format to **Currency**.
4. Selected the US Dollar (`$`) currency symbol.
5. Set the decimal places to 2.
6. Verified that all charge values were displayed in the `$` currency format.

**Result:** The Charges column was successfully formatted as US Dollar currency (`$`) with two decimal places.



