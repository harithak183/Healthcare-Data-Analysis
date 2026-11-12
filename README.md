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

4.Exploratory Data Analysis
Create a new sheet named “Healthcare", combine all three tables into one, using 
Customer ID as the common column, utilizing VLOOKUP. 

### 4.1. Combining the Three Tables Using VLOOKUP

1. Created a new worksheet named **Healthcare** in the Excel workbook.
2. Used **Customer ID** as the common column to combine data from the three existing tables:
   - Customer Names
   - Medical Examinations
   - Hospitalisation Details
3. Retained **Customer ID** as the first column in the new Healthcare worksheet.
4. Used the `VLOOKUP` function with an exact match (`0`) to retrieve the required information from the respective worksheets.

**VLOOKUP formulas used:**

**Customer Names table:**
- `=VLOOKUP(A2,'Customer Names'!A:D,4,0)` – First Name
- `=VLOOKUP(A2,'Customer Names'!A:E,5,0)` – Last Name

**Medical Examinations table:**
- `=VLOOKUP(A2,'Medical Examinations'!A:B,2,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:D,4,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:F,6,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:G,7,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:H,8,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:I,9,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:J,10,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:C,3,0)`
- `=VLOOKUP(A2,'Medical Examinations'!A:E,5,0)`

**Hospitalisation Details table:**
- `=VLOOKUP(A2,'Hospitalisation Details'!A:J,10,0)`
- `=VLOOKUP(A2,'Hospitalisation Details'!A:F,6,0)`
- `=VLOOKUP(A2,'Hospitalisation Details'!A:G,7,0)`
- `=VLOOKUP(A2,'Hospitalisation Details'!A:H,8,0)`
- `=VLOOKUP(A2,'Hospitalisation Details'!A:I,9,0)`
- `=VLOOKUP(A2,'Hospitalisation Details'!A:K,11,0)`

5. Used the appropriate column index in each formula to retrieve the required field from its source worksheet.
6. Applied the formulas to all customer records by dragging the fill handle down.
7. Retained the following required columns in the combined Healthcare table:

   - Customer ID
   - First Name
   - BMI
   - HBA1C
   - Heart Issues
   - Any Transplants
   - Cancer History
   - NumberOfMajorSurgeries
   - Smoker
   - Weight Status
   - Diabetes Status
   - Date of Birth
   - Charges
   - Hospital Tier
   - City Tier
   - State ID
   - Age

8. Verified the retrieved values against the original worksheets to ensure that the records were correctly matched using Customer ID.

**Result:** Successfully combined the required information from the three worksheets into a single **Healthcare** table using Customer ID as the common column and VLOOKUP with exact-match lookup.

4.2 Analysis using Pie/Donut Chart
1) What is the distribution of cancer history among smokers and non-smokers? 
### Distribution of Cancer History Among Smokers and Non-Smokers

**Step 1: Creating a PivotTable**

1. Opened the **Healthcare** worksheet containing the combined healthcare dataset.
2. Selected any cell within the dataset.
3. Navigated to the **Insert** tab in Excel for the Web.
4. Clicked **PivotTable**.
5. Selected the complete dataset as the source.
6. Chose **New Worksheet** and clicked **Create**.
7. Renamed the new worksheet to **Cancer History Analysis**.

**Step 2: Naming the PivotTable**

1. Selected any cell inside the PivotTable.
2. Navigated to the **PivotTable** tab.
3. Located the **PivotTable Name** field.
4. Entered the name `Pivot_Cancer_History` and pressed Enter.

**Step 3: Configuring PivotTable Fields**

1. In the PivotTable Fields pane, dragged the following fields into their respective areas:
   - **Rows:** Smoker
   - **Columns:** Cancer History
   - **Values:** Customer ID
2. Opened the Customer ID value field settings and selected **Count** to calculate the number of customers in each category.
3. Verified the PivotTable results to ensure that cancer history was distributed between smokers and non-smokers.

**Step 4: Creating a Pie Chart**

1. Selected the PivotTable results.
2. Navigated to the **Insert** tab.
3. Selected **Pie Chart** from the chart options.
4. Added the chart title: **Distribution of Cancer History Among Smokers and Non-Smokers**.
5. Verified the chart legend and data labels to ensure that the cancer history categories were clearly represented.

**Result:** Created a named PivotTable (`Pivot_Cancer_History`) and Pie Chart to analyze and visualize the distribution of cancer history among smokers and non-smokers based on the count of Customer IDs.


### Analysis: Major Surgeries and Average HbA1C by Transplant History

**Objective:** To compare the total number of major surgeries and the average HbA1C levels between patients with and without a history of transplants.

2)How does the total number of major surgeries and average HbA1C differ between patients with and without a history of transplants? 

**Steps Performed:**

**Step 1: Creating a PivotTable**
1. Selected the Healthcare dataset.
2. Navigated to **Insert → PivotTable**.
3. Selected **New Worksheet** and created the PivotTable.
4. Renamed the worksheet to `Transplant Analysis`.

**Step 2: Configuring PivotTable Fields**
1. Dragged **Any Transplants** to the Rows area.
2. Dragged **NumberOfMajorSurgeries** to the Values area.
3. Changed the summarisation method to **Sum** to calculate the total number of major surgeries.
4. Dragged **HBA1C** to the Values area.
5. Changed the summarisation method to **Average** to calculate the average HbA1C level.
6. Verified that the PivotTable displayed separate results for patients with and without a history of transplants.

**Step 3: Creating a Chart**
1. Selected the PivotTable data.
2. Navigated to **Insert → Chart**.
3. Selected a suitable chart type to visualise the comparison.
4. Added an appropriate chart title and data labels.
5. Verified that the chart represented both transplant groups and the two calculated measures.

4.3 Analysis using Column/Bar Chart:
1) How do healthcare charges vary based on different weight statuses and diabetes statuses? 
### Analysis: Healthcare Charges by Weight Status and Diabetes Status

**Objective:** To analyse how healthcare charges vary among patients with different weight statuses and diabetes statuses.

**Steps Performed:**

**Step 1: Creating a PivotTable**
1. Selected the Healthcare dataset.
2. Navigated to **Insert → PivotTable**.
3. Selected **New Worksheet** and created the PivotTable.
4. Renamed the worksheet to `Charges Analysis`.

**Step 2: Configuring PivotTable Fields**
1. Dragged **Weight Status** to the Rows area.
2. Dragged **Diabetes Status** to the Columns area.
3. Dragged **Charges** to the Values area.
4. Changed the summarisation method to **Sum** to calculate the total healthcare charges.
5. Verified that the PivotTable displayed total charges for each combination of weight status and diabetes status.

**Step 3: Creating a Bar Chart**
1. Selected the PivotTable data.
2. Navigated to **Insert → Chart**.
3. Selected **Bar Chart** to visualise the comparison.
4. Added an appropriate chart title: `Healthcare Charges by Weight Status and Diabetes Status`.
5. Added data labels to display the charge values.
6. Verified that the chart clearly represented the differences in healthcare charges across the categories.
Observations:

Patients in the Obesity category also include individuals with Normal diabetes status.

The number of patients with Normal diabetes status is higher than the number of patients with Diabetes status in the Obesity category.

The number of patients in the Underweight category is comparatively lower than in the Normal Weight, Overweight, and Obesity categories.

Healthcare charges vary across different combinations of weight status and diabetes status.

 2) Can you compare the average charges for each hospital tier within different states?
### Analysis: Average Healthcare Charges by Hospital Tier and State

**Objective:** To compare the average healthcare charges across different hospital tiers within different states.

**Steps Performed:**

**Step 1: Creating a PivotTable**
1. Selected the Healthcare dataset.
2. Navigated to **Insert → PivotTable**.
3. Selected **New Worksheet** and created the PivotTable.
4. Renamed the worksheet to `Hospital Tier Analysis`.

**Step 2: Configuring PivotTable Fields**
1. Dragged **State ID** to the Rows area.
2. Dragged **Hospital Tier** to the Columns area.
3. Dragged **Charges** to the Values area.
4. Changed the summarisation method from **Sum** to **Average** to calculate the average healthcare charges.
5. Verified that the PivotTable displayed the average charges for each hospital tier within each state.

**Step 3: Creating a Bar Chart**
1. Selected the PivotTable data.
2. Navigated to **Insert → Chart**.
3. Selected **Bar Chart** to visualise the comparison.
4. Added an appropriate chart title: `Average Healthcare Charges by Hospital Tier and State`.
5. Added data labels to display the average charge values.
6. Verified that the chart clearly represented the differences in average healthcare charges across hospital tiers and states.

 3.Is there any correlation between age and both BMI and HbA1C in the dataset? 
 ### Analysis: Correlation Between Age, BMI and HbA1C

**Objective:** To identify whether there is any relationship between patient age and their BMI and HbA1C levels.

**Steps Performed:**

**Step 1: Selecting the Data**
1. Opened the `Healthcare` worksheet.
2. Selected the **Age**, **BMI**, and **HBA1C** columns from the dataset.

**Step 2: Creating a Scatter Chart**
1. Navigated to **Insert → Chart**.
2. Selected **Scatter Chart** to visualise the relationship between the variables.
3. Configured the first data series:
   - **Series Name:** BMI
   - **X-axis Values:** Age column
   - **Y-axis Values:** BMI column
4. Added a second data series to the same chart:
   - **Series Name:** HBA1C
   - **X-axis Values:** Age column
   - **Y-axis Values:** HBA1C column
5. Used the Age column as the X-axis for both data series.

**Step 3: Formatting the Chart**
1. Added the chart title: `Relationship Between Age, BMI and HbA1C`.
2. Added axis titles:
   - **X-axis:** Age
   - **Y-axis:** BMI and HbA1C
3. Displayed the legend to distinguish the BMI and HbA1C data series.
4. Verified that both series were plotted against the same Age values.

4.Explore the relationship between age and healthcare charges. 

### Analysis: Relationship Between Age and Healthcare Charges

**Objective:** To explore the relationship between patient age and average healthcare charges.

**Steps Performed:**

**Step 1: Creating a PivotTable**
1. Selected the Healthcare dataset.
2. Navigated to **Insert → PivotTable**.
3. Selected **New Worksheet** and created the PivotTable.
4. Renamed the worksheet to `Age and Charges Analysis`.

**Step 2: Configuring PivotTable Fields**
1. Dragged **Age** to the Rows area.
2. Dragged **Charges** to the Values area.
3. Opened the Value Field Settings for Charges.
4. Selected **Average** from the summarisation options to calculate the average healthcare charges for each age group.
5. Verified that the PivotTable displayed the average charges for each age.

**Step 3: Creating a Scatter Chart**
1. Selected the **Age** and **Charges** columns from the original Healthcare dataset.
2. Navigated to **Insert → Chart**.
3. Selected **Scatter Chart**.
4. Configured the data series as follows:
   - **Series Name:** Healthcare Charges
   - **X-axis Values:** Age column
   - **Y-axis Values:** Charges column

**Step 4: Formatting the Chart**
1. Added the chart title: `Relationship Between Age and Healthcare Charges`.
2. Added the X-axis title as **Age**.
3. Added the Y-axis title as **Charges**.
4. Verified that the scatter chart displayed the relationship between patient age and healthcare charges.

5 Dashboard Creation:  
1)Build an interactive dashboard that consolidates all key insights using the above 
visualizations. Ensure visual clarity and ease of interpretation for all chart types. 

2)Add slicers for the fields “Weight Status” and “Diabetes Status” to enable filtering across 
all visualizations, supporting comparison of health outcomes and charges based on body 
weight and diabetes condition.
### Interactive Healthcare Dashboard

**Objective:** To create an interactive healthcare dashboard that consolidates key insights from the dataset using KPIs, PivotTables, charts, and slicers. The dashboard allows users to explore healthcare information dynamically based on Diabetes Status and Weight Status.

**Steps Performed:**

**Step 1: Creating a KPI PivotTable**
1. Selected the Healthcare dataset.
2. Navigated to **Insert → PivotTable**.
3. Selected **New Worksheet** and created the PivotTable.
4. Renamed the worksheet to `Dashboard`.
5. Configured the PivotTable fields to display the following key performance indicators (KPIs):
   - **Diabetes Status:** Used to identify the diabetes categories.
   - **Total Patients:** Added Customer ID to the Values area and selected **Count**.
   - **Diabetes Prevalence:** Calculated the percentage of patients with diabetes among the total patients.
   - **Total Hospital Charges:** Added Charges to the Values area and selected **Sum**.
   - **Average Hospital Charges:** Added Charges to the Values area and selected **Average**.
6. Formatted the KPI values for clear presentation.

**Step 2: Adding Slicers**
1. Selected a PivotTable in the dashboard.
2. Navigated to **Insert → Slicer**.
3. Selected the following fields:
   - Diabetes Status
   - Weight Status
4. Inserted the slicers into the dashboard.
5. Positioned the slicers appropriately for easy access.

**Step 3: Connecting Slicers to All PivotTables**
1. Selected the Diabetes Status slicer.
2. Opened the **Report Connections** or **PivotTable Connections** option.
3. Selected all relevant PivotTables created for the healthcare analysis.
4. Repeated the same process for the Weight Status slicer.
5. Verified that both slicers were connected to all relevant PivotTables.

**Step 4: Consolidating Dashboard Visualizations**
1. Arranged the KPI summary and previously created charts on the Dashboard worksheet.
2. Included the visualizations covering:
   - Cancer History by Smoker Status.
   - Major Surgeries and Average HbA1C by Transplant History.
   - Healthcare Charges by Weight Status and Diabetes Status.
   - Average Healthcare Charges by Hospital Tier and State.
   - Relationship Between Age, BMI and HbA1C.
   - Relationship Between Age and Healthcare Charges.
3. Arranged the charts and KPI summary to maintain a clear and organised dashboard layout.
4. Added appropriate chart titles, axis labels, legends, and data labels wherever required.

**Step 5: Testing Dashboard Interactivity**
1. Selected different categories using the Diabetes Status slicer.
2. Observed the changes in the KPI values and connected PivotTable visualizations.
3. Selected different categories using the Weight Status slicer.
4. Verified that the KPI values and connected charts updated according to the selected filters.
5. Tested different combinations of both slicers to ensure that the dashboard responded correctly.

### Dashboard Layout

The interactive healthcare dashboard is organised into the following sections to ensure visual clarity and ease of interpretation.

**1. Top Section – Dashboard Title and Slicers**
- Dashboard title.
- Diabetes Status slicer.
- Weight Status slicer.

**2. KPI Section – Key Performance Indicators**
- Total Patients.
- Diabetes Prevalence.
- Total Hospital Charges.
- Average Hospital Charges.

**3. Middle Section – Healthcare Analysis Charts**
- Healthcare Charges by Weight Status and Diabetes Status.
- Weight Status Analysis.
- Diabetes Status Analysis.
- Average Healthcare Charges by Hospital Tier and State.

**4. Bottom Section – Age-Related and Other Analysis Charts**
- Relationship Between Age, BMI and HbA1C.
- Relationship Between Age and Healthcare Charges.
- Other healthcare analysis visualizations.

**Dashboard Interactivity:**

The Diabetes Status and Weight Status slicers are connected to all relevant PivotTables. Selecting different slicer values dynamically updates the KPI values and associated charts, allowing users to explore the healthcare data using different filter combinations.


