# Healthcare Data Analysis

## 1. Project Overview

### Problem Statement

This project focuses on analysing healthcare data to identify patient health trends, medical histories, hospitalisation patterns, and healthcare costs.

The project involves data cleaning, transformation, data integration, exploratory data analysis (EDA), and data visualisation to generate meaningful insights that support informed decision-making and improve healthcare management.

### Project Objectives

- Identify patient health conditions and medical histories.
- Analyse the relationship between smoking habits and cancer history.
- Examine BMI and HbA1c levels across different patient groups.
- Explore hospitalisation patterns across hospital tiers and states.
- Analyse healthcare expenditure across different patient categories.
- Develop an interactive dashboard to present key healthcare insights.

## 2. Dataset Description

The healthcare dataset consists of three related tables containing 2,335 customer records, covering customer information, medical examinations, and hospitalisation details.

The tables are linked through the common **Customer ID** column, enabling comprehensive analysis of patient health conditions, medical history, hospitalisation patterns, and healthcare expenses.

### 2.1 Dataset Structure

| Worksheet | Rows | Columns | Description |
|---|---:|---:|---|
| Customer Names | 2,335 | 2 | Customer identification and names |
| Medical Examinations | 2,335 | 8 | Health conditions, examination results, and medical history |
| Hospitalisation Details | 2,343 | 9 | Hospitalisation records, hospital tiers, and healthcare charges |
| **Total** | **7,013** | **19** | **Combined dataset records and columns** |

### 2.2 Key Data Attributes

- **Customer Information:** Customer IDs and names.
- **Medical Information:** BMI, HbA1c, heart issues, transplant history, cancer history, major surgeries, and smoking status.
- **Hospitalisation Information:** Admission dates, number of children, healthcare charges, hospital tiers, city tiers, and state IDs.

### 2.3 Data Overview

- Customer Names and Medical Examinations contain 2,335 unique Customer IDs each.
- Hospitalisation Details contains 2,343 records, indicating that some customers have multiple hospitalisation records.
- Hospital Tier 2 has the highest number of records, with 1,339.
- Healthcare charges range from ₹563.84 to ₹63,770.43, based on the reported dataset values.

## 3. Data Preprocessing and Cleaning

### 3.1 Identifying Missing Values

**Objective:** Identify the number of missing values represented by `?` in each column of the Medical Examinations and Hospitalisation Details tables.

**Methodology:**

Used the Excel `COUNTIF` function to identify and count occurrences of `?`.

**Formula:**

```excel
=COUNTIF(A2:A2344,"~?")
```

The tilde (`~`) acts as an escape character, allowing Excel to treat `?` as a literal character rather than a wildcard.

**Results:**

- **Hospitalisation Details:** Missing-value counts across the nine columns were 6, 2, 3, 0, 0, 0, 1, 1, and 2.
- **Medical Examinations:** Missing-value counts across the eight columns were 0, 0, 0, 0, 0, 0, 0, and 2.

**Outcome:** Identified missing values across both worksheets to determine the appropriate cleaning strategies.

### 3.2 Handling Missing Month and Year Values

**Objective:** Replace missing values in the Month and Year columns.

**Month column:**

1. Inserted a temporary helper column next to the Month column.
2. Used an IF formula to replace missing values with `Sep` while retaining existing values.
3. Applied the formula to all records.
4. Copied the resulting values.
5. Used Paste Special → Values to update the original column.

**Year column:**

1. Calculated the average of the available Year values using `AVERAGE()`.
2. Used `ROUND()` to round the calculated average to the nearest integer.
3. Replaced missing Year values with the rounded average.
4. Verified the updated values.

### 3.3 Handling Missing Values in Smoker, Hospital Tier, and City Tier

**Objective:** Identify the most frequently occurring value (Mode) in each column and use it to replace missing values.

**Methodology:**

Used `COUNTIF` to calculate the frequency of each category.

**Smoker:**

```excel
=COUNTIF(H2:H2336,"yes")
=COUNTIF(H2:H2336,"no")
```

**Hospital Tier:**

```excel
=COUNTIF(W2:W2344,"tier - 1")
=COUNTIF(W2:W2344,"tier - 2")
=COUNTIF(W2:W2344,"tier - 3")
```

**City Tier:**

```excel
=COUNTIF(X2:X2344,"tier - 1")
=COUNTIF(X2:X2344,"tier - 2")
=COUNTIF(X2:X2344,"tier - 3")
```

**Steps:**

1. Calculated the frequency of each category using `COUNTIF`.
2. Compared the counts to identify the most frequently occurring value.
3. Used the IF function to replace missing values with the corresponding Mode.
4. Retained existing values without modification.
5. Applied the formulas to all records.
6. Used Copy → Paste Special → Values to update the original columns.

**Outcome:** Replaced missing categorical values using the most frequently occurring category in each column.

### 3.4 Handling Missing State ID Values

**Objective:** Handle missing State ID values without assigning an incorrect geographical identifier.

**Steps:**

1. Used `COUNTIF` to examine the frequency of existing State ID values.
2. Checked the State ID column for missing entries.
3. Replaced missing values with `Unknown` because the actual State ID could not be established reliably.
4. Used the following helper-column formula:

```excel
=IF(Y2="","Unknown",Y2)
```

5. Applied the formula to all records.
6. Copied the results and used Paste Special → Values.

**Outcome:** Missing State ID values were identified and represented as `Unknown`.

## 4. Data Transformation

### 4.1 Splitting Customer Names

**Objective:** Split the Names column in the Customer Names worksheet into three meaningful columns: Title, First Name, and Last Name.

**Original Name Format:**

`Last Name, Title. First Name`

Example: `Hawks, Mr. John`

**Steps:**

1. Created three new columns: Title, First Name, and Last Name.
2. Extracted the title using:

```excel
=LEFT(TRIM(MID(B2,FIND(",",B2)+1,LEN(B2))),FIND(".",TRIM(MID(B2,FIND(",",B2)+1,LEN(B2))))-1)
```

3. Extracted the first name using:

```excel
=TRIM(LEFT(TRIM(RIGHT(B2,LEN(B2)-FIND(".",B2))),FIND(" ",TRIM(RIGHT(B2,LEN(B2)-FIND(".",B2)))&" ")-1))
```

4. Extracted the last name using:

```excel
=LEFT(B2,FIND(",",B2)-1)
```

5. Applied the formulas to all records using the fill handle.
6. Verified the extracted values.

**Expected Output:**

| Original Name | Title | First Name | Last Name |
|---|---|---|---|
| Hawks, Mr. John | Mr | John | Hawks |

### 4.2 Converting NumberOfMajorSurgeries into Numerical Data

**Objective:** Convert non-numeric values in the NumberOfMajorSurgeries column into meaningful numerical values.

**Steps:**

1. Created a new column named `NumberOfMajorSurgeries_Cleaned`.
2. Used the IF function to convert `No major surgery` into `0`, while retaining existing numerical values.

**Formula:**

```excel
=IF(G2="No major surgery",0,VALUE(G2))
```

3. Applied the formula to all records.
4. Verified that:
   - `No major surgery` was converted to `0`.
   - Existing numerical values such as 1.00, 2.00, and 3.00 were retained as numbers.
5. Copied the cleaned column and used Paste Special → Values.

**Outcome:** Converted the NumberOfMajorSurgeries column into numerical data for further analysis.

### 4.3 Standardising Heart Issues and Smoker Columns

**Objective:** Identify and correct inconsistencies in the Heart Issues and Smoker columns.

**Steps:**

1. Reviewed both columns to identify inconsistent text formatting.
2. Used the `PROPER` function to standardise capitalization.

**Heart Issues:**

```excel
=PROPER(A2)
```

**Smoker:**

```excel
=PROPER(B2)
```

3. Applied the formulas to all records.
4. Reviewed the results to identify capitalization differences such as `YES`, `yes`, and `Yes`.
5. Replaced the original values using Copy → Paste Values.
6. Verified the updated columns.

**Outcome:** Standardised the capitalization of categorical values in both columns.

### 4.4 Creating Weight Status

**Objective:** Create a new column named Weight Status to categorise patients based on their BMI values.

**BMI Classification:**

| BMI Range | Weight Status |
|---|---|
| Below 18.5 | Underweight |
| 18.5 – Below 25 | Normal Weight |
| 25 – Below 30 | Overweight |
| 30 and Above | Obesity |

**Formula:**

```excel
=IF(M3<18.5,"Underweight",IF(M3<25,"Normal Weight",IF(M3<30,"Overweight","Obesity")))
```

**Steps:**

1. Created a new column named Weight Status.
2. Applied nested IF conditions to classify BMI values.
3. Filled the formula down for all records.
4. Verified the resulting classifications against the specified BMI ranges.

**Outcome:** Categorised patients into four weight-status groups based on BMI.

### 4.5 Creating Diabetes Status

**Objective:** Create a new column named Diabetes Status based on HbA1c values.

**HbA1c Classification:**

| HbA1c Range | Diabetes Status |
|---|---|
| Below 5.7 | Normal |
| 5.7 – Below 6.5 | Prediabetes |
| 6.5 and Above | Diabetes |

**Formula:**

```excel
=IF(O3="","",IF(O3<5.7,"Normal",IF(O3<6.5,"Prediabetes","Diabetes")))
```

**Steps:**

1. Created a new column named Diabetes Status.
2. Used nested IF conditions to categorise HbA1c values.
3. Retained blank values when the corresponding HbA1c value was missing.
4. Applied the formula to all records.
5. Verified the classifications against the specified ranges.

**Outcome:** Categorised patients into Normal, Prediabetes, and Diabetes groups based on HbA1c values.

### 4.6 Combining Year, Month, and Date into a Single Date Column

**Objective:** Combine the Year, Month, and Date columns in the Hospitalisation Details worksheet into a single date column.

**Steps:**

1. Created a new column for the combined date.
2. Used the DATE, DATEVALUE, and MONTH functions to combine the individual components.

Assuming:
- Z3 contains Year.
- AA3 contains Month.
- AB3 contains Date.

**Formula:**

```excel
=DATE(VALUE(Z3),MONTH(DATEVALUE("1 "&AA3&" 2000")),VALUE(AB3))
```

3. Applied the formula to all records.
4. Formatted the resulting column using the custom date format:

```text
DD-MMM-YYYY
```

5. Verified the combined dates.

**Outcome:** Combined the separate date components into a single date column.

*Note: The column should be named Date of Birth only if these source fields represent patients' birth dates. If they represent hospitalisation dates, use Admission Date instead.*

### 4.7 Calculating Customer Age

**Objective:** Calculate the age of each customer using the date of birth and the dataset collection date, 8 June 2023.

**Steps:**

1. Created a new column named Age.
2. Used the DATEDIF function to calculate completed years as of the collection date.

Assuming AH3 contains the customer's Date of Birth:

```excel
=DATEDIF(AH3,DATE(2023,6,8),"Y")
```

3. Applied the formula to all records.
4. Verified the calculated ages.

**Outcome:** Calculated customer age using the specified dataset collection date.

### 4.8 Formatting Healthcare Charges

**Objective:** Format the Charges column as currency.

**Steps:**

1. Selected the Charges column.
2. Navigated to Home → Number Format.
3. Selected Currency.
4. Selected the appropriate currency symbol based on the original dataset.
5. Set the decimal places to two.
6. Verified the formatting.

**Outcome:** Formatted healthcare charges as currency with two decimal places.

## 5. Exploratory Data Analysis (EDA)

### 5.1 Combining the Three Tables Using VLOOKUP

**Objective:** Create a consolidated Healthcare dataset by retrieving information from the three worksheets using Customer ID as the common identifier.

**Steps:**

1. Created a new worksheet named `Healthcare`.
2. Retained Customer ID as the common lookup column.
3. Used VLOOKUP with an exact-match condition (`0`) to retrieve information from the source worksheets.

**Example formulas:**

**Customer Names:**

```excel
=VLOOKUP(A2,'Customer Names'!A:D,4,0)
```

**Medical Examinations:**

```excel
=VLOOKUP(A2,'Medical Examinations'!A:H,8,0)
```

**Hospitalisation Details:**

```excel
=VLOOKUP(A2,'Hospitalisation Details'!A:K,11,0)
```

The column index was adjusted according to the required field in each source worksheet.

4. Applied the formulas to the required records.
5. Verified the retrieved values against the original worksheets.
6. Retained the following fields in the consolidated dataset:

- Customer ID
- First Name
- BMI
- HbA1c
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

**Outcome:** Created a consolidated Healthcare dataset for further analysis.

**Data integration consideration:** Since Hospitalisation Details contains 2,343 records for 2,335 unique customers, the consolidated dataset should preserve all hospitalisation records when analysing charges and hospitalisation patterns. A single VLOOKUP against a one-to-many relationship retrieves only the first matching record. Therefore, use the Hospitalisation Details table as the base when retaining all hospitalisation records and retrieve customer and medical information using Customer ID.

## 6. Data Analysis and Visualisation

### 6.1 Analysis Using Pie/Donut Charts

#### 6.1.1 Distribution of Cancer History Among Smokers and Non-Smokers

**Objective:** Analyse the distribution of cancer history among smokers and non-smokers.

**PivotTable Configuration:**

- Rows: Smoker
- Columns: Cancer History
- Values: Count of Customer ID

**Steps:**

1. Selected the Healthcare dataset.
2. Navigated to Insert → PivotTable.
3. Created a PivotTable using the Healthcare dataset.
4. Named the PivotTable `cancer_history`.
5. Configured the PivotTable fields as mentioned above.
6. Changed the Customer ID aggregation to Count.
7. Verified the results.

**Chart Creation:**

1. Selected the PivotTable results.
2. Navigated to Insert → Chart.
3. Created a chart to compare cancer history among smokers and non-smokers.
4. Added a descriptive title, legend, and data labels.

**Analysis:**

The PivotTable shows the following distribution:

| Smoking Status | No Cancer History | Cancer History | Total |
|---|---:|---:|---:|
| Non-Smokers | 1,540 | 307 | 1,847 |
| Smokers | 404 | 84 | 488 |
| **Total** | **1,944** | **391** | **2,335** |

- Non-smokers account for 1,847 patients, of whom 307 have a cancer history.
- Smokers account for 488 patients, of whom 84 have a cancer history.
- Approximately 16.62% of non-smokers have a cancer history.
- Approximately 17.21% of smokers have a cancer history.

**Conclusion:**

The proportion of patients with a cancer history is slightly higher among smokers than non-smokers in this dataset. However, the difference is relatively small, and the findings do not establish a causal relationship between smoking and cancer history.

#### 6.1.2 Major Surgeries and Average HbA1c by Transplant History

**Objective:** Compare the total number of major surgeries and average HbA1c levels between patients with and without a history of transplants.

**PivotTable Configuration:**

- Rows: Any Transplants
- Values: Sum of NumberOfMajorSurgeries
- Values: Average of HbA1c

**Steps:**

1. Selected the Healthcare dataset.
2. Navigated to Insert → PivotTable.
3. Created a PivotTable named `major_surgeries`.
4. Added Any Transplants to Rows.
5. Added NumberOfMajorSurgeries to Values and selected Sum.
6. Added HbA1c to Values and selected Average.
7. Verified the calculated results.

**Chart Creation:**

1. Selected the PivotTable results.
2. Created separate charts or a suitable combination chart to compare the two measures.
3. Added appropriate chart titles, legends, and data labels.

**Analysis:**

- Patients without a transplant history have an average HbA1c level of 6.67.
- Patients with a transplant history have an average HbA1c level of 5.19.
- A total of 1,417 major surgeries are recorded for patients without transplants.
- A total of 162 major surgeries are recorded for patients with transplants.
- Patients without transplants account for approximately 89.74% of the recorded major surgeries.
- Patients with transplants account for approximately 10.26%.

The overall average HbA1c level is 6.58, with 1,579 major surgeries recorded across both groups.

**Conclusion:**

Patients without a history of transplants account for the majority of recorded major surgeries and have a higher average HbA1c level than patients with a transplant history.

These findings describe differences between the two groups but do not establish a causal relationship between transplant history, HbA1c levels, and major surgeries.

### 6.2 Analysis Using Column/Bar Charts

#### 6.2.1 Healthcare Charges by Weight Status and Diabetes Status

**Objective:** Analyse how healthcare charges vary across different weight-status and diabetes-status categories.

**PivotTable Configuration:**

- Rows: Weight Status
- Columns: Diabetes Status
- Values: Average of Charges

An additional PivotTable can be created using Sum of Charges to compare total expenditure.

**Steps:**

1. Selected the Healthcare dataset.
2. Navigated to Insert → PivotTable.
3. Created a PivotTable in a new worksheet.
4. Renamed the worksheet `weight_statuses`.
5. Added Weight Status to Rows.
6. Added Diabetes Status to Columns.
7. Added Charges to Values.
8. Changed the aggregation to Average to calculate average healthcare charges.
9. Created a separate PivotTable using Sum of Charges if total expenditure comparison was required.
10. Created a clustered column or bar chart to visualise the results.

**Patient Distribution Findings:**

The patient distribution across weight categories is:

| Weight Status | Number of Patients |
|---|---:|
| Obesity | 1,221 |
| Overweight | 498 |
| Normal Weight | 464 |
| Underweight | 152 |
| **Total** | **2,335** |

- Obesity is the largest weight category, representing approximately 52.29% of the patients.
- Overweight represents 498 patients.
- Normal Weight represents 464 patients.
- Underweight is the smallest category, with 152 patients.

Across diabetes categories:

- Normal: 1,095 patients.
- Diabetes: 784 patients.
- Prediabetes: 456 patients.

Among obese patients, Normal diabetes status has the highest representation, with 561 patients, followed by Diabetes (423) and Prediabetes (237).

**Conclusion:**

Obesity is the most represented weight category in the dataset, while Underweight is the least represented.

Normal diabetes status has the highest patient count across the dataset.

The patient distribution findings describe the composition of the dataset. The average and total healthcare charges must be examined separately to determine how expenditure varies across weight and diabetes categories.

#### 6.2.2 Average Healthcare Charges by Hospital Tier and State

**Objective:** Compare average healthcare charges across different hospital tiers within different states.

**PivotTable Configuration:**

- Rows: State ID
- Columns: Hospital Tier
- Values: Average of Charges

**Steps:**

1. Selected the Healthcare dataset.
2. Navigated to Insert → PivotTable.
3. Created a new PivotTable.
4. Renamed the worksheet `Hospital Tier Analysis`.
5. Added State ID to Rows.
6. Added Hospital Tier to Columns.
7. Added Charges to Values.
8. Changed the summarisation method to Average.
9. Verified the average charges across hospital tiers and states.
10. Created a bar chart titled `Average Healthcare Charges by Hospital Tier and State`.

**Analysis:**

- **Overall Trend:** Tier 1 hospitals have the highest average healthcare charges (₹31,239.99), followed by Tier 2 (₹15,760.51) and Tier 3 (₹10,220.83).

- **State-Level Comparison:** State R1016 records the highest average charge for Tier 1 hospitals (₹39,868.62).

- **Tier 2 Comparison:** State R1011 records the highest average charge for Tier 2 hospitals (₹18,997.28).

- **Notable Exception:** State R1014 shows a different pattern, where Tier 3 hospitals (₹16,950.49) have higher average charges than Tier 1 (₹13,891.86) and Tier 2 (₹2,395.17).

- **Geographical Variation:** Average healthcare charges vary across the ten states, indicating differences in expenditure across geographical locations and hospital tiers.

**Conclusion:**

The analysis shows that Tier 1 hospitals generally have higher average healthcare charges than Tier 2 and Tier 3 hospitals.

However, variations such as those observed in State R1014 indicate that hospital tier alone does not explain all differences in healthcare expenditure.

#### 6.2.3 Relationship Between Age, BMI, and HbA1c

**Objective:** Explore the relationship between patient age and two health indicators: BMI and HbA1c.

**Steps:**

1. Opened the Healthcare worksheet.
2. Selected Age, BMI, and HbA1c data.
3. Created separate scatter charts for:
   - Age vs BMI.
   - Age vs HbA1c.
4. Configured Age as the X-axis and the respective health indicator as the Y-axis.
5. Added appropriate chart titles, axis titles, and data labels where required.
6. Verified that both charts referenced the correct data columns.

**Analysis:**

- **Age vs BMI:** Average BMI remains relatively stable across different age groups, with most observed values ranging between 31 and 34. This indicates no clear increasing or decreasing trend in average BMI with age.

- **Age vs HbA1c:** The HbA1c values show fluctuations across different age groups. The trend should be interpreted using the corrected HbA1c data and verified chart references.

- **Comparison:** BMI appears relatively consistent across age groups, whereas HbA1c shows greater variation.

**Conclusion:**

The analysis suggests that average BMI remains relatively stable across age groups. A reliable conclusion about the relationship between age and HbA1c requires verification of the plotted values and, if necessary, additional correlation analysis.

#### 6.2.4 Relationship Between Age and Healthcare Charges

**Objective:** Explore the relationship between patient age and average healthcare charges.

**PivotTable Configuration:**

- Rows: Age
- Values: Average of Charges

**Steps:**

1. Selected the Healthcare dataset.
2. Navigated to Insert → PivotTable.
3. Created a PivotTable in a new worksheet.
4. Renamed the worksheet `Age and Charges Analysis`.
5. Added Age to Rows.
6. Added Charges to Values.
7. Changed the summarisation method to Average.
8. Verified the average healthcare charges for each age group.

**Chart Creation:**

1. Selected Age and Charges from the original Healthcare dataset.
2. Navigated to Insert → Chart.
3. Selected Scatter Chart.
4. Configured Age as the X-axis and Charges as the Y-axis.
5. Added a descriptive chart title and axis labels.
6. Verified the plotted data.

**Analysis:**

- **Overall Healthcare Expenditure:** The overall average healthcare charge across the dataset is ₹17,700.73.

- **Variation Across Age Groups:** Average healthcare charges fluctuate across different age groups, with most observed values ranging between ₹15,000 and ₹23,000.

- **Higher Average Charges:** Age 65 records the highest average healthcare charge among the selected age groups at ₹23,275.53, followed by age 62 at ₹22,024.46.

- **Lower Average Charges:** Age 58 records a comparatively lower average healthcare charge of ₹15,539.13.

- **Age-Related Spending Pattern:** Although some older age groups show relatively higher average healthcare charges, the overall pattern does not demonstrate a consistent increase in expenditure with age.

**Conclusion:**

Average healthcare charges vary across age groups, but the observed pattern does not establish a consistent relationship between increasing age and healthcare expenditure.

Other factors, including medical conditions, hospitalisation frequency, treatment requirements, and hospital tier, may also contribute to differences in healthcare costs.

## 7. Interactive Dashboard Development

### 7.1 Dashboard Objective

Develop an interactive healthcare dashboard that consolidates key analytical findings using KPIs, PivotTables, charts, and slicers.

The dashboard enables users to explore patient health information, medical conditions, and healthcare expenditure by applying Diabetes Status and Weight Status filters.

### 7.2 Creating KPI Metrics

Created a new worksheet named `Dashboard` to display key performance indicators.

The following KPIs were included:

| KPI | Description |
|---|---|
| Total Patients | Number of unique patients |
| Diabetes Prevalence | Percentage of patients classified as Diabetes |
| Total Hospital Charges | Sum of healthcare charges |
| Average Hospital Charges | Average healthcare charges |

The KPI calculations were configured using PivotTables and appropriate aggregation methods.

### 7.3 Adding Interactive Slicers

Added slicers for the following fields:

- Diabetes Status
- Weight Status

**Steps:**

1. Selected a PivotTable in the Dashboard worksheet.
2. Navigated to Insert → Slicer.
3. Selected Diabetes Status and Weight Status.
4. Inserted the slicers.
5. Positioned them at the top of the dashboard.
6. Used Report Connections or PivotTable Connections to connect the slicers to the relevant PivotTables.
7. Tested the slicers to verify that the connected PivotTables and charts responded to the selected filters.

### 7.4 Dashboard Visualisations

The dashboard consolidates the following analytical visualisations:

- Cancer History by Smoker Status.
- Major Surgeries and Average HbA1c by Transplant History.
- Healthcare Charges by Weight Status and Diabetes Status.
- Patient Distribution by Weight Status.
- Patient Distribution by Diabetes Status.
- Average Healthcare Charges by Hospital Tier and State.
- Relationship Between Age and BMI.
- Relationship Between Age and HbA1c.
- Relationship Between Age and Healthcare Charges.

### 7.5 Dashboard Layout

The dashboard is organised into four main sections.

**1. Top Section – Dashboard Title and Slicers**

- Healthcare Dashboard title.
- Diabetes Status slicer.
- Weight Status slicer.

**2. KPI Section – Key Performance Indicators**

- Total Patients.
- Diabetes Prevalence.
- Total Hospital Charges.
- Average Hospital Charges.

**3. Middle Section – Healthcare Analysis**

- Healthcare Charges by Weight Status and Diabetes Status.
- Weight Status Distribution.
- Diabetes Status Distribution.
- Average Healthcare Charges by Hospital Tier and State.

**4. Bottom Section – Age-Related and Other Analysis**

- Age vs BMI.
- Age vs HbA1c.
- Age vs Healthcare Charges.
- Other healthcare analysis visualisations.

### 7.6 Testing Dashboard Interactivity

1. Selected different categories using the Diabetes Status slicer.
2. Observed changes in the connected KPI values and charts.
3. Selected different categories using the Weight Status slicer.
4. Verified that the relevant visualisations responded to the selected filters.
5. Tested different combinations of both slicers.
6. Verified that the dashboard displayed the corresponding filtered results.

**Outcome:**

Developed an interactive healthcare dashboard that consolidates the key findings from the analysis and enables users to explore healthcare trends through Diabetes Status and Weight Status filters.

## 8. Key Findings and Insights

Based on the exploratory analysis:

1. **Cancer History:** The proportion of patients with a cancer history is slightly higher among smokers (17.21%) than non-smokers (16.62%).

2. **Transplant History:** Patients without a transplant history account for the majority of recorded major surgeries and have a higher average HbA1c level.

3. **Weight Distribution:** Obesity is the largest weight category, representing approximately 52.29% of the patients.

4. **Diabetes Distribution:** Normal diabetes status accounts for the highest number of patients, followed by Diabetes and Prediabetes.

5. **Hospital Tier Analysis:** Tier 1 hospitals generally have higher average healthcare charges than Tier 2 and Tier 3 hospitals.

6. **Age and BMI:** Average BMI remains relatively stable across different age groups.

7. **Age and Healthcare Charges:** Healthcare expenditure varies across age groups, without a consistent increasing trend.

8. **Interactive Dashboard:** Diabetes Status and Weight Status slicers enable users to explore patient categories and associated healthcare metrics.

## 9. Tools and Technologies

- **Microsoft Excel for the Web**
- Excel Formulas and Functions:
  - IF
  - COUNTIF
  - AVERAGE
  - ROUND
  - PROPER
  - LEFT
  - RIGHT
  - MID
  - FIND
  - TRIM
  - VALUE
  - DATE
  - DATEVALUE
  - MONTH
  - DATEDIF
  - VLOOKUP
- PivotTables
- PivotCharts
- Bar and Column Charts
- Pie and Donut Charts
- Scatter Charts
- Interactive Slicers
- Data Cleaning and Transformation
- Exploratory Data Analysis (EDA)
- Interactive Dashboard Development

## 10. Project Limitations and Considerations

- The dataset contains multiple hospitalisation records for some customers. Customer-level and hospitalisation-level analyses should be distinguished.
- Missing values were handled using the specified replacement strategies, which may influence some analytical results.
- The analysis is based on the available dataset and does not establish causal relationships between patient characteristics and healthcare outcomes.
- Average healthcare charges can vary due to multiple factors, including medical conditions, hospital tier, and hospitalisation frequency.
- Age-related trends should be interpreted using the dataset collection date rather than the current date.

## 11. Conclusion

This project demonstrates an end-to-end healthcare data analysis workflow using Microsoft Excel.

Starting with data cleaning and transformation, the project integrates multiple datasets using Customer ID and applies PivotTables, formulas, and visualisations to explore patient health conditions, medical histories, hospitalisation patterns, and healthcare expenses.

The exploratory analysis identifies differences in cancer history, transplant status, weight categories, diabetes status, hospital tiers, and healthcare expenditure across patient groups.

Finally, an interactive dashboard was developed using KPI metrics, charts, and slicers to provide a consolidated view of the key healthcare insights.

The project strengthened practical skills in data preprocessing, Excel functions, data integration, exploratory data analysis, data visualisation, and interactive dashboard development.
