 Healthcare-Data-Analysis
# 1.Problem Statement
This project focuses on analyzing healthcare data to identify patient health trends, medical histories, hospitalization patterns, and healthcare costs. It involves data cleaning, transformation, exploration, and visualization to generate meaningful insights that support informed decision-making and improve healthcare management.
## 2.Dataset Description
The healthcare dataset consists of three tables: Customer Names, Medical Examinations, and Hospitalisation Details, containing information on 2,335 customers. It includes customer profiles, medical conditions, health examination results, hospitalisation records, and healthcare expenses. These tables are linked through Customer ID, enabling comprehensive analysis of patient health profiles, medical histories, hospitalisation trends, and healthcare costs.
It contains three worksheets (tables), each representing a different aspect of healthcare information.
The three tables can be connected using the common Customer ID column to perform combined analysis.
###1. Customer Names
2,335 rows · 2 columns
Purpose: This table contains the basic customer information. It helps identify individual customers and link their details with medical examination and hospitalisation records.
Column	Description	Data values
Customer ID	Unique identifier assigned to each customer	Id1, Id2, Id3, etc.
name	Customer's name	Full names with titles such as Mr. and Ms.
Brief summary of data:
•	Contains 2,335 customer records.
•	All 2,335 Customer IDs are unique.
•	All customer names are unique.
•	No missing values were identified.
Analysis opportunity: This table can be used to identify customers and connect their personal information with their medical and hospitalisation details.
###2. Medical Examinations
2,335 rows · 8 columns
Purpose: This table contains customers' medical examination results, health conditions, medical histories, and lifestyle information. It helps analyse patient health profiles and identify potential health risks.
Column	Description	Data values
Customer ID	Unique customer identifier	Id1, Id2, Id3, etc.
BMI	Body Mass Index, indicating body weight relative to height	15.01 – 55.05
HBA1C	HbA1c measurement used to assess average blood glucose levels	4.00 – 12.00
Heart Issues	Indicates whether the customer has heart-related health issues	Yes / No
Any Transplants	Indicates whether the customer has undergone an organ transplant	Yes / No
Cancer history	Indicates whether the customer has a history of cancer	Yes / No
NumberOfMajorSurgeries	Number of major surgeries undergone	No major surgery, 1, 2, 3
smoker	Indicates the customer's smoking status	Yes / No / ?
Brief summary of data values:
Medical attribute	Data distribution
Heart Issues – No	1,409
Heart Issues – Yes	926
Transplants – No	2,191
Transplants – Yes	144
Cancer history – No	1,944
Cancer history – Yes	391
No major surgery	1,074
1 major surgery	965
2 major surgeries	274
3 major surgeries	22
Non-smokers	1,845
Smokers	488

Analysis opportunity: This table can be used to analyse health conditions, smoking habits, surgery history, BMI distribution, and HbA1c levels among customers.
###3. Hospitalisation Details
2,343 rows · 9 columns
Purpose: This table contains customer hospitalisation records, including admission dates, number of children, medical expenses, hospital categories, city categories, and state information. It helps analyse healthcare costs and hospitalisation trends.
Column	Description	Data values
Customer ID	Customer identifier used to connect records	Id1, Id2, Id3, etc.
year	Year associated with the hospitalisation record	Year values
month	Month associated with the hospitalisation record	Jan, Feb, Mar, etc.
date	Day of the month	1–30
children	Number of children associated with the customer	0–5
charges	Healthcare expenses incurred by the customer	563.84 – 63,770.43
Hospital tier	Classification of the hospital based on its tier	Tier 1, Tier 2, Tier 3
City tier	Classification of the city based on its tier	Tier 1, Tier 2, Tier 3
State ID	Unique code representing the state	R1011, R1012, R1013, etc.
Brief summary of data values:
Attribute	Data distribution
Hospital Tier 1	309
Hospital Tier 2	1,339
Hospital Tier 3	694
City Tier 1	734
City Tier 2	812
City Tier 3	796
Customers with 0 children	
Numerical data summary:
Column	Average	Minimum	Maximum
Children	1.03	0	5
Charges	13,559.07	563.84	63,770.43
Date	15.55	1	30
Additional observations:
•	September has the highest number of records (369).
•	The dataset contains 17 distinct State ID values.
•	Hospital Tier 2 has the highest number of records (1,339).
•	The median healthcare charge is 9,634.54.
Analysis opportunity: This table can be used to analyse healthcare expenses, hospitalisation patterns, city and hospital tier distributions, and regional healthcare trends.
4. Overall Dataset Summary
Worksheet	Rows	Columns	Main purpose
Customer Names	2,335	2	Customer identification
Medical Examinations	2,335	8	Patient health information
Hospitalisation Details	2,343	9	Hospitalisation and expenses
Total	7,013	19	Healthcare analysis



