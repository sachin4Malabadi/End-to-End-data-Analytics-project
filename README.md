# End-to-End-data-Analytics-project

#STEP 0 — Project Introduction

##“ Build a complete Netflix Data Analytics Project. I have uncleaned dummy data on customer viewing and subscriptions.  will clean this data with Python Pandas, analyze it with SQL, and build an interactive dashboard in Power BI.”
#Project Objective
 need to answer the management’s questions:
1.	How many customers does Netflix have?
2.	Which subscription plan is the most popular?
3.	Which countries have the most customers?
4.	How many hours of content are customers watching?
5.	Which genre is the most popular?
6.	Which device is used the most?
7.	Which plan does the revenue come from?
8.	How many customers are Active vs Cancelled?
9.	How is customer engagement?
10.	Who are the high-value customers?
________________________________________
# STEP 1 — Understanding the Raw Excel Data
First of all:
##Netflix_Unclean_Project.xlsx
Open it.
“Before starting data cleaning,  need to understand the raw data.”
Explain the columns
Column	Meaning
Customer_ID	Unique customer
Customer_Name	Customer name
Age	Customer age
Title	Watched content
Content_Type	Movie / TV Show
Genre	Content genre
Country	Customer country
Subscription_Plan	Basic / Standard / Premium
Device	Watching device
Payment_Method	Payment method
Account_Status	Active / Cancelled / Paused
Join_Date	Customer joining date
Last_Watch_Date	Last activity
Watch_Hours	Total watch hours
Monthly_Fee	Monthly subscription fee
Movies_Watched	Movies watched
Episodes_Watched	Episodes watched
Avg_Rating	Average rating
Monthly_Sessions	Monthly sessions
Download_Count	Downloads
________________________________________# STEP 2 — Python Environment
Open Jupyter Notebook / VS Code / Google Colab.
##Import Libraries
import pandas as pd
import numpy as np
Explain:
•	pandas → Data manipulation 
•	numpy → Numerical operations 
________________________________________# STEP 3 — Load Excel
df = pd.read_excel("Netflix_Unclean_Project.xlsx")
Check:
df.head()
Then:
df.tail()
________________________________________
# STEP 4 — Initial Data Exploration
Shape
df.shape
Explain :
How many rows and columns are there?
________________________________________
Column Names
df.columns
________________________________________
Data Types
df.info()
Identify here:
•	Numeric columns 
•	Text columns 
•	Date columns 
•	Potentially wrong data types 
________________________________________
Statistical Summary
df.describe()
Explain:
•	Mean 
•	Minimum 
•	Maximum 
•	Standard deviation 
•	Quartiles 
________________________________________
# STEP 5 — Missing Values
df.isnull().sum()
You can also calculate the percentage:
(df.isnull().sum()/len(df)*100).sort_values(ascending=False)
Explain:
Missing values can affect business analysis, so cleaning is necessary.
________________________________________
# STEP 6 — Duplicate Records
df.duplicated().sum()
View the duplicate records:
df[df.duplicated()]
Remove:
df.drop_duplicates(inplace=True)
Then:
df.duplicated().sum()
Expected:
0
________________________________________
# STEP 7 — Dirty Values
The dataset may contain:
N/A
NULL
blank
spaces
can be present.
Replace:
df.replace(
    ["N/A", "NULL", "", " "],
    np.nan,
    inplace=True
)
Again:
df.isnull().sum()
________________________________________
# STEP 8 — Remove Extra Spaces
Remove leading/trailing spaces from text columns:
df = df.apply(
    lambda x: x.str.strip()
    if x.dtype == "object"
    else x
)
Explain:
"  India  "
becomes:
"India"
________________________________________
# STEP 9 — Standardize Text
For example:
df["Country"] = df["Country"].str.title()

df["Genre"] = df["Genre"].str.title()

df["Subscription_Plan"] = (
    df["Subscription_Plan"].str.title()
)

df["Account_Status"] = (
    df["Account_Status"].str.title()
)
Check:
df["Country"].unique()
________________________________________
# STEP 10 — Numeric Columns
Numeric columns:
numeric_cols = [
    "Age",
    "Watch_Hours",
    "Monthly_Fee",
    "Movies_Watched",
    "Episodes_Watched",
    "Avg_Rating",
    "Monthly_Sessions",
    "Download_Count"
]
Convert:
for col in numeric_cols:
    df[col] = pd.to_numeric(
        df[col],
        errors="coerce"
    )
Explain errors="coerce":
Any value that cannot be converted to a number will be turned into NaN by Pandas.
________________________________________
#STEP 11 — Invalid Age
Check:
df["Age"].describe()
If there are values such as:
-5
0
150
200
these are invalid values.
Filter:
df = df[
    (df["Age"] >= 18) &
    (df["Age"] <= 100)
]
________________________________________
# STEP 12 — Negative Values
Watch hours:
df[df["Watch_Hours"] < 0]
Remove:
df = df[df["Watch_Hours"] >= 0]
Monthly fee:
df = df[df["Monthly_Fee"] >= 0]
________________________________________
# STEP 13 — Date Cleaning
Convert:
df["Join_Date"] = pd.to_datetime(
    df["Join_Date"],
    errors="coerce"
)

df["Last_Watch_Date"] = pd.to_datetime(
    df["Last_Watch_Date"],
    errors="coerce"
)
Check:
df[["Join_Date","Last_Watch_Date"]].isnull().sum()
Invalid dates will automatically become NaT.
________________________________________
# STEP 14 — Handle Missing Values
Numeric:
df["Age"].fillna(
    df["Age"].median(),
    inplace=True
)

df["Watch_Hours"].fillna(
    df["Watch_Hours"].median(),
    inplace=True
)

df["Monthly_Fee"].fillna(
    df["Monthly_Fee"].median(),
    inplace=True
)
Categorical:
df["Payment_Method"].fillna(
    "Unknown",
    inplace=True
)

df["Device"].fillna(
    "Unknown",
    inplace=True
)

df["Genre"].fillna(
    "Unknown",
    inplace=True
)
________________________________________
# STEP 15 — Feature Engineering
Now the important Data Analyst part begins.
1. Customer Revenue
df["Customer_Revenue"] = (
    df["Monthly_Fee"] *
    df["Monthly_Sessions"]
)
________________________________________
2. Engagement Score
df["Engagement_Score"] = (
    df["Watch_Hours"] +
    df["Monthly_Sessions"] +
    df["Download_Count"]
)
________________________________________
3. Watch Category
df["Watch_Category"] = pd.cut(
    df["Watch_Hours"],
    bins=[0, 20, 50, 100, np.inf],
    labels=[
        "Low",
        "Medium",
        "High",
        "Very High"
    ]
)
________________________________________
4. Customer Age Group
df["Age_Group"] = pd.cut(
    df["Age"],
    bins=[17, 25, 35, 50, 100],
    labels=[
        "18-25",
        "26-35",
        "36-50",
        "51+"
    ]
)
________________________________________
5. Account Active Flag
df["Active_Flag"] = np.where(
    df["Account_Status"] == "Active",
    1,
    0
)
________________________________________
# STEP 16 — Final Data Quality Check
df.info()
df.isnull().sum()
df.duplicated().sum()
df.describe()
Important line:
“After cleaning,  will make sure the dataset is ready for analysis.”
________________________________________
# STEP 17 — Export Clean Excel
df.to_excel(
    "Netflix_Clean_Data.xlsx",
    index=False
)
CSV:
df.to_csv(
    "Netflix_Clean_Data.csv",
    index=False
)
________________________________________
# STEP 18 — SQL Database
Create the database in MySQL:
CREATE DATABASE netflix_db;
Use:
USE netflix_db;
Python to SQL connection:
from sqlalchemy import create_engine

engine = create_engine(
    "mysql+pymysql://root:password@localhost/netflix_db"
)
Upload:
df.to_sql(
    "netflix_customers",
    con=engine,
    if_exists="replace",
    index=False
)
________________________________________
# STEP 19 — SQL Analysis
Now solve the business questions.
Q1. Total Customers
SELECT COUNT(*) AS Total_Customers
FROM netflix_customers;
Q2. Active Customers
SELECT COUNT(*) AS Active_Customers
FROM netflix_customers
WHERE Account_Status = 'Active';
Q3. Customers by Subscription Plan
SELECT
    Subscription_Plan,
    COUNT(*) AS Customers
FROM netflix_customers
GROUP BY Subscription_Plan
ORDER BY Customers DESC;
Q4. Revenue by Subscription Plan
SELECT
    Subscription_Plan,
    SUM(Customer_Revenue) AS Revenue
FROM netflix_customers
GROUP BY Subscription_Plan
ORDER BY Revenue DESC;
Q5. Popular Genres
SELECT
    Genre,
    COUNT(*) AS Views
FROM netflix_customers
GROUP BY Genre
ORDER BY Views DESC;
Q6. Country-wise Customers
SELECT
    Country,
    COUNT(*) AS Customers
FROM netflix_customers
GROUP BY Country
ORDER BY Customers DESC;
Q7. Device Usage
SELECT
    Device,
    COUNT(*) AS Users
FROM netflix_customers
GROUP BY Device
ORDER BY Users DESC;
Q8. Average Watch Hours by Plan
SELECT
    Subscription_Plan,
    ROUND(AVG(Watch_Hours),2) AS Avg_Watch_Hours
FROM netflix_customers
GROUP BY Subscription_Plan;
Q9. Top 10 Customers by Revenue
SELECT
    Customer_Name,
    Customer_Revenue
FROM netflix_customers
ORDER BY Customer_Revenue DESC
LIMIT 10;
Q10. Movie vs TV Show
SELECT
    Content_Type,
    COUNT(*) AS Total
FROM netflix_customers
GROUP BY Content_Type;
Q11. Average Rating by Genre
SELECT
    Genre,
    ROUND(AVG(Avg_Rating),2) AS Avg_Rating
FROM netflix_customers
GROUP BY Genre
ORDER BY Avg_Rating DESC;
Q12. Account Status
SELECT
    Account_Status,
    COUNT(*) AS Customers
FROM netflix_customers
GROUP BY Account_Status;
________________________________________
# STEP 20 — Power BI
Import the clean Excel/SQL table into Power BI.
Dashboard Layout
Row 1 — KPI Cards
Card 1
•	Total Customers 
Card 2
•	Active Customers 
Card 3
•	Total Revenue 
Card 4
•	Average Watch Hours 
Card 5
•	Average Rating 
________________________________________
📈 Row 2 — Customer Analysis
Visual 1
Customers by Subscription Plan
Clustered Column Chart
Visual 2
Customers by Country
Map / Bar Chart
Visual 3
Account Status Distribution
Donut Chart
________________________________________
🎬 Row 3 — Content Analysis
Visual 4
Popular Genres
Bar Chart
Visual 5
Movie vs TV Show
Donut Chart
Visual 6
Top 10 Titles
Bar Chart
________________________________________
📱 Row 4 — Engagement Analysis
Visual 7
Watch Hours by Subscription Plan
Column Chart
Visual 8
Device Usage
Donut Chart
Visual 9
Monthly Sessions by Plan
Column Chart
Visual 10
Downloads by Genre
Bar Chart
________________________________________
# STEP 21 — Slicers
Slicers in the dashboard:
•	Country 
•	Subscription Plan 
•	Content Type 
•	Genre 
•	Device 
•	Payment Method 
•	Account Status 
•	Age Group 
•	Watch Category 
________________________________________
🧮 STEP 22 — Power BI DAX Measures
Total Customers
Total Customers =
COUNTROWS(netflix_customers)
Active Customers
Active Customers =
CALCULATE(
    [Total Customers],
    netflix_customers[Account_Status] = "Active"
)
Cancelled Customers
Cancelled Customers =
CALCULATE(
    [Total Customers],
    netflix_customers[Account_Status] = "Cancelled"
)
Total Revenue
Total Revenue =
SUM(netflix_customers[Customer_Revenue])
Average Watch Hours
Average Watch Hours =
AVERAGE(netflix_customers[Watch_Hours])
Average Rating
Average Rating =
AVERAGE(netflix_customers[Avg_Rating])
Average Sessions
Average Sessions =
AVERAGE(netflix_customers[Monthly_Sessions])
________________________________________
# STEP 23 — Final Business Insights

#Explaining the insights.
Example:
“The analysis shows which subscription plan is the most popular among customers.”
“Some countries may be important customer markets for Netflix.”
“Shows which genre is consumed the most.”
“How does the engagement of Premium customers compare with Basic customers.”
“Which device is dominant: Mobile, Smart TV or Laptop.”
“By identifying customers with high watch hours, Netflix can provide personalized recommendations for them.”
