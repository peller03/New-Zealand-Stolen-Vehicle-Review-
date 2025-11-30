# New-Zealand-Stolen-Vehicle-Review-
----
## Table of Contents
- [Project Overview](#project-overview)
- [Data Sources](#data-sources)
- [Tools Used](#tools-used)
- [Data Cleaning/Preparation](#data-cleaning--preparation)
- [Exploratory Data Analysis Performed](#exploratory-data-analysis-performed)
- [Data Analysis](#data-analysis)
- [Results/Findings](#results--findings)
- [Recommendations](#recommendations)
- [Limitations](#limitations)

### Project Overview
This project analyzes stolen vehicle trends across New Zealand using Excel for data cleaning, transformation, and visualization.
The goal is to understand which vehicles are most stolen, high-risk regions, seasonal theft patterns, and how demographic factors like population density influence crime.
 

### Data Sources
TDI (The Data Immersed) Group Provided the dataset as a ZIP file containing multiple structured datasets related to stolen vehicles.

Supporting files such as:

- stolen_vehicles.csv (Main data)

- make_details.csv

- locations.csv
 
- stolen_vehicle_db (data dictionary)

### Tools Used
- Microsoft Excel

  - Data cleaning

  - VLOOKUP-based data merging

  - PivotTables & PivotCharts

  - Dashboard creation

- Excel Formulas: VLOOKUP, COUNT, AVERAGE, sorting, filtering

- Basic statistical analysis

- Data visualization tools available in Excel

 
### Data Cleaning / Preparation

Key steps in preparing the dataset:

1. Removed duplicates:
Ensured each stolen vehicle report appeared only once.

3. Handled missing values:
Cleaned null or blank entries in location, make details, and other critical fields.

4. Standardized text fields:
Ensured consistency in spellings for regions, cities, vehicle types, and colors.

5. Corrected data types

   - Dates standardized

   - Model years cleaned

   - Numeric values validated

5. Merged datasets using VLOOKUP

- Added key attributes such as:

   - make_name

   - make_type

   - region

   - population

   - density

6. Validated against the data dictionary:
Used stolen_vehicle_db to ensure correct interpretation of each column.

### Exploratory Data Analysis Performed
The EDA focused on understanding:

#### Key Metrics (KPIs)

- Total stolen vehicles: 4,527

- Most stolen vehicle type: Station Wagon

- Region with most cases: Auckland

- Region with highest population: Auckland

- Peak theft month: March 2022

#### Major EDA Visuals

1. Most Stolen Vehicle Types

    - Station Wagons & Hatchbacks lead

     - Specialized vehicles like tractors appear least stolen

2. Monthly Theft Trends

    - Peak in March 2022 (1,049 cases)

    - Noticeable seasonal drop between Oct 2021 – Jan 2022

3. Vehicle Model Year Distribution

    - Older models (1976–1994 & 1994–2012) most vulnerable

    - Newer models less stolen → stronger security technologies

4. Make Type Distribution

    - Standard vehicles: 96%

    - Luxury vehicles: 4%

    - Indicates thieves prefer common, easy-to-resell models

5. Regional Theft Patterns

    - Auckland highest

    - Canterbury, Waikato & Bay of Plenty also high

    - Rural regions like Southland & Ruapehu lowest

6. Population vs Theft Correlation

    - Higher population → higher theft frequency

    - But exceptions exist, showing other socio-economic factors may influence crime

7. Vehicle Color Distribution

    - Most stolen colors: White, Silver, Blue

    - Least stolen: Pink, Purple, Cream
 
### Data Analysis

Excel formulas and pivot tools used:

- COUNT: Total number of stolen vehicles

- AVERAGE: Average population per region

- VLOOKUP: Merge datasets across multiple sources

Merged make_details data with stolen_vehicles excel record

#### a. VLOOKUP to Get Make Name 
```
=VLOOKUP(C2,make_details.xlsx!$A$1:$B$139,2,FALSE)

```

#### b. VLOOKUP to Get Make Name 
Merged locations data with stolen_vehicles excel record

```
=VLOOKUP(I2,locations.xlsx!$A$2:$E$17,2,FALSE)
```

- PivotTables:

   - Theft by region

   - Theft by vehicle type

   - Theft by color

   - Theft by model year

   - Theft over time

Insights driven by this analysis include demographic influence, vehicle characteristics, and geographic patterns contributing to theft.


### Results / Findings

Key insights summarize the crime landscape:

1. Station Wagons are the most stolen vehicle type

2. Auckland is the highest-risk region

3. March 2022 is the peak theft month

4. Older vehicles are far more vulnerable

5. Standard car types dominate theft cases

6. High population regions experience more theft

7. Neutral colors (white, silver) are targeted more

8. Theft patterns show both seasonal and regional behavior

### Recommendations
Based on the analysis:

1. Increase surveillance in Auckland & other high-risk regions

2. Strengthen security for older vehicles (immobilizers, steering locks)

3. Awareness campaigns for owners of high-risk vehicle types

4. Deploy targeted policing during peak months (e.g., March)

5. Encourage anti-theft technology adoption (GPS tracking systems)

### Limitations
- No socio-economic or crime-rate data included

- Dataset does not include recovery outcomes

- Population metrics are limited to region-level

- No data on offender characteristics

- Time period limited  trends may shift beyond dataset

  
