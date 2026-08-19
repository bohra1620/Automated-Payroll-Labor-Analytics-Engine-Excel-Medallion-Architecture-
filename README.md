# 📊 Automated Payroll & Labor Analytics Engine (Excel Medallion Architecture)

## 📌 Business Problem
A client required an administrative assistant to manually calculate monthly salaries for 50+ employees using raw, messy timesheet dumps. The manual process was prone to human error, struggled with missing clock-outs, and took hours to complete, yielding no broader business insights.

## 💡 The Solution
Instead of manual data entry, I engineered a zero-touch **Medallion Architecture pipeline entirely within the native Excel ecosystem**, utilizing Power Query (M), Power Pivot (DAX), and VBA for final-mile automation. 

**Impact:**
* ⏳ **Time Saved:** Reduced processing time from ~8 hours/month to a 10-second "Refresh All" click.
* 🎯 **Accuracy:** 100% automated calculation eliminating copy-paste errors.
* 📈 **Business Intelligence:** Added a dashboard highlighting excessive departmental overtime, saving the client estimated thousands in unoptimized labor costs.

## 🛠️ Architecture & Tech Stack
* **Bronze Layer (Raw):** Folder-based connection to messy system-generated CSV timesheets.
* **Silver Layer (Transform):** Power Query (M-Code) utilized to normalize mixed datetime formats, handle `null` clock-outs via fallback logic, and structure tabular data.
* **Gold Layer (Semantic Model):** Power Pivot utilized to build a relational model. DAX measures created to enforce complex business rules (1.5x Overtime after 8 hours daily, Weekend differentials).
* **Automation Layer:** VBA scripts to refresh the pipeline and batch-export PDF paychecks.

## 🗂️ Project Repository Contents
* `Payroll_Engine_V1.xlsm`: The final automated tool (Macro-enabled).
* `synthetic_data_gen.py`: A Python script utilizing `pandas` and `faker` to generate the realistic, noisy dataset used to test the ETL pipeline.
* `/Data_Drop`: Sample raw CSV files showcasing the system's ability to append multiple files dynamically.

## 🚀 How It Works
1. User drops weekly/monthly timesheet CSVs into the `/Data_Drop` folder.
2. User opens the Excel file and clicks the **Refresh Dashboard** macro button.
3. Power Query automatically ingests, cleans, and updates the DAX model.
4. The Payroll tab instantly generates the precise payout amounts, while the Executive Dashboard updates labor cost trends.
5. The user clicks **Generate Paychecks** to export individual PDF pay slips for every employee.

---

## 🧠 Under the Hood: Power Query, DAX, & VBA Logic

### 1. Power Query (M-Code) Transformations
**Handling Missing Punch-Outs:**
To handle scenarios where an employee forgot to clock out, I capped their pay at a standard 8-hour shift by adding a Custom Column (`Clean_Punch_Out`).
```powerquery
if [Punch_Out] = null then [Punch_In] + #duration(0, 8, 0, 0) else [Punch_Out] 
```

**Calculating Hours Worked:**
With the clean punch-out times, calculating the exact duration was done via another Custom Column (`Hours_Worked`).
```powerquery
Duration.TotalHours([Clean_Punch_Out] - [Punch_In])
```

### 2. Data Modeling & Calculated Columns (Power Pivot)
**Handling Exceptions (The Trick):**
Because Power Query mathematically fixed the nulls, I built a quick flag column inside Power Pivot's `Fact_Timesheets` table to catch these exceptions for reporting.
```dax
Exception_Flag = IF(ISBLANK([Punch_Out]), "Missing Punch", "OK")
```

**Time Intelligence Grouping:**
Calculated columns added to group the data for the Executive Dashboard.
```dax
Day_Name = FORMAT([Shift_Date], "dddd")
Week_Number = WEEKNUM([Shift_Date])
```

### 3. Core Analytics Measures (DAX)
These measures drive the executive dashboard and the automated payroll list.
```dax
Total Hours = SUM(Timesheets[Hours_Worked])

Regular Hours = IF(Timesheets[Hours_Worked] > 8, 8, Timesheets[Hours_Worked])

Overtime Hours = IF(Timesheets[Hours_Worked] > 8, Timesheets[Hours_Worked] - 8, 0)

Total Pay = 
    (SUMX(Timesheets, [Regular Hours]) * RELATED(Employee[Hourly_Rate])) + 
    (SUMX(Timesheets, [Overtime Hours]) * RELATED(Employee[Hourly_Rate]) * 1.5)
```

### 4. VBA Automation (The Final Mile)
To make the tool completely foolproof for the end user and provide maximum utility, I implemented VBA scripts tied to interactive dashboard buttons:

*   **One-Click Data Refresh Macro:** A script assigned to a button on the dashboard that automatically refreshes all Power Query data connections, updates the DAX data model, and refreshes all pivot tables and slicers in the correct sequence. 
*   **Automated PDF Paycheck Generator:** A script that loops through the finalized payroll list and automatically exports individual PDF pay slips for each employee. This eliminates the need to manually create and print or save individual records.
*   **Creative Executive Summary:** Integrated a dynamic text box summarizing top-level metrics for quick executive reading without having to parse the charts.
