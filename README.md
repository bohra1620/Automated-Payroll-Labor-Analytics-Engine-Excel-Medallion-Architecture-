# Automated-Payroll-Labor-Analytics-Engine-Excel-Medallion-Architecture-
A client required an administrative assistant to manually calculate monthly salaries for 50+ employees using raw, messy timesheet dumps. The manual process was prone to human error, struggled with missing clock-outs, and took hours to complete, yielding no broader business insights. The strict constraint: Only Excel could be used.

💡 The Solution
Instead of manual data entry, I engineered a zero-touch Medallion Architecture pipeline entirely within the native Excel ecosystem, utilizing Power Query (M) and Power Pivot (DAX).

Impact:

⏳ Time Saved: Reduced processing time from ~8 hours/month to a 10-second "Refresh All" click.

🎯 Accuracy: 100% automated calculation eliminating copy-paste errors.

📈 Business Intelligence: Added a dashboard highlighting excessive departmental overtime, saving the client estimated thousands in unoptimized labor costs.

🛠️ Architecture & Tech Stack
Constraint Adherence: 100% Native Excel (No macros required, safe for enterprise security policies).

Bronze Layer (Raw): Folder-based connection to messy system-generated CSV timesheets.

Silver Layer (Transform): Power Query (M-Code) utilized to normalize mixed datetime formats, handle null clock-outs via fallback logic, and structure tabular data.

Gold Layer (Semantic Model): Power Pivot utilized to build a relational model. DAX measures created to enforce complex business rules (1.5x Overtime after 8 hours daily, Weekend differentials).

🗂️ Project Repository Contents
Payroll_Engine_V1.xlsx: The final automated tool.

synthetic_data_gen.py: A Python script utilizing pandas and faker to generate the realistic, noisy dataset used to test the ETL pipeline.

/Data_Drop: Sample raw CSV files showcasing the system's ability to append multiple files dynamically.

🚀 How It Works
User drops weekly/monthly timesheet CSVs into the /Data_Drop folder.

User opens the Excel file and clicks Data > Refresh All.

Power Query automatically ingests, cleans, and updates the DAX model.

The Payroll tab instantly generates the precise payout amounts, while the Executive Dashboard updates labor cost trends.
