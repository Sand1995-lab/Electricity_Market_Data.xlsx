⚡ Electricity Market Data Automation – Python
📌 Overview
This project automates the process of fetching, processing, and analyzing electricity market data from the PJM wholesale market using Python. The script downloads electricity pricing data for 2024 & 2025, processes it, and updates an Excel file with a rolling year dataset and average calculations.

🔍 Features
✅ Automatic Data Fetching: Retrieves data from the PJM electricity market for 2024 & 2025
✅ Rolling Year Analysis: Combines historical data for trend analysis
✅ Data Cleaning & Processing: Filters and structures data for better insights
✅ Automatic Average Calculation: Computes key metrics and applies formatting
✅ Scheduled Execution: Runs twice daily at 5 AM & 5 PM using Python's schedule module
✅ Excel Integration: Saves processed data in an Excel file (Electricity_Market_Data.xlsx) with formatted output

🛠️ Technologies Used
Python for automation and data processing

Pandas for data manipulation

Requests for fetching data from the PJM market

OpenPyXL for writing & formatting Excel files

Schedule for setting up automated execution

🚀 How It Works
The script fetches electricity market price data for 2024 & 2025.

Data is stored in an Excel file (Electricity_Market_Data.xlsx) with separate sheets for each year.

A combined dataset for the last 365 days is created.

The script calculates averages for numeric fields and applies formatting (bold text & light blue background).

The process runs automatically twice daily (at 5 AM & 5 PM) using a scheduler.

📈 Expected Insights
Electricity price trends over time

Comparison of energy costs between years

Regional price variations and market trends
