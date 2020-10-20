# Automate MCI Report

## Description
The MCI Report is run manually on a monthly basis that requires various steps in excel to manipulate the data to get the required outcome. To automate the process I use the Pandas libary and Openpyxl. 

## Automation Tasks
- Ask user for file paths for enter, exit, previous enter, and previous exit.
- Ask user for date range for current report.
- Ask user where to save new reports.
- The files that are downloaded come as an old .xls file which are actually an html file.
- Determine if it is an .xls file then read it as a html and if it is an .xlsx file then read it as an excel file.
- Create Pandas Dataframes for each file.
- Remove last six rows because it is not needed for the files that end with xlsx.
- Keep rows that meet the Case Type criteria.
- Add a new column with the rank value of the Case Type.
- Add a new column for the rank value for the Case Status in the Exit file.
- Multisort the data in the by PID, Ranks, and Case Open Date.
- Delete the Rank columns.
- Check to see if the current report's PID is in the previous report and only keep the rows that were not in the previous report.
- Highlight the cell where the Custody Start Date was two months ago.
- Save the new data into a new sheet called  Dups Removed.
- Create seperate workbooks by Insurance compnany. Each workbook should have an Enter tab and exit tab.
- Add another tab to the Highmark workbook for all kids.
- Adjust the columns widths dynamically. 
