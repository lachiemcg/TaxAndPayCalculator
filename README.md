# Tax and Salary Sacrifice Calculator

This project provides an Excel-based tool to assist payroll teams in calculating the effects of salary sacrifice on employees' pay. It generates detailed reports that include a breakdown of taxable income, tax liabilities, and net pay before and after salary sacrificing. Additionally, the tool creates a pay schedule showing the adjustments for each pay cycle for the remainder of the financial year.

## How It Works
1. **Input Employee Details:**
    - The user inputs information such as the employee's name, financial year, annual salary, HECS-HELP debt status, payroll cycle, next payroll date, and the amount to be sacrificed per payroll cycle.
  
2. **Generate Calculations:**
    - Once all the inputs are filled, the user clicks the "Generate Calcs" button, and the VBA script runs the calculations.
  
3. **Generated Reports:**
    - Two reports are generated:
        - **[Employee Name]-Comparison**: Provides a detailed comparison between the original and salary sacrifice-adjusted financial details, including gross pay, total salary sacrifice, taxable income, income tax, HECS-HELP, Medicare Levy, and a summary of amounts paid to date and remaining to be paid.
        - **[Employee Name]-Pay Schedule**: Displays a detailed schedule of each remaining pay cycle, including pay dates, gross pay, amount sacrificed, taxable income, deductions (income tax, HECS-HELP, Medicare), and net pay after deductions.

## Setup Instructions
1. **Ensure Macros are Enabled:**
    - This tool requires VBA macros. Make sure to enable macros when opening the Excel file.

2. **User Input:**
    - Fill in the required fields (e.g., employee name, salary, payroll cycle) in the input section of the "Donation_Tax_Calc" worksheet.

3. **Generate Report:**
    - After inputting the details, click "Generate Calcs" to create both a comparison and a pay schedule worksheet.

4. **Saving and Sharing:**
    - Save the final report with an appropriate filename and securely share the file with payroll or relevant stakeholders.

## Features
- **Automated Income and Tax Calculations:** Calculates taxable income and tax liabilities both before and after salary sacrifice.
- **Detailed Reports:** Generates a comparison summary of the original vs. salary-sacrificed income, and a pay cycle schedule for the remaining year.
- **User-Friendly Interface:** Allows easy input and automatic generation of outputs with just one click.

## Troubleshooting
- Ensure the payroll cycle (fortnightly or monthly) and next payroll date are correctly entered to avoid errors.
- If the next payroll date falls outside the FY25 range, an error will be prompted.

## Files
- **SalarySacrificeCalculator.xlsm**: The Excel workbook containing the VBA script and the tool for salary sacrifice calculations.
