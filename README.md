# VBA-challenge
KU Boot Camp Week 2 Challenge - Stock Analysis VBA Script

## Introduction
This VBA script is designed to analyze stock data across consecutive years. It leverages three primary For Loops to aggregate and calculate key financial indicators for a variety of stocks. The script dynamically accommodates data from different years by incrementing the year as it processes through each worksheet. The output includes the ticker symbol, yearly change, percentage change, and total stock volume, along with the greatest increase, decrease, and total volume metrics.

## Overview of For Loops

### First For Loop - Stock Ticker and Volume Calculation
- Sets initial variables for calculating ticker symbol and volume.
- Sets an initial stock volume to zero to ensure accurate summation.
- Prints column names for data headers.
- Compares adjacent cells to identify unique ticker symbols and prints the output.
- Aggregates total stock volume per ticker and prints the output.
- Resets total stock volume to zero to ensure accurate per-stock aggregation.

### Second For Loop - Yearly and Percent Change Calculation
- Sets initial variables for calculating yearly and percent change.
- Identify and store the opening price at the beginning of the year and the closing price at the end of the year.
- Calculates and prints the yearly change and percentage change.
- Ensures proper formatting for yearly change (red and green cells) and percentage display.

### Third For Loop - Greatest Increase, Decrease, and Volume Calculation
- Sets initial variables for greatest volume, increase, and decrease.
- Iterates through rows to determine the greatest values in each category and performs the comparison.
- Formats and prints the calculated greatest values in a summary table.

## Adjustments for Multi-Year Analysis
- The script runs through each sheet, increments the year by 1, and replicates the stock analysis for each subsequent year.

## Usage
1. Open the Excel workbook titled Multiple_year_stock_data.xlsm, with the stock data organized by year in separate worksheets.
2. Access the VBA editor by navigating to the 'Developer' tab, then clicking on the 'Visual Basic' icon.
3. Run the script by pressing 'Run Sub'.
4. Review the analysis results presented on each worksheet, which will include the ticker symbol, yearly change, percentage change, and total stock volume, along with the greatest increase, decrease, and total volume metrics.

## Contribution
Contributions are welcome. To propose changes, please reach out with your suggestions or open an issue in the project repository.

## Credits
Developed by Albert J. Lee. Inspired by coursework and collaborative learning sessions.
