# VBA-challenge
This challenge required us to process and analyze multi year stock data using VBA in Excel
# Background
For this project, I created a VBA (Visual Basic) script to analyze some stock market data. The data is inside a Microsoft Excel Spread Sheet and includes stock data for 4 Quarters of 2022. Each Quarter is in a separate sheet in the Excel sheet
# Testing
For the testing purpose the script was ran through the alphabetical testing and then on multi year stock data.
# About the Script
The Script is under the name VBA script which I was able to generate with the help of the my friends.
Open the Developer Tab:

If the Developer tab is not already visible, enable it by going to File > Options > Customize Ribbon and check the box next to "Developer".
Open the Visual Basic Editor:

Click the Developer tab in the ribbon, then click Visual Basic to open the Visual Basic for Applications (VBA) editor.
Import the Script:

In the Visual Basic editor, click on File > Import File and select the MultipleYearStockData file from this repository to import it into the editor.
Run the Script:

Open the MultipleYearStockData file within the VBA editor.
Click the Run Macro button (green play icon) in the toolbar to execute the script.
Important Notes:
The script will take some time to run, as it processes data for every sheet in the workbook. There's no need to run it more than once.
As the script executes, it loops through the stock data for each year and computes the following information:
Ticker Symbol: The stock's unique identifier.
Yearly Change: The difference between the opening price at the start of the year and the closing price at the end of the year.
Percent Change: The percentage change between the opening and closing prices.
Total Stock Volume: The total trading volume of the stock during the year.
Conditional Formatting:

The script will apply conditional formatting to highlight positive yearly changes in green and negative yearly changes in red.
Final Results:

Once the script finishes running, it will identify:
The stock with the greatest percent increase.
The stock with the greatest percent decrease.
The stock with the highest total volume for the year.
