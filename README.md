#Module 2 Challenge

##Problem
Create a script that loops through all the stocks for one year and outputs the following information:
•	The ticker symbol
•	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
•	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
•	The total stock volume of the stock
•	Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
•	Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

##Location of the code and Filename

**FileName:** vba-challenge.vba

##Screenshot File Location and names

ScreenShot-2018-worksheet: 2018 Worksheet Screenshot.png
ScreenShot-2019-worksheet: 2019 Worksheet Screenshot.png
ScreenShot-2020-worksheet: 2020 Worksheet Screenshot.png

##Explanation of the code

In code there are two loops. The outer loop goes through the worksheets in the workbook. The inner loop goes through each row of the worksheet to identify open value, close value,
and total volume for a ticker in a given year. We then use open and close value to determine the yearly and percent change. These values are then written to a separate summary
table. In the same loop, we also track the greatest percent increase, greatest percent decrease and greatest total volume. We then update these values in a third summary table. To
give separate visual representation for positive and negative change we also did conditional formatting and used red color for negative yearly change and percent change and greem
for positive yearly change and percent change.

##Learnings
There were lot of learning while solving this problem. Some important ones are as follows:
1.	Using loops to go through data in and across worksheets.
2.	Using If-else conditions to determine state of the data, like when ticker values are changing, when to change the greatest increase or decrease values, and take actions accordingly.
3.	Using variables to track the values that needs to be updated in the Summary Tables.
4.	Programmatically determine values of number of rows and number of sheets so that the same solution can be used for any number of rows and sheets.
5.  Used conditional formating to assign colors to positive and negative values.
