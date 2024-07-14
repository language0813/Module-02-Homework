# Module-02-Homework

In this VBA challenge, I separated the project into four steps and added additional script on top of each previous step's script. Therefore, there are four separate VBA script files called "StockDataPart1", "StockDataPart2", "StockDataPart3", and "StockDataPart4".

I saved VBA script files by right-clicking on the Modules, clicking on "Export Files". I could only save as ".bas" file, and I'm able to open ".bas" files in VS Code.

Please see below for details regarding each step:

**Part 1:**

In part 1, I worked on getting info for first summary table in "Q1" worksheet, which includes Ticker, QuarterlyChange, PercentChange, and TotalStockVolume. I successfully retrieved the ticker symbol, volume of stock, and close price based on what I've learned from the class. But I had trouble about setting open price, and had to do research on the internet. I was able to find useful resource that helped me with the logic for setting the open price. The resource is listed below.

<https://github.com/shrawantee/VBA-Scripting---Stock-Market-Analysis/blob/master/HW2_Moderate_DS.vbs>

**Part 2:**

In part 2, I added scripts to obtain info for the second summary table in the "Q1" worksheet, which includes Greatest%increase, Greatest%decrease, and GreatestTotalVolume. I referred to code from the website Stackoverflow to find the maximum value of a column. So that I could use the maximum and minimum values to locate the correct ticker.

<https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba>

**Part 3:**

After creating two summary tables in the previous steps, I organized the worksheet's format in this part. I recorded a macro using relative references to refer to the formatting syntax. I adjusted the number's format, style, and column's width to match the data size by applying autofit.

Additionally, I added scripts to apply the conditional formatting in this part. I referred to below two websites for code to set the cell's background color and excel color indexes.

<https://www.excel-easy.com/vba/examples/background-colors.html>

<http://dmcritchie.mvps.org/excel/colors.htm>

**Part 4:**

In the last part, I added script to loop codes across all the worksheets. I successfully completed it through what I've learned from the class.

Thank you for your time review this!
