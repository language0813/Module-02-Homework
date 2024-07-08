Attribute VB_Name = "Module2"
Sub StockDataPart1()

' Create variables to hold the counters and names
Dim i As Long
Dim Ticker As Long
Dim QuarterlyChange As Double
Dim PercentChange As Double
Dim TotalStockVol As Long
Dim OpenPrice As Double
Dim ClosePrice As Double

' Set initial value to the counters
i = 2
Ticker = 2
TotalStockVol = 2
OpenPrice = Cells(2, 3).Value

' To print the first part headers to the specific cell for the first summary table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quarterly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Use "do while" loop method to loop through column A while cells' value is not equal to empty
Do While Cells(i, 1) <> ""

    ' Set condition "when the next cell in column A is not equal to the current cell"
    ' To locate the current ticker in column A and retrieve its information, and fill it in the summary table
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Cells(Ticker, 9).Value = Cells(i, 1).Value
        
        ClosePrice = Cells(i, 6).Value
        QuarterlyChange = ClosePrice - OpenPrice
        Cells(Ticker, 10).Value = QuarterlyChange
        
        PercentChange = (ClosePrice - OpenPrice) / OpenPrice
        Cells(Ticker, 11).Value = PercentChange
                
        Cells(TotalStockVol, 12).Value = Cells(TotalStockVol, 12).Value + Cells(i, 7).Value
        
        ' To reset the counters
        Ticker = Ticker + 1
        OpenPrice = Cells(i + 1, 3).Value
        TotalStockVol = TotalStockVol + 1
        
    ' If the next cell in column A equal to the current cell
    ' Just add the current row's volumn to summary table
    Else
    
        Cells(TotalStockVol, 12).Value = Cells(TotalStockVol, 12).Value + Cells(i, 7).Value
                   
    End If
    
i = i + 1
Loop

End Sub

