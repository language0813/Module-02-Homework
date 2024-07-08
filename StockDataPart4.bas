Attribute VB_Name = "Module5"
Sub StockDataPart4()

' To loop across all the worksheets Q1, Q2, Q3, and Q4
Dim ws As Worksheet
For Each ws In Worksheets

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
OpenPrice = ws.Cells(2, 3).Value

' To print the first part headers to the specific cell for the first summary table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

' Use "do while" loop method to loop through column A while cells' value is not equal to empty
Do While ws.Cells(i, 1) <> ""

    ' Set condition "when the next cell in column A is not equal to the current cell"
    ' To locate the current ticker in column A and retrieve its information, and fill it in the summary table
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
        
        ClosePrice = ws.Cells(i, 6).Value
        QuarterlyChange = ClosePrice - OpenPrice
        ws.Cells(Ticker, 10).Value = QuarterlyChange
        
        PercentChange = (ClosePrice - OpenPrice) / OpenPrice
        ws.Cells(Ticker, 11).Value = PercentChange
                
        ws.Cells(TotalStockVol, 12).Value = ws.Cells(TotalStockVol, 12).Value + ws.Cells(i, 7).Value
        
        ' To reset the counters
        Ticker = Ticker + 1
        OpenPrice = ws.Cells(i + 1, 3).Value
        TotalStockVol = TotalStockVol + 1
        
    ' If the next cell in column A equal to the current cell
    ' Just add the current row's volumn to summary table
    Else
    
        ws.Cells(TotalStockVol, 12).Value = ws.Cells(TotalStockVol, 12).Value + ws.Cells(i, 7).Value
                   
    End If
    
i = i + 1
Loop



' To print the second part headers to the specific cell for second summary table
ws.Cells(2, 14).Value = "Greatest % increase"
ws.Cells(3, 14).Value = "Greatest % decrease"
ws.Cells(4, 14).Value = "Greatest total volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

' Set variables to hold values
Dim MaxPercentIncrease As Double
Dim MaxPercentDecrease As Double
Dim MaxTotalVol As Double
Dim LastRow As Integer

' Applied the syntax learned from the class content to find the last row
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

' To find the max and min numbers in column K, and max number in column L
MaxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
MaxPercentDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
MaxTotalVol = Application.WorksheetFunction.Max(ws.Range("L:L"))

' To loop from row 2 to the last row of the first summary table
For i = 2 To LastRow

' Set conditions to locate the tickers that have same % change and stock volume as MaxPercentIncrease
' MaxPercentDecrease, and MaxTotalVol. And print the values to the second summary table
    If ws.Cells(i, 11).Value = MaxPercentIncrease Then
        ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
        
    ElseIf ws.Cells(i, 11).Value = MaxPercentDecrease Then
        ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
    
    ElseIf ws.Cells(i, 12).Value = MaxTotalVol Then
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
        
    End If

Next i


' To organize the numbers' format for both summary tables
' And adjust the column width to match the data size by applying autofit
ws.Columns("J:J").NumberFormat = "0.00"
ws.Columns("K:K").Style = "Percent"
ws.Columns("K:K").NumberFormat = "0.00%"
ws.Range("P2").Style = "Percent"
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").Style = "Percent"
ws.Range("P3").NumberFormat = "0.00%"
ws.Columns("I:P").EntireColumn.EntireColumn.AutoFit
    
' Set conditions to color the cells to the quarterly change and percent change column
For i = 2 To LastRow

    If ws.Cells(i, 10) > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 43
        
    ElseIf ws.Cells(i, 10) < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 46
        
    End If
    
Next i

For i = 2 To LastRow
 
    If ws.Cells(i, 11) > 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 43
        
    ElseIf ws.Cells(i, 11) < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 46
        
    End If
    
Next i

Next ws

End Sub

