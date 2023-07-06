Attribute VB_Name = "Module1"
Option Explicit


Sub StockMarketAnalysis()

'Create a script that loops through all the stocks for one year and outputs the following information:
    Dim sheetObject As Worksheet
    For Each sheetObject In Worksheets
    
    Dim WorkSheetName As String
    Dim StockCount As Long
    Dim LastRow1 As Long
    Dim Lastrow2 As Long
    Dim i As Long
    Dim j As Long
    Dim PerChange As Double
    Dim GreatInc As Double
    Dim GreatDec As Double
    Dim GreatVol As Double
    
    WorkSheetName = sheetObject.Name
    
'Headers
    sheetObject.Cells(1, 9).Value = "Ticker"
    sheetObject.Cells(1, 10).Value = "Yearly Change"
    sheetObject.Cells(1, 11).Value = "Percent Change"
    sheetObject.Cells(1, 12).Value = "Total Stock Volume"
    sheetObject.Cells(1, 16).Value = "Ticker"
    sheetObject.Cells(1, 17).Value = "Value"
    sheetObject.Cells(2, 15).Value = "Greatest % Increase"
    sheetObject.Cells(3, 15).Value = "Greatest % Decrease"
    sheetObject.Cells(4, 15).Value = "Greatest Total Volume"
    
    StockCount = 2
    j = 2
    
    LastRow1 = sheetObject.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow1

'Ticker symbol output
    If sheetObject.Cells(i + 1, 1).Value <> sheetObject.Cells(i, 1).Value Then
    sheetObject.Cells(StockCount, 9).Value = sheetObject.Cells(i, 1).Value
    
'Annual change from the opening price at the beginning of a given year to the closing price at the end of that year
    sheetObject.Cells(StockCount, 10).Value = sheetObject.Cells(i, 6).Value - sheetObject.Cells(i, 3).Value

'Conditional Format
    If sheetObject.Cells(StockCount, 10).Value < 0 Then
    sheetObject.Cells(StockCount, 10).Interior.ColorIndex = 3
    Else
    sheetObject.Cells(StockCount, 10).Interior.ColorIndex = 4
    
    End If
    
'Price Change
    If sheetObject.Cells(j, 3).Value <> 0 Then
    PerChange = ((sheetObject.Cells(i, 6).Value - sheetObject.Cells(j, 3).Value) / sheetObject.Cells(j, 3).Value)
    
'Format
    sheetObject.Cells(StockCount, 11).Value = Format(PerChange, "Percent")
    Else
    sheetObject.Cells(StockCount, 11).Value = Format(0, "Percent")
    
    End If
    
'Total Volume
    sheetObject.Cells(StockCount, 12).Value = WorksheetFunction.Sum(Range(sheetObject.Cells(j, 7), sheetObject.Cells(i, 7)))
    
    StockCount = StockCount + 1
    j = i + 1
    
    End If
    
    Next i
    
    Lastrow2 = sheetObject.Cells(Rows.Count, 9).End(xlUp).Row
  
'Functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"

'Summaries
    GreatVol = sheetObject.Cells(2, 12).Value
    GreatInc = sheetObject.Cells(2, 11).Value
    GreatDec = sheetObject.Cells(2, 11).Value
    
    For i = 2 To Lastrow2
    
'Total Greatest Volume
    If sheetObject.Cells(i, 12).Value > GreatVol Then
    GreatVol = sheetObject.Cells(i, 12).Value
    sheetObject.Cells(4, 16).Value = sheetObject.Cells(i, 9).Value
    Else
    GreatVol = GreatVol
    
    End If
    
'Total Greatest Increase
    If sheetObject.Cells(i, 11).Value > GreatInc Then
    GreatInc = sheetObject.Cells(i, 11).Value
    sheetObject.Cells(2, 16).Value = sheetObject.Cells(i, 9).Value
    Else
    GreatInc = GreatInc
    
    End If
    
'Total Greatest Decrease
    If sheetObject.Cells(i, 11).Value < GreatDec Then
    GreatDec = sheetObject.Cells(i, 11).Value
    sheetObject.Cells(3, 16).Value = sheetObject.Cells(i, 9).Value
    Else
    GreatDec = GreatDec
    
    End If
    
'Format
    sheetObject.Cells(2, 17).Value = Format(GreatInc, "Percent")
    sheetObject.Cells(3, 17).Value = Format(GreatDec, "Percent")
    sheetObject.Cells(4, 17).Value = Format(GreatVol, "Scientific")
    
    Next i
    
    sheetObject.Columns("A:Z").EntireColumn.AutoFit
    
    Next sheetObject


End Sub
