Attribute VB_Name = "Module1"
Sub multiple_year_stock_data():



Dim Ticker As String
Dim YearlyChange, PercentChange, TotalStockVolume, OpeningPrice, ClosingPrice, HighPrice, LowPrice, Volume As Double
Dim LastRow, SummaryRow As Long
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Activate
    

percent_change = 0

Dim summary As Integer
summary = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = " Greatest % Increase "
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Value"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"




LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

worksheetname = ws.Name

For i = 2 To LastRow


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
        
        'Date = Cells(i, 2).Value
        
        OpeningPrice = Cells(i, 3).Value
        
        HighPrice = Cells(i, 4).Value
        
        LowPrice = Cells(i, 5).Value
        
        ClosingPrice = Cells(i, 6).Value
        
        Volume = Cells(i, 7).Value
        
        YearlyChange = ClosingPrice - OpeningPrice
        
        PercentChange = YearlyChange / OpeningPrice * 100
        
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        
        Range("i" & summary).Value = Ticker
        
        Range("j" & summary).Value = YearlyChange
        
        Range("k" & summary).Value = PercentChange
        
        Range("l" & summary).Value = TotalStockVolume
        
        summary = summary + 1
        
        PercentChange = 0
        
    Else
        
        PercentChange = PercentChange = Cells(i, 3).Value
        
        
    End If
    
Next i

Next ws

End Sub

