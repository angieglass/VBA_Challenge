Attribute VB_Name = "Module3"
Sub StockMarket_Moderate_OneRun()

'Set initial variables
Dim Ticker As String
Dim Volume As LongLong
Dim Summary_Table As Integer
Dim Closing As Double
Dim Opening As Double
Dim YearlyChange As Double
Dim PercentChange As Double


For Each ws In Worksheets

'Variables
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
Volume = 0
Summary_Table = 2
Opening = ws.Cells(2, 3).Value


'Looping through the tickers
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

'Comparing row values to obtain tickers
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    'Setting the variables values on the Summary_Table
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        Closing = ws.Cells(i, 6).Value
        YearlyChange = Closing - Opening
       
       If Opening = 0 Then
            PercentChange = 0
        Else
            PercentChange = YearlyChange / Opening
    End If

        'Putting the variables values on the Summary_Table

        ws.Range("I" & Summary_Table).Value = Ticker
        ws.Range("J" & Summary_Table).Value = YearlyChange
        ws.Range("K" & Summary_Table).Value = PercentChange
        ws.Range("L" & Summary_Table).Value = Volume
    
        Summary_Table = Summary_Table + 1
        Volume = 0
        Opening = ws.Cells(i + 1, 3).Value
        
        'Format to percentage
        ws.Range("K:K").NumberFormat = "0.00%"
    
    Else
        'Adding the volumen to each Ticker
        Volume = Volume + ws.Cells(i, 7).Value
        
    End If
Next i

  'Format to YearlyChange
For i = 2 To lastrow
    If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 43
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i

Next ws

End Sub


