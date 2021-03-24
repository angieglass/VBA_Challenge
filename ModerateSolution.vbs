Attribute VB_Name = "Module2"
Sub StockMarket_Moderate()

'Set initial variables
Dim Ticker As String
Dim Volume As LongLong
Dim Summary_Table As Integer
Dim Closing As Double
Dim Opening As Double
Dim YearlyChange As Double
Dim PercentChange As Double


'Variables
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Volume = 0
Summary_Table = 2
Opening = Cells(2, 3).Value


'Looping through the tickers
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

'Comparing row values to obtain tickers
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    'Setting the variables values on the Summary_Table
        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        Closing = Cells(i, 6).Value
        YearlyChange = Closing - Opening
       
       If Opening = 0 Then
            PercentChange = 0
        Else
            PercentChange = YearlyChange / Opening
    End If

        'Putting the variables values on the Summary_Table

        Range("I" & Summary_Table).Value = Ticker
        Range("J" & Summary_Table).Value = YearlyChange
        Range("K" & Summary_Table).Value = PercentChange
        Range("L" & Summary_Table).Value = Volume
    
        Summary_Table = Summary_Table + 1
        Volume = 0
        Opening = Cells(i + 1, 3).Value
        
        'Format to percentage
        Range("K:K").NumberFormat = "0.00%"
    
    Else
        'Adding the volumen to each Ticker
        Volume = Volume + Cells(i, 7).Value
        
    End If
Next i

  'Format to YearlyChange
For i = 2 To lastrow
    If Cells(i, 10).Value >= 0 Then
        Cells(i, 10).Interior.ColorIndex = 43
    Else
        Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i


End Sub


