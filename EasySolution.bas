Attribute VB_Name = "Module1"
Sub StockMarket_EasySolution()

'Set initial variables
Dim Ticker As String
Dim Volume As LongLong
Dim Summary_Table As Integer


'Variables
Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"
Volume = 0
Summary_Table = 2

'Looping through the tickers
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

'Comparing rows values to obtain tickers
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        'Setting the Ticker and Volumen Value
        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        
        'Putting the Ticker and Volumen on the Summary_Table
        Range("I" & Summary_Table).Value = Ticker
        Range("J" & Summary_Table).Value = Volume
        Summary_Table = Summary_Table + 1
        Volume = 0
    
    Else
    'Adding the volumen to each Ticker
        Volume = Volume + Cells(i, 7).Value
    
    End If
Next i

End Sub

