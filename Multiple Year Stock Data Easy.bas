Attribute VB_Name = "Module1"
Sub multiple_year_stock_data()
For Each ws In Worksheets
    Dim WorksheetName As String
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    WorksheetName = ws.Name
    MsgBox WorksheetName
    
    Dim Ticker As String
    Dim Total_Volume As Double
    Total_Volume = 0
    
    Dim Ticker_Table_Row As Integer
    Ticker_Table_Rows = 2
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Total Stock Volume"
    Dim i As Long

    For i = 2 To LastRow
        Ticker = ws.Cells(i, 1).Value
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            ws.Range("I" & Ticker_Table_Rows).Value = Ticker
            ws.Range("J" & Ticker_Table_Rows).Value = Total_Volume
            Ticker_Table_Rows = Ticker_Table_Rows + 1
            Total_Volume = 0
        Else
            Total_Volume = Total_Volume + Cells(i, 7).Value
        End If
    Next i
Next

End Sub

