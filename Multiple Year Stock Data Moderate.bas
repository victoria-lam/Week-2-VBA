Attribute VB_Name = "Module1"
Sub multiple_year_stock_data()
'Create loop for all worksheets
For Each ws In Worksheets
    'Declare the variables
    Dim WorksheetName As String
    WorksheetName = ws.Name
    MsgBox WorksheetName
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim open_amt As Double
    Dim close_amt As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim i As Long
    Dim Ticker_Table_Rows As Integer
    
    'Values for variables
    Total_Volume = 0
    Ticker_Table_Rows = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    open_amt = ws.Cells(2, 3).Value
    close_amt = ws.Cells(2, 6).Value
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    'Create a loop
    For i = 2 To LastRow
    
        'Conditionals
        'Find Ticker values for yearly change, percent change, and total stock
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            close_amt = ws.Cells(i, 6).Value
            Yearly_Change = close_amt - open_amt
            
                If open_amt <> 0 Then
                    Percent_Change = Yearly_Change / open_amt
                End If
                
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
                'Assign colors to positive or negative yearly change
                If Yearly_Change > 0 Then
                    ws.Cells(Ticker_Table_Rows, 10).Interior.Color = vbGreen
                ElseIf Yearly_Change <= 0 Then
                    ws.Cells(Ticker_Table_Rows, 10).Interior.Color = vbRed
                End If
                
            ws.Cells(Ticker_Table_Rows, 9).Value = Ticker
            ws.Cells(Ticker_Table_Rows, 10).Value = Yearly_Change
            ws.Cells(Ticker_Table_Rows, 11).Value = Percent_Change
            ws.Cells(Ticker_Table_Rows, 11).NumberFormat = "0.00%"
            ws.Cells(Ticker_Table_Rows, 12).Value = Total_Volume
            Ticker_Table_Rows = Ticker_Table_Rows + 1
            open_amt = Cells(i + 1, 3).Value
            Total_Volume = 0
        Else
            close_amt = Cells(i, 6).Value
            Total_Volume = Total_Volume + Cells(i, 7).Value
        End If
    Next i

Next
End Sub



