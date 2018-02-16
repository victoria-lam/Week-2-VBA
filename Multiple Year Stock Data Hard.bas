Attribute VB_Name = "Module1"
Sub multiple_year_stock_data()
    
    'Define variables
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
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    open_amt = Cells(2, 3).Value
    close_amt = Cells(2, 6).Value
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    'Create a loop
    For i = 2 To LastRow
    
        'Conditionals
        'Find Ticker values for yearly change, percent change, and total stock
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            close_amt = Cells(i, 6).Value
            Yearly_Change = close_amt - open_amt
            
                If open_amt <> 0 Then
                    Percent_Change = Yearly_Change / open_amt
                End If
                
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
                'Assign colors to positive or negative yearly change
                If Yearly_Change > 0 Then
                    Cells(Ticker_Table_Rows, 10).Interior.Color = vbGreen
                ElseIf Yearly_Change <= 0 Then
                    Cells(Ticker_Table_Rows, 10).Interior.Color = vbRed
                End If
                
            Cells(Ticker_Table_Rows, 9).Value = Ticker
            Cells(Ticker_Table_Rows, 10).Value = Yearly_Change
            Cells(Ticker_Table_Rows, 11).Value = Percent_Change
            Cells(Ticker_Table_Rows, 11).NumberFormat = "0.00%"
            Cells(Ticker_Table_Rows, 12).Value = Total_Volume
            Ticker_Table_Rows = Ticker_Table_Rows + 1
            open_amt = Cells(i + 1, 3).Value
            Total_Volume = 0
        Else
            close_amt = Cells(i, 6).Value
            Total_Volume = Total_Volume + Cells(i, 7).Value
        End If
    Next i

   'Find greatest % increase/decrease and total volume
    Dim percent_range As Range
    Dim percent_max As Double
    Dim percent_min As Double
    Dim vol_range As Range
    Dim vol_val As Double
    Dim curr_val As Double
    Dim ticker_val As String
    
    Set percent_range = Range("K:K")
    Set vol_range = Range("L:L")
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    percent_max = 0
    percent_min = 0
    curr_val = 0
    LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To LastRow
        curr_val = Cells(i, 11)
        
        If curr_val > percent_max Then
            percent_max = curr_val
        End If
        
        If curr_val < percent_min Then
            percent_min = curr_val
        End If
    Next i
    
    curr_val = 0
    
    For i = 2 To 10
        curr_val = Cells(i, 12)
        
        If curr_val > vol_val Then
            vol_val = curr_val
        End If
    Next i
    
    'Find corresponding ticker value
    For i = 2 To LastRow
       If Range("K" & i).Value = Cells(2, 17).Value Then
            Cells(2, 16).Value = Range("I" & i).Value
            MsgBox Range("I" & i).Value
        End If
        
        If Range("K" & i).Value = Cells(3, 17).Value Then
            Cells(3, 16).Value = Range("I" & i).Value
            MsgBox Range("I" & i).Value
        End If
        
        If Range("L" & i).Value = Cells(4, 17).Value Then
            Cells(4, 16).Value = Range("I" & i).Value
            MsgBox Range("I" & i).Value
         End If
    Next i
    
    Cells(2, 17).Value = percent_max
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).Value = percent_min
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 17).Value = vol_val

End Sub



