Attribute VB_Name = "VBAChallengeModule"
Sub VBA_challenge()
    'Creating variable to hold the counter
    Dim i As Long
    'Creating ws variable for loop
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        'Setting initial variable to hold the ticker name and total
        Dim Ticker_Name As String
        
        Dim Ticker_Total As Double
        Ticker_Total = 0
        
        'Setting up the open variable
        Dim Open_Ticker As Double
        Dim Close_Ticker As Double
        
        Open_Ticker = ws.Cells(2, 3).Value
        
        'Setting up the summary table where it'll print out all of the totals
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2 'Setting at 2 to compensate for the header
        
        'Find the end row to make it able to loop through all of the tickers on the page
        endRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Adding in the Column Names
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Adding in Addtl. Column and Row Names for "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        For i = 2 To endRow
        'Checking to see if we are still in the same ticker and moving onto the next if we are not.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_Name = ws.Cells(i, 1).Value
                Ticker_Total = ws.Cells(i, 7).Value + Ticker_Total
                Close_Ticker = ws.Cells(i, 6).Value
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ws.Range("J" & Summary_Table_Row).Value = (Close_Ticker - Open_Ticker)
                ws.Range("K" & Summary_Table_Row).Value = ((Close_Ticker - Open_Ticker) / Open_Ticker)
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Ticker_Total = 0
                
                Open_Ticker = ws.Cells(i + 1, 3).Value
                
            Else
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            End If
        Next i
        
        'Formatting Column J
        endRow2 = ws.Cells(Rows.Count, "J").End(xlUp).Row
        For i = 2 To endRow2
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
            
            End If
         Next i
         
         'Finding Greatest % Increase, Greatest % Decrease, and Greatest Volume
         Dim Max_PercentTicker As String
         Dim Min_PercentTicker As String
         
         Dim Max_VolumeTicker As String
         
         Max_Percent = 0
         For i = 2 To endRow2 'Finding greatest % Increase looping through endRow2
            If ws.Cells(i, 11).Value > Max_Percent Then
            Max_Percent = ws.Cells(i, 11).Value
            Max_PercentTicker = ws.Cells(i, 9).Value
            End If
         Next i
         
         Min_Percent = 0
         For i = 2 To endRow2 'Finding greatest % Decrease looping through endRow2
               If ws.Cells(i, 11).Value < Min_Percent Then
               Min_Percent = ws.Cells(i, 11).Value
               Min_PercentTicker = ws.Cells(i, 9).Value
               End If
         Next i
         
         Max_Volume = 0
         For i = 2 To endRow2 'Finding the Greatest Trade Volume looping through endRow2
               If ws.Cells(i, 12).Value > Max_Volume Then
               Max_Volume = ws.Cells(i, 12).Value
               Max_VolumeTicker = ws.Cells(i, 9).Value
               End If
         Next i
         
         'Pasting and formatting the values in their own table
         
         ws.Cells(2, 16).Value = Max_PercentTicker
         ws.Cells(2, 17).Value = Max_Percent
         ws.Cells(2, 17).NumberFormat = "0.00%"
         
         ws.Cells(3, 16).Value = Min_PercentTicker
         ws.Cells(3, 17).Value = Min_Percent
         ws.Cells(3, 17).NumberFormat = "0.00%"
         
         ws.Cells(4, 16).Value = Max_VolumeTicker
         ws.Cells(4, 17).Value = Max_Volume
         
         ws.Range("I:Q").Columns.AutoFit
         
    Next ws
    
    MsgBox ("All pages searched")

End Sub

