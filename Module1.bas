Attribute VB_Name = "Module1"
Sub LoopSheets()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim CurrentTicker As String
    Dim TotalVol As Double
    Dim Counter As Long
    Dim OpenPrice As Double
    
    'Loop through each worksheet in the workbook.
    For Each ws In Worksheets
    
        'Initializes the column headers for each worksheet in the workbook.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Column headers for summary statistics
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Initialize variables for the first stock entry
        CurrentTicker = ws.Cells(2, 1).Value
        'Initialize the total volume variable
        TotalVol = 0
        'Start a counter for the output row number
        Counter = 2
        'Store the opening price for the first ticker
        OpenPrice = ws.Cells(2, 3)
        'Determine the last row with data
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
'       'Loop through each row, starting from row 2 to the last row with data
        For i = 2 To LastRow
           
            If CurrentTicker = ws.Cells(i, 1).Value Then
                'Add the current row's volume to the total volume
                TotalVol = TotalVol + ws.Cells(i, 7).Value
                
            Else
                'Once we hit the new ticker, output the current ticker symbol and its calculated values
                ws.Cells(Counter, 9) = CurrentTicker
                ws.Cells(Counter, 10) = ws.Cells(i - 1, 6) - OpenPrice
                ws.Cells(Counter, 11) = ws.Cells(Counter, 10).Value / OpenPrice
                ws.Cells(Counter, 12) = TotalVol
                'Update variables for the new ticker
                OpenPrice = ws.Cells(i, 3)
                Counter = Counter + 1
                CurrentTicker = ws.Cells(i, 1).Value
                TotalVol = ws.Cells(i, 7)
            End If
        Next i
        
        Dim MaxValRow As Long
        Dim MinValRow As Long
        Dim MaxVolRow As Long
        
        'Used built in functions to find the maximum and minimum values
        ws.Range("Q2") = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
        ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
        ws.Range("Q4") = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        
        'Loop through the range to find the row numbers for the respective maximum and minimum values
        For j = 2 To LastRow
            If ws.Cells(j, "K").Value = ws.Range("Q2") Then
                MaxValRow = j
            End If
            
            If ws.Cells(j, "K").Value = ws.Range("Q3") Then
                MinValRow = j
            End If
            
            If ws.Cells(j, "L").Value = ws.Range("Q4") Then
                MaxVolRow = j
            End If
            
        Next j
        
        ws.Range("P2") = ws.Cells(MaxValRow, "I").Value
        ws.Range("P3") = ws.Cells(MinValRow, "I").Value
        ws.Range("P4") = ws.Cells(MaxVolRow, "I").Value
        
        'The code below formats the data in Yearly Change and Percentage change
        Dim LastRow1 As Long
        LastRow1 = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        With ws.Range("K2:K" & LastRow1)
            .NumberFormat = "0.00%"
        End With
        
        With ws.Range("J2:J" & LastRow1)
            .NumberFormat = "0.00"
            
            .FormatConditions.Delete
            
            With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0) 'Red
            End With
            
            With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
            .Interior.Color = RGB(57, 255, 20) ' Neon Green
            End With
            
        End With
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
    Next ws
End Sub
