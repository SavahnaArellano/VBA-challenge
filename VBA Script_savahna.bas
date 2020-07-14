Attribute VB_Name = "Module1"


Sub stock_analysis()

    Dim tot_volume As Variant
    Dim RowCount As Long
    Dim yr_change As Variant
    Dim j As Integer
    Dim strt_pr As Variant
    Dim end_pr As Variant
    Dim rowNo As Integer
    Dim perChange As Variant

'To insert column title
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Year Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"


'j = 0
    tot_volume = 0
'yr_change = 0
'strt_pr = 2

' get the row number of the last row with data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    enterCellRow = 1


' 701937 and 78
    For i = 1 To RowCount
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(enterCellRow + 1, 9).Value = Cells(i + 1, 1).Value
            
            If i > 1 Then
                tot_volume = tot_volume + Cells(i, 7).Value
                Cells(enterCellRow, 12).Value = tot_volume
                tot_volume = 0
            End If
            
            If strt_pr Then
                end_pr = Cells(i, 6)
                yr_change = strt_pr - end_pr
                Cells(enterCellRow, 10) = yr_change
                
            If yr_change > 0 Then
                Cells(enterCellRow, 10).Interior.ColorIndex = 4
            Else
                Cells(enterCellRow, 10).Interior.ColorIndex = 3
            End If
            
            perChange = Format((yr_change / strt_pr), "percent")
            Cells(enterCellRow, 11) = perChange
            
        End If

        
        
        strt_price = Cells(i + 1, 3)
        enterCellRow = enterCellRow + 1
        
    Else
        tot_volume = tot_volume + Cells(i, 7).Value
        
    End If
    
    Next i
   

Cells(5, 14) = "Greatest % Increase"
Cells(8, 14) = "Greatest % Decrease"
Cells(11, 14) = "Greatest Total Volume"
Cells(2, 16) = "Value"
Cells(2, 15) = "Ticker"



' take the max and min and place them in a separate part in the worksheet
Cells(5, 16) = "%" & WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100
Cells(8, 16) = "%" & WorksheetFunction.Min(Range("K2:K" & RowCount)) * 100
Cells(11, 16) = WorksheetFunction.Max(Range("L2:L" & RowCount))


' returns one less because header row not a factor
incrNo = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
dcrNo = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
avgNo = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)


' final ticker symbol for  total, greatest % of increase and decrease, and average
Cells(5, 15) = Cells(incrNo + 1, 1)
Cells(8, 15) = Cells(dcrNo + 1, 1)
Cells(11, 15) = Cells(avgNo + 1, 1)

End Sub

Sub LoopWorksheet():

    Dim sht As Worksheet
    
    For Each sht In Worksheets
        sht.Select
        Call stock_analysis
        
        MsgBox sht.Name
        
    Next
    
End Sub








