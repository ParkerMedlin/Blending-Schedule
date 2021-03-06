Sub LogCheckOutCounts()
'///Assigned to the blue arrow on the CheckOutCounts sheet/////////////////////////////////////////////////////////////////
'///Copies active data body from CheckOutCounts and inserts copied cells underneath headers on CHECKOUTdiff.log sheet//////

'Turn off date logger to prevent fuckery
    Call macrosOff

'Clear all the filters
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=5
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=6
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=7
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=8
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=9
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=10
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=11
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=12
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=13
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=14
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=15
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=16
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=17
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.AutoFilter Field:=18

'Select only the visible cells in a filtered table
    ActiveSheet.ListObjects("CheckOutCounts_query").DataBodyRange.SpecialCells(xlCellTypeVisible).Select

'Get the rowcount and insert correct num of rows
    Dim RowCount As Integer
    RowCount = Selection.Rows.Count
    Sheets("CountLog").Select
    Range("A2").EntireRow.Offset(1).Resize(RowCount).Insert Shift:=xlDown
    Sheets("CheckOutCounts").Select

'Copy and paste active data body
    ActiveSheet.ListObjects("CheckOutCounts_query").DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets("CountLog").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'loop through the first 100 rows and delete ones that don't have a count in them
    Dim inc As Integer
    Dim rowNum As Integer
    rowNum = 3
    
    For inc = 3 To 100
        If IsEmpty(Range("K" & rowNum).Value) Then
             Rows(rowNum).EntireRow.Delete
        Else
            rowNum = rowNum + 1
        End If
    Next inc
    
'bottom border of pasted selection
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

'Timestamp on CHECKOUTdiff.log
    Range("S3").Value = Now()

'Turn Macros back on
    Sheets("CheckOutCounts").Select
    Call macrosOn

End Sub
Sub LogAllCounts()
'///Assigned to the blue arrow on the CheckOutCounts sheet/////////////////////////////////////////////////////////////////
'///Copies active data body from CheckOutCounts and inserts copied cells underneath headers on CHECKOUTdiff.log sheet//////

'Turn off date logger to prevent fuckery
    Call macrosOff

'Select only the visible cells in a filtered table
    ActiveSheet.ListObjects("AllCounts_query").DataBodyRange.SpecialCells(xlCellTypeVisible).Select

'Get the rowcount and insert correct num of rows
    Dim RowCount As Integer
    RowCount = Selection.Rows.Count
    Sheets("AllCounts").Select
    Range("A2").EntireRow.Offset(1).Resize(RowCount).Insert Shift:=xlDown
    Sheets("MasterCounts").Select

'Copy and paste active data body
    ActiveSheet.ListObjects("AllCounts_query").DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets("CountLog").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'bottom border of pasted selection
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

'Timestamp on CountsLog
    Range("P3").Value = Now()

'Top border for timestamp cell
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

'Turn date logger back on
    Sheets("AllCounts").Select
    Call macrosOn

End Sub
