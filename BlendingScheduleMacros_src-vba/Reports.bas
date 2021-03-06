Sub HistoryReport()
'///Triggers when you type in a part number on the HistoryReport page//////////////////////////////////////////////////////
'///Gives you a report of all BI and BR transactions as well as all of our counts//////////////////////////////////////////

'Navigation variables for the workbook names and the blend PN
    Dim blendPN As String
    blendPN = ActiveCell.Offset(-1, 0).Value
    Dim src As String
    src = ActiveWorkbook.Name
    Dim reportWB As String
    reportWB = "C:\OD\Kinpak, Inc\Blending - Documents\03 Projects\ReportGen-Destination\HistoryReport.xlsb"
   
'Open History Report and clear the unnecessary sheets
    Workbooks.Open (reportWB)
    reportWB = ActiveWorkbook.Name
    Call clearitAllHistoryReport
    Windows(src).Activate
    
'BI_BR transactions:
    With ThisWorkbook
        Sheets("BI_BR_Hist").Visible = True
    End With
    Sheets("BI_BR_Hist").Activate
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1, Criteria1:= _
        blendPN
    Range("A1:F1").Select
    Selection.Copy
    Windows(reportWB).Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows(src).Activate
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Windows(reportWB).Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Name = "transactHist"
    Columns(3).Delete
    Columns(4).Insert
    Columns(4).Insert
    Range("C:C").NumberFormat = "mm/dd/yyyy"
    Worksheets("transactHist").Columns("A:F").AutoFit
    Sheets.Add
    Windows(src).Activate
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1

'countLog:
    With ThisWorkbook
        Sheets("CountLog").Visible = True
        Sheets("CountLog").Activate
    End With
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5, Criteria1:= _
        blendPN
    Range("E1:O1").Select
    Selection.Copy
    Windows(reportWB).Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False
    Windows(src).Activate
    ActiveSheet.ListObjects("countLog").DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Windows(reportWB).Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Name = "countHist"
    Columns(14).EntireColumn.Delete
    Columns(13).EntireColumn.Delete
    Columns(12).EntireColumn.Delete
    Columns(11).EntireColumn.Delete
    Columns(10).EntireColumn.Delete
    Columns(9).EntireColumn.Delete
    Columns(5).EntireColumn.Delete
    Columns(4).EntireColumn.Delete
    Columns(3).EntireColumn.Delete
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("C:C").NumberFormat = "mm/dd/yyyy"
    Worksheets("countHist").Columns("A:I").AutoFit
    Windows(src).Activate
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5
    Windows(reportWB).Activate
   
'Composite report ordered by date
    Sheets.Add
    ActiveSheet.Name = "timeline"
    Range("A1").Value = "Blend PN"
    Range("B1").Value = "Description"
    Range("C1").Value = "Date"
    Range("D1").Value = "Exp OH"
    Range("E1").Value = "Count"
    Range("F1").Value = "TransacType"
    Range("G1").Value = "TransacQty"
    Sheets("countHist").Activate
    Range("A2:F200").Select
    Selection.Copy
    Sheets("timeline").Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
           :=False, Transpose:=False
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    Sheets("transactHist").Activate
    Range("A2:G200").Select
    Selection.Copy
    Sheets("timeline").Activate
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
           :=False, Transpose:=False
    Range("A2:G200").Select
    ActiveWorkbook.Worksheets("timeline").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("timeline").Sort.SortFields.Add2 Key:=Range( _
        "C2:C111"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("timeline").Sort
        .SetRange Range("A1:G111")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:G").AutoFit
    Range("C:C").NumberFormat = "mm/dd/yyyy"
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$G$200"), , xlYes).Name = _
        "timelineTable"
  
End Sub


Sub todaysCounts()
'///Assigned to the Sheet icon on CheckOutCounts page//////////////////////////////////////////////////////////////////////
'///Gives you a report of counts that are still on the main checkoutcounts page////////////////////////////////////////////

'Navigation variables for the workbook names
    Dim src As String
    src = ActiveWorkbook.Name
    Dim reportWB As String
    reportWB = "C:\OD\Kinpak, Inc\Blending - Documents\03 Projects\ReportGen-Destination\DailyCountReport.xlsm"
    
'Turn macros off
    Call macrosOff

'Open workbook and clear info from last time
    Workbooks.Open (reportWB)
    reportWB = ActiveWorkbook.Name
    Columns("A:E").EntireColumn.Delete
    
'Go back, get the blanks out, and copy the list
    Windows(src).Activate
    ActiveSheet.ListObjects("CheckOutCounts_query").DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    
'Paste and insert row for headers
    Windows(reportWB).Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

'Headers copypasta
    Windows(src).Activate
    Range("A1:R1").Select
    Selection.Copy
    Windows(reportWB).Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Clean up unnecessary columns, right to left. then Autofit
    Columns("P:R").EntireColumn.Delete
    Columns("M:N").EntireColumn.Delete
    Columns("G:I").EntireColumn.Delete
    Columns("A:E").EntireColumn.Delete
    Columns("A:E").AutoFit
    

'Select the active stuff and format
    Dim addy As String
    Range("A1").End(xlDown).Select
    addy = ActiveCell.Address
    Range(addy & ":E1").Select
    
'Gross formatting borders ew this is so very chunk
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    Range("A1:E1").Select
    Selection.Font.Bold = True
    
'Delete blanks and format date column
    Columns("C:C").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("D:D").Select
    Selection.NumberFormat = "m/d/yyyy"
   
'Back to other workbook, macrosOn, then end on destination workbook
    Windows(src).Activate
    Call macrosOn
    Windows(reportWB).Activate


End Sub


Sub startronReport()
'///Assigned to the startron logo on the BlendThese sheet//////////////////////////////////////////////////////////////////
'///Selects the blendData sheet and filters the table there by the blend PN of the row you click on////////////////////////

'Navigation variables for the workbook names
    Dim src As String
    src = ActiveWorkbook.Name
    Dim reportWB As String
    reportWB = "C:\OD\Kinpak, Inc\Blending - Documents\03 Projects\ReportGen-Destination\StartronReport.xlsm"

'Open workbook and clear info from last time
    Workbooks.Open (reportWB)
    reportWB = ActiveWorkbook.Name
    Columns("A:I").EntireColumn.Delete

'Unhide and select BlendData sheet
    Windows(src).Activate
    Sheets("blendData").Visible = True
    Sheets("blendData").Select
   
'filter the list by all the different startron blend PNs and copy the table to the report workbook each time
        ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2, Criteria1:= _
            Array("14308.B", "14308AMBER.B", "93100DSL.B", "93100GAS.B", "93100TANK.B", "93100GASBLUE.B", "93100GASAMBER.B"), _
            Operator:=xlFilterValues
        Range("A1:I2200").Copy
        Windows(reportWB).Activate
        With ThisWorkbook
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        Windows(src).Activate

'Bring over the headers
    Windows(src).Activate
    ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2
    Range("A1:I1").Copy
    Sheets("BlendThese").Select
    Windows(reportWB).Activate
    With ThisWorkbook
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End With
    
'Make a table
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$I$200"), , xlYes).Name = _
        "startronTable"
    
'loop through the first 100 rows and delete ones that don't have a count in them
    Dim inc As Integer
    Dim rowNum As Integer
    rowNum = 3

    For inc = 3 To 200
        If IsEmpty(Range("B" & rowNum).Value) Then
             Rows(rowNum).EntireRow.Delete
        Else
            rowNum = rowNum + 1
        End If
    Next inc
    
'Autofit columns
    Columns("A:G").AutoFit

'Sort by StartTime
    ActiveWorkbook.Worksheets("Report").ListObjects("startronTable").Sort. _
        SortFields.Add2 Key:=Range("startronTable[[#All],[StartTime]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Report").ListObjects("startronTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub
