
Sub potatoPrinterCheckOut()
'///Assigned to printer icon on CheckOutCounts sheet///////////////////////////////////////////////////////////////////////
'///Formats the sheet and then prints it for distribution to inventory crew////////////////////////////////////////////////

'hide rows
    Range("A:D,G:I,L:L,N:AA").Select
    Selection.EntireColumn.Hidden = True
    
'turn off macros
    Call macrosOff

'select visible data body
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.SpecialCells(xlCellTypeVisible).Select
    ActiveSheet.ListObjects("CheckOutCounts_query").DataBodyRange.Select
    
'set row height
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.RowHeight = 32
    
'Insert row, write in reason for counting
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Reason for counting: if our info is wrong, we will run short on blends"
    With Selection.Font
        .Name = "Calibri"
        .Size = 24
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

'Insert row, write in Count and Now()
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Count for " & Now() & " - In Order of Production Needs"
    With Selection.Font
        .Name = "Calibri"
        .Size = 24
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
'print selection
    copyNumber = InputBox("Number of Printed Copies")
    If Not copyNumber = "" Then
        ActiveWindow.SelectedSheets.PrintOut Copies:=copyNumber
    End If
    
'delete the inserted rows
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
     
'show rows again and change row height back
    Range("A:D,G:I,L:L,N:AA").Select
    Selection.EntireColumn.Hidden = False
    ActiveSheet.ListObjects("CheckOutCounts_query").Range.RowHeight = 21
    Range("A1").Select
    
'turn on macros
    Call macrosOn
    
End Sub


Sub printBlendThese()
'///Assigned to printer on BlendThese//////////////////////////////////////////////////////////////////////////////////////
'///Formats the sheet and then prints it///////////////////////////////////////////////////////////////////////////////////

'hide columns
    Range("A:D,L:P,T:AA").Select
    Selection.EntireColumn.Hidden = True
    
'off with the Macros
    Call macrosOff

'input box
    copyNumber = InputBox("Number of Printed Copies")
    If Not copyNumber = "" Then
        
'insert row
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Printed " & Now()
        With Selection.Font
            .Name = "Calibri"
            .Size = 24
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
    
'print selection
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True
        ActiveSheet.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .CenterHorizontally = True
            .CenterVertically = False
            .Orientation = xlPortrait
            .Draft = False
            .PaperSize = xlPaperLetter
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 100
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        Application.PrintCommunication = True
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True
        ActiveSheet.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .CenterHorizontally = True
            .CenterVertically = False
            .Orientation = xlLandscape
            .Draft = False
            .PaperSize = xlPaperLetter
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 150
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        Application.PrintCommunication = True
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
        
'delete row
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
        
    End If
   
   
'show columns again
    Range("A:D,L:P,T:AA").Select
    Selection.EntireColumn.Hidden = False
    Range("E1").Select

'turn on macros
    Call macrosOn
    
End Sub


Sub printIssueSchedule()
'///Assigned to printer on IssueSheetTable/////////////////////////////////////////////////////////////////////////////////
'///Formats the sheet and then prints it///////////////////////////////////////////////////////////////////////////////////
    
'declare int
    Dim i As Integer
    Dim numRows As Integer
    
'find numRows
 '   ActiveSheet.ListObjects("issueSheetTable").DataBodyRange.SpecialCells(xlCellTypeVisible).Select
 '   numRows = Selection.Rows.Count + 4
    NextFree = Range("D2:D" & Rows.Count).Cells.SpecialCells(xlCellTypeBlanks).Row
    Range("D" & NextFree).Select
    numRows = Selection.Row + 1

'insert row
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    Dim theDate As Date
    theDate = InputBox("Enter the date as MM/DD/YYYY")
    ActiveCell.FormulaR1C1 = "Runs for " & theDate
    With Selection.Font
        .Name = "Calibri"
        .Size = 22
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

'hide columns and select appropriate thing
    Range("H:AB").Select
    Selection.EntireColumn.Hidden = True
    Range("A1:G" & numRows).Select

    Dim numberOfCopies As Integer
    numberOfCopies = Application.InputBox(Prompt:="How many copies?", Type:=1)

'Print it however many times
    For i = 1 To numberOfCopies
    Selection.PrintOut
    Next i

'delete row
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

'Unhide things
    Range("H:I").Select
    Selection.EntireColumn.Hidden = False

    
End Sub

Sub printIssueSheets()
'///Assigned to the print button on the PrintButton worksheet of each line's workbook//////////////////////////////////////
'///Delete all unused sheets from the workbook/////////////////////////////////////////////////////////////////////////////

'Turn off the warnings so that sheets can be deleted without interruption
Application.DisplayAlerts = False

'Start at the beginning
Worksheets(1).Activate

'Loop through all sheets and delete the ones that start with "Blending Issue Sheet"
Dim i As Integer
For i = 1 To (ActiveWorkbook.Worksheets.Count - 1)
ActiveSheet.PrintOut
Sheets(i + 1).Select
Next i

'Turn sheet delete warnings back on
Application.DisplayAlerts = True

End Sub

Sub potatoPrinterChemstoCheck()
'///Assigned to printer icon on chemstocheck sheet/////////////////////////////////////////////////////////////////////////
'///Formats the sheet and then prints it for distribution to inventory crew////////////////////////////////////////////////

'hide rows
    Range("A:B").Select
    Selection.EntireColumn.Hidden = True
    
'select visible data body
    ActiveSheet.ListObjects("bom_ChemsToCheck_query").Range.SpecialCells(xlCellTypeVisible).Select
    ActiveSheet.ListObjects("bom_ChemsToCheck_query").DataBodyRange.Select
    
'set row height
    ActiveSheet.ListObjects("bom_ChemsToCheck_query").Range.RowHeight = 32
    
'remove duplicates
    ActiveSheet.Range("bom_ChemsToCheck_query[#All]").RemoveDuplicates Columns:=3 _
        , Header:=xlYes
    
'Insert row, write in Count and Now()
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Count for " & Now() + 1 & " - In Order of Production Needs"
    With Selection.Font
        .Name = "Calibri"
        .Size = 24
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
'print selection
    copyNumber = InputBox("Number of Printed Copies")
    If Not copyNumber = "" Then
        ActiveWindow.SelectedSheets.PrintOut Copies:=copyNumber
    End If
    
'delete the inserted date row
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
          
'show rows again and change row height back
    Range("A:B").Select
    Selection.EntireColumn.Hidden = False
    ActiveSheet.ListObjects("bom_ChemsToCheck_query").Range.RowHeight = 21
    Range("A1").Select
    
'refresh to restore order
    Range("C2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
   
End Sub

Sub BlendTheseCounts()
'///Assigned to arrow icon on BlendThese sheet/////////////////////////////////////////////////////////////////////////////
'///Formats the sheet and then prints it for distribution to inventory crew////////////////////////////////////////////////

'off with the Macros
    Call macrosOff
    
'Filter for items that have not been checked since last change in inventory
    ActiveSheet.ListObjects("timeTable_BlendThese_query").Range.AutoFilter Field _
        :=9, Criteria1:=RGB(146, 208, 80), Operator:=xlFilterCellColor
    
        ActiveSheet.ListObjects("timeTable_BlendThese_query").Range.RowHeight = 32

'Hide unneccessary columns
    Columns("A:D").Select
    Selection.EntireColumn.Hidden = True
    Columns("H:U").Select
    Selection.EntireColumn.Hidden = True
    Columns("V:V").Select
    Selection.EntireColumn.Hidden = False
    Columns("W:AA").Select
    Selection.EntireColumn.Hidden = True
    
    
    
'set row height

    
'select visible data body
    ActiveSheet.ListObjects("timeTable_BlendThese_query").Range.SpecialCells(xlCellTypeVisible).Select
    ActiveSheet.ListObjects("timeTable_BlendThese_query").DataBodyRange.Select
    
'print selection
    copyNumber = InputBox("Number of Printed Copies")
    If Not copyNumber = "" Then
        ActiveSheet.PageSetup.Orientation = xlPortrait
        ActiveWindow.SelectedSheets.PrintOut Copies:=copyNumber
    End If
    
'clear color filter, reset row height, then unhide the columns
    ActiveSheet.ListObjects("timeTable_BlendThese_query").Range.AutoFilter Field:=9
    ActiveSheet.ListObjects("timeTable_BlendThese_query").Range.RowHeight = 15
    Columns("A:D").Select
    Selection.EntireColumn.Hidden = False
    Columns("H:AA").Select
    Selection.EntireColumn.Hidden = False
    Columns("V:V").Select
    Selection.EntireColumn.Hidden = True
    Columns("T:T").Select
    Selection.EntireColumn.Hidden = True
    
End Sub
