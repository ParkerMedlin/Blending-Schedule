
Sub clearAndReturn_CheckOutCounts()
'///Assigned to the cell labeled CheckOutCounts////////////////////////////////////////////////////////////////////////////
'///clears all filters from blendData and CountLog table///////////////////////////////////////////////////////////////////
    
    Range("B2").Select
    Sheets("BI_BR_Hist").Select
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1
    Sheets("blendData").Visible = True
    Sheets("blendData").Select
    ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2
    Sheets("blendData").Visible = False
    Sheets("BlendThese").Select
    Sheets("CountLog").Select
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5
    Sheets("CheckOutCounts").Select
    Range("E2").Select
    
End Sub

Sub clearAndReturn_AllCounts()
'///Assigned to the arrows labeled CheckOutCounts//////////////////////////////////////////////////////////////////////////
'///clears all filters from blendData and CountLog table///////////////////////////////////////////////////////////////////
    
    Range("B2").Select
    Sheets("BI_BR_Hist").Select
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1
    Sheets("blendData").Visible = True
    Sheets("blendData").Select
    ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2
    Sheets("blendData").Visible = False
    Sheets("BlendThese").Select
    Sheets("CountLog").Select
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5
    Sheets("AllCounts").Select
    Range("E2").Select
    
End Sub

Sub clearAndReturn_BlendThese()
'///Assigned to the arrows labeled CheckOutCounts//////////////////////////////////////////////////////////////////////////
'///clears all filters from blendData and CountLog table///////////////////////////////////////////////////////////////////
    
    Range("B2").Select
    Sheets("BI_BR_Hist").Select
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1
    Sheets("blendData").Visible = True
    Sheets("blendData").Select
    ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2
    Sheets("blendData").Visible = False
    Sheets("BlendThese").Select
    Sheets("CountLog").Select
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5
    Sheets("BlendThese").Select
    Range("E2").Select
    
End Sub

Sub clearAndReturn_IssueSheetTable()
'///Assigned to the arrows labeled CheckOutCounts//////////////////////////////////////////////////////////////////////////
'///clears all filters from blendData and CountLog table///////////////////////////////////////////////////////////////////
    
    Range("B2").Select
    Sheets("BI_BR_Hist").Select
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1
    Sheets("blendData").Visible = True
    Sheets("blendData").Select
    ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2
    Sheets("blendData").Visible = False
    Sheets("BlendThese").Select
    Sheets("CountLog").Select
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5
    Sheets("IssueSheetTable").Select
    Range("E2").Select
    
End Sub

Sub clearTransactionFilters()
'///Assigned to the shuttlecock on BI_BR_Hist sheet////////////////////////////////////////////////////////////////////////
'///why the shuttlecock, you ask?//////////////////////////////////////////////////////////////////////////////////////////
'///because.///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///clears filters from the TransactionCode and Desc columns, for when you're done investigating blend transactions////////

    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=2, Criteria1:="<>not a blend", _
        Operator:=xlAnd
    ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=5, _
        Criteria1:=Array("BI", "BR", "II"), Operator:=xlFilterValues

End Sub

Sub clearChemsToCheckFilter()
'///Assigned to cell N1////////////////////////////////////////////////////////////////////////////////////////////////////
'///Clears the filter only for [BlendPN], also sets M1 to blank////////////////////////////////////////////////////////////

   ActiveSheet.ListObjects("bom_ChemsToCheck_query").Range.AutoFilter Field:=1
   Rows.EntireRow("1:1").Delete
   Range("M1").Value = ""
 
End Sub

Sub clearCheckOutCountsFilter()
'///Assigned to cell the x button//////////////////////////////////////////////////////////////////////////////////////////
'///Clears the filter only for [BlendPN], also sets M1 to blank////////////////////////////////////////////////////////////

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

End Sub

Sub clearitAllHistoryReport()
'///Assigned to the x icon on the CLEARSHEETS worksheet in the HistoryReport workbook//////////////////////////////////////
'///Deletes the first three worksheets in the document/////////////////////////////////////////////////////////////////////

'Turn off the warnings so that sheets can be deleted without interruption
Application.DisplayAlerts = False

'Start at the beginning
Worksheets(1).Activate

'Loop through all sheets and delete the ones that start with "Blending Issue Sheet"
Dim i As Integer
For i = 1 To (ActiveWorkbook.Worksheets.Count - 1)
ActiveSheet.Delete
Next i

'Turn sheet delete warnings back on
Application.DisplayAlerts = True

Range("A2").Select

End Sub
