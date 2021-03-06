Sub refreshAndFilter()
'///Assigned to refresh icon on Blend These sheet//////////////////////////////////////////////////////////////////////////
'///Refreshes everything in the workbook and then auto sizes the columns on CheckOutCounts and Blend These/////////////////

    Dim uInput As Integer
    uInput = MsgBox("Proceed with Refresh?", vbQuestion + vbYesNo)
    
    If uInput = vbYes Then

'Deactivate Blend These selection-based Sheet Code functions
        Call macrosOff
    
'Deactivate CheckOutCounts date logger function
        Sheets("CheckOutCounts").Select
        Call macrosOff
    
'Refresh All
        ActiveWorkbook.RefreshAll
    
'Auto size CheckOutCounts columns
        Columns("E:U").Select
        Columns("E:U").EntireColumn.AutoFit
    
'Reactivate CheckOutCounts date logger function
        Call macrosOn
        
'Reactivate Blend These selection-based sheet code functions
        Sheets("BlendThese").Select
        Columns("P").ColumnWidth = 3
        Call macrosOn
   
        
'Re-filter chemShortages
        Sheets("ChemShortages").Select
        ActiveSheet.ListObjects("bom_Chem_query").Range.AutoFilter Field:=6
        ActiveSheet.ListObjects("bom_Chem_query").Range.AutoFilter Field:=6, _
        Criteria1:="<0", Operator:=xlAnd
        
     Else
        MsgBox "Refresh cancelled."

    End If

End Sub


Sub refreshOH()
'///Assigned to the Abacus shape on CheckOutCounts sheet//////////////////////////////////////////////////////////////////
'///refreshes the qty we have on hand as well as the runs listed on CheckOutCounts/////////////////////////////////////////

    Dim uInput2 As Integer
    uInput2 = MsgBox("Proceed with OnHand Refresh?", vbQuestion + vbYesNo)
    
    If uInput2 = vbYes Then
    
'turn off the macros
    Sheets("CheckOutCounts").Select
    Call macrosOff

'unhide blendqtyonhand and refresh and then rehide it
    With ThisWorkbook
       Sheets("im.blendQty.onHand").Visible = True
       Sheets("im.blendQty.onHand").Select
    End With
    Range("B2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("im.blendQty.onHand").Visible = False
     

'Refresh CheckOutCounts
    Sheets("CheckOutCounts").Select
    Range("B2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
        
'Autofit columns
    Columns("A:L").Select
    Columns("A:L").EntireColumn.AutoFit
    
'Re-enable date logger
    Call macrosOn

    Else
        MsgBox "Refresh cancelled."
    End If
    
End Sub


Sub refreshIssueSheet()
'///Used in the IssueSheetLooper macro to restore the rows removed when filtering duplicates///////////////////////////////
'///refreshes then re-hides the appropriate columns////////////////////////////////////////////////////////////////////////

    Range("A2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

    Range("J:O").EntireColumn.Hidden = True

End Sub


Sub refreshNFilter2()
'///Assigned to arrow on the IssueSheetTable sheet/////////////////////////////////////////////////////////////////////////
'///Refreshes the schedule and then closes/reopens so that IssueSheetTable will actually display correctly/////////////////

    Call refreshAndFilter
    
'Open the blend issue sheet and refresh it
    ChDir "G:\Blending\02 Blending Issue Sheet"
    Workbooks.Open Filename:= _
        "G:\Blending\02 Blending Issue Sheet\NOT the Blending Issue Sheet.xlsb"
    ActiveWorkbook.RefreshAll
 

End Sub
