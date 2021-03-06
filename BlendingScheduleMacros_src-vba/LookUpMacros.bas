Sub lookup_lotNum()
'///Assigned to column A of Blend These, CheckOutCounts, and ChemShortages sheets (symbol: hashtag)////////////////////////
'///Opens the lot number generator workbook and filters the table by the blend PN of the row you click on//////////////////

'Get Blend PN
    ActiveCell.Offset(0, 4).Select
    Dim blendPN As String
    blendPN = ActiveCell.Value

'Get filepath of Blenging Lont Ntumbner Grenerartror dot xlsx
    Dim LotNumGenPath As String
    LotNumGenPath = "C:\OD\Kinpak, Inc\Blending - Documents\01 Spreadsheet Tools\Blending Lot Number Generator\LotNumGenerator-Prod\Blending Lot Number Generator"
    Workbooks.Open LotNumGenPath

'Filter the list
    ActiveSheet.Range("$A:$A").AutoFilter Field:=1, Criteria1:=blendPN _
        , Operator:=xlAnd

End Sub

Sub lookup_runs()
'///Assigned to column B of Blend These, CheckOutCounts, and ChemShortages sheets (symbol: magnifying glass)///////////////
'///Selects the blendData sheet and filters the table there by the blend PN of the row you click on////////////////////////

    Dim filter As String
    filter = ActiveCell.Offset(0, 3).Value
    With ThisWorkbook
       Sheets("blendData").Visible = True
       Sheets("blendData").Select
       ActiveSheet.ListObjects("blendData").Range.AutoFilter Field:=2, Criteria1:= _
            filter
    End With
        
End Sub

Sub lookup_countHistory()
'///Assigned to column C of Blend These, CheckOutCounts, and ChemShortages sheets (symbol: bunch of grapes)////////////////
'///Selects the CountLog sheet and filters the table by the blend PN of the row you click on///////////////////////////////

    Dim filter As String
    filter = ActiveCell.Offset(0, 2).Value
    Sheets("CountLog").Select
    ActiveSheet.ListObjects("CountLog").Range.AutoFilter Field:=5, Criteria1:= _
        filter

End Sub

Sub lookup_BI_BR_transactions()
'///Assigned to column D of Blend These, CheckOutCounts, and ChemShortages sheets (symbol: up/down arrows)/////////////////
'///Selects the BI_BR transaction sheet and filters the table by the blend PN of the row you selected//////////////////////

    Dim filter As String
    filter = ActiveCell.Offset(0, 1).Value
    With ThisWorkbook
       Sheets("BI_BR_Hist").Visible = True
       Sheets("BI_BR_Hist").Select
       ActiveSheet.ListObjects("BI_BR_Hist_SQLquery").Range.AutoFilter Field:=1, Criteria1:= _
            filter
    End With
    

End Sub