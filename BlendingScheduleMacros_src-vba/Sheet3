Private Sub Worksheet_SelectionChange(ByVal target As Range)

If Selection.Count = 1 Then

'Clear filter and return to CheckOutCounts sheet
    If Not Intersect(target, Range("$L$1")) Is Nothing Then
        Call clearAndReturn_CheckOutCounts
    End If
'Clear filter and return to BlendThese sheet
    If Not Intersect(target, Range("$M$1")) Is Nothing Then
        Call clearAndReturn_BlendThese
    End If
'Clear filter and return to IssueSheetTable sheet
    If Not Intersect(target, Range("$N$1")) Is Nothing Then
        Call clearAndReturn_IssueSheetTable
    End If


End If

End Sub