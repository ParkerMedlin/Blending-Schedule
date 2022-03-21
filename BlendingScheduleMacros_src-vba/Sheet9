Private Sub Worksheet_Change(ByVal target As Range)

'Timestamper for when I enter a count for a blend
    If InStr(1, (Range("AA1").Value), "On") > 0 Then
        If Not Intersect(target, Range("$C$3")) Is Nothing Then
            If target.Row > 1 Then
               Call HistoryReport
            End If
        End If
    End If

End Sub