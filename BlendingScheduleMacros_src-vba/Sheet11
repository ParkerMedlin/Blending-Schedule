Private Sub Worksheet_SelectionChange(ByVal target As Range)

'A:A Triggers Lot Number Generator lookup macro
If InStr(1, (Range("AA2").Value), "On") > 0 Then
    If Selection.Count = 1 Then
        If Not Intersect(target, Range("$A:$A")) Is Nothing Then
            Call lookup_lotNum
        End If
    End If
End If
    
'B:B Triggers blendData lookup macro
If InStr(1, (Range("AA3").Value), "On") > 0 Then
    If Selection.Count = 1 Then
        If Not Intersect(target, Range("$B:$B")) Is Nothing Then
            Call lookup_runs
        End If
    End If
End If
    
'C:C Triggers count history lookup macro
If InStr(1, (Range("AA4").Value), "On") > 0 Then
    If Selection.Count = 1 Then
        If Not Intersect(target, Range("$C:$C")) Is Nothing Then
            Call lookup_countHistory
        End If
    End If
End If

'D:D Triggers BI_BR transaction lookup
If InStr(1, (Range("AA4").Value), "On") > 0 Then
    If Selection.Count = 1 Then
        If Not Intersect(target, Range("$D:$D")) Is Nothing Then
            Call lookup_BI_BR_transactions
        End If
    End If
End If
   
'--------------------------------------------------------------------

'R1 triggers MacrosOn
If Selection.Count = 1 Then
    If Not Intersect(target, Range("$S$1")) Is Nothing Then
        Call macrosOn
    End If
'S1 triggers Macros Off
    If Not Intersect(target, Range("$T$1")) Is Nothing Then
        Call macrosOff
    End If
End If

End Sub



Private Sub Worksheet_Change(ByVal target As Range)

'Timestamper for when I enter a count for a blend
    If InStr(1, (Range("AA1").Value), "On") > 0 Then
        If Not Intersect(target, Range("$K:$K")) Is Nothing Then
            If target.Row > 1 Then
               target.Offset(0, 1) = Now()
            End If
        End If
    End If

End Sub
