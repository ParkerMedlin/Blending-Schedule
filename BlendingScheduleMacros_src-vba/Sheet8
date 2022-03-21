Private Sub Worksheet_SelectionChange(ByVal target As Range)

'Assigns the blend sheet gen macros to the shortage cells
    If InStr(1, (Range("AA1").Value), "On") > 0 Then
        If Selection.Count = 1 Then
            If Not Intersect(target, Range("$Q:$Q")) Is Nothing Then
                Call blndShtGen1wk
                End If
            If Not Intersect(target, Range("$R:R")) Is Nothing Then
                Call blndShtGen2wk
                End If
            If Not Intersect(target, Range("$S:S")) Is Nothing Then
                Call blndShtGen3wk
            End If
        End If
     End If
'--------------------------------------------------------------------

'A:A triggers lot Number lookup macro
    If InStr(1, (Range("AA1").Value), "On") > 0 Then
        If Selection.Count = 1 Then
            If Not Intersect(target, Range("$A:$A")) Is Nothing Then
                Call lookup_lotNum
            End If
        End If
    End If
    
'B:B triggers blend data future run lookup macro
    If InStr(1, (Range("AA2").Value), "On") > 0 Then
        If Selection.Count = 1 Then
            If Not Intersect(target, Range("$B:$B")) Is Nothing Then
                Call lookup_runs
                End If
        End If
    End If

'C:C triggers count history lookup macro
    If InStr(1, (Range("AA4").Value), "On") > 0 Then
        If Selection.Count = 1 Then
            If Not Intersect(target, Range("$C:$C")) Is Nothing Then
                Call lookup_countHistory
            End If
        End If
    End If
    
'D:D triggers BI_BR transaction lookup macro
    If InStr(1, (Range("AA4").Value), "On") > 0 Then
        If Selection.Count = 1 Then
            If Not Intersect(target, Range("$D:$D")) Is Nothing Then
                Call lookup_BI_BR_transactions
            End If
        End If
    End If
   
'U:U Calls Lot Number Generator macro
    If InStr(1, (Range("AA4").Value), "On") > 0 Then
        If Selection.Count = 1 Then
            If Not Intersect(target, Range("$U:$U")) Is Nothing Then
                Call LotNumGen
            End If
        End If
    End If
   
'--------------------------------------------------------------------

'Macros On
If Selection.Count = 1 Then
    If Not Intersect(target, Range("$X$1")) Is Nothing Then
        Call macrosOn
    End If
'Macros Off
    If Not Intersect(target, Range("$Y$1")) Is Nothing Then
        Call macrosOff
    End If
End If
    
End Sub