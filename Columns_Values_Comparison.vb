
Sub Columns_Values_Comparison()
'

' Comparison on values on columns A & B returning on C & D repectively values not found

Dim rngCell As Range
For Each rngCell In Range("A2:A1000")
    If WorksheetFunction.CountIf(Range("B2:B1000"), rngCell) = 0 Then
        Range("C" & Rows.Count).End(xlUp).Offset(1) = rngCell
    End If
Next
For Each rngCell In Range("B2:B1000")
    If WorksheetFunction.CountIf(Range("A2:A1000"), rngCell) = 0 Then
        Range("D" & Rows.Count).End(xlUp).Offset(1) = rngCell
    End If
Next


'
End Sub
