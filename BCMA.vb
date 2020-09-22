Sub BCMA()

'Last Update: 4/17/2020

'Script to clean & prep. data from SSRS to excel dataset feeding Tableau Report'

'_GLOBAL VARIABLES_
'------------------
Facility1_arr = Array("HOLTZ - CHILDREN", "HOLTZ - WOMEN", "JMH", "JNMC", "JSCH", "MENTAL HEALTH", "REHAB")

lRow = ActiveCell.SpecialCells(xlLastCell).Row


'_REMOVING ROWS_
'---------------------
Rows("1:7").Select
Selection.Delete Shift:=xlUp

For iCntr = lRow To 3 Step -1
    If IsInArray(Cells(iCntr, 4), Facility1_arr) Then
    Else
        Rows(iCntr).Delete
    End If
Next

Rows("2:2").Select
Selection.Delete Shift:=xlUp

'_REMOVING SOME COLUMNS_
'---------------------
Columns("A:B").Select
Selection.Delete Shift:=xlToLeft

Columns("D:F").Select
Selection.Delete Shift:=xlToLeft

Range("D:E").Select
Selection.UnMerge
Columns("E:E").Select
Selection.Delete Shift:=xlToLeft

Range("I:J").Select
Selection.UnMerge
Columns("J:J").Select
Selection.Delete Shift:=xlToLeft

Range("C:C").Select
Selection.ClearContents
Range("C1").Value2 = "Month"
'
End Sub

