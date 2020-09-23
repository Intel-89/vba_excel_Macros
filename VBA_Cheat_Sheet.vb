'_Creating a macro timer_
'------------------------
Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer 'Remember time when macro starts
'*** Macro***'
SecondsElapsed = Round(Timer - StartTime, 2) 'Determine how many seconds code took to run
MsgBox ("Macro executed successfully in " & SecondsElapsed & " seconds") '--'


'_Easy Reference Method_
'------------------------
[A1].Value="Certain Value"

'_Disabling worksheet recalculation, screen updating, statusbar updating_ Method 1_
'----------------------------------------------------------------------------------
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'*** Macro***'
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
ActiveWindow.DisplayGridlines = False


'_Disabling worksheet recalculation, screen updating, statusbar updating_ Method 2_
'----------------------------------------------------------------------------------
'Creating a funtion, later on calling it with True, and then at then calling with False
Option Explicit
Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not(isOn)
    Application.ScreenUpdating = Not(isOn)
	Application.DisplayAlerts = Not(isOn)
    ActiveSheet.DisplayPageBreaks = Not(isOn)
End Sub

OptimizeVBA True
'*** Macro***'
ActiveWindow.DisplayGridlines = False
OptimizeVBA False


'_MESSAGE BOX EXAMPLES_
'--------------------------------------
MsgBox ("Macro executed successfully in " & SecondsElapsed & " seconds")

'_Simple inline modification to sheets_
'--------------------------------------
Sheets("License_Plate").Select               '_Selecting sheets knowing name_
Sheets("License_Plate").Name = "Inventory"   '_Renaming sheets knowing name_
Sheets(1).Name = "Inventory"                 '_Renaming sheets knowing position/counter_
Sheets.Add After:=Sheets(Sheets.Count)       '_Adding Sheet after last available sheet_
Sheets("Table 3").Delete                     '_Deleting sheet with "certain" name_
Sheets(1).Delete                             '_Deleting sheet with position/counter_
Rows(iCntr).Delete                           '_Deleting row on "certain" position/counter_
Cell(iCntr1,iCntr2) or Range(iCntr1,iCntr2)  '_Selecting certain cell or range_

'_REMOVING FIRST 2 EMPTY ROWS_
'-----------------------------
' First positing on desire sheet and then
 Rows("1:2").Select
 Selection.Delete Shift:=xlUp
End If


'_Clearing content on cell_
'--------------------------
Range("N1").Select
Selection.ClearContents


'_Deleting some columns_
'-----------------------
Columns("K:K").Select
Selection.Delete Shift:=xlToLeft


'_Copying Range from 1 worksheet to another one_
'-----------------------------------------------
Sheets("Falls and Med Events").Range("1:1").Copy Destination:=Sheets("Holtz Falls").Range("1:1")


'_Defining variables_
'-------------------
lRow = ActiveCell.SpecialCells(xlLastCell).Row ' For last row in this case
Location_Array = Array("a", "b", "c", "d", "e")'arrays initial position is 0 on -VBA-


'_Inserting an textual formula as cell value_
'--------------------------------------------
Cells(x,y).Select
ActiveCell.FormulaR1C1 = "=IF((TIMEVALUE(RC[-9])>TIME(14,29,0)),""2nd Shift"",""1st Shift"")"


'_Finding last cell
'-------------------------------------------------------------------------------------------------------------------------------
'_Finding last cell with value on rage - Method 1_
ActiveCell.SpecialCells(xlLastCell).Select

'_Finding last cell with value on rage - Method 2 (since is using last used cell sometime it can be an empty cell)_
MsgBox Range("A1").SpecialCells(xlCellTypeLastCell).Address

'_Finding last cell with value on rage - Method 3 ('Finds the last non-blank cell on a sheet/range)_
Dim lRow As Long
Dim lCol As Long
    
lRow = Cells.Find(What:="*",After:=Range("A1"),LookAt:=xlPart,LookIn:=xlFormulas,SearchOrder:=xlByRows,SearchDirection:=xlPrevious,MatchCase:=False).Row
'-----------------------------------------------------------------------------------------------------------------------------------


'_Looping from the bottom to delete complete row not matching "certain value" on a "certain cell" of the row_
'------------------------------------------------------------------------------------------------------------
Dim lRow As Long
Dim iCntr As Long

lRow = 5000
For iCntr = lRow To 1 Step -1
    If Cells(iCntr, 2) = "PRIDE OF  AMERICA" Or Cells(iCntr, 2) = "NORWEGIAN JADE" Then .' .Value2 in case of error
        Rows(iCntr).Delete
	'ElseIf - also posible at this point
    End If
Next


'_Creating a counter and increasing with the value inside a cell_
'----------------------------------------------------------------
Dim HVCounterPA1 As Long
Counter = Counter + Val(ActiveCell.Value) 'Cells(x,y).Select before the ActiveCell or Cells(x,y).Value can be used


'_IsInArray() funtion will compare value from cell with array (Boolean return)_
'------------------------------------------------------------------------------
'Creating a funtion, later on calling it with True, and then at then calling with False
Function IsInArray(ByVal VarToBeFound As Variant, ByVal Arr As Variant) As Boolean
    Dim Element As Variant
    For Each Element In Arr
        If Element = VarToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next Element

    IsInArray = False
End Function


'Removing rows where substring appears_
'--------------------------------------
aVariable = "No Person Involved"
For iCntr = lRow To 1 Step -1
    If InStr(Cells(iCntr, 9), aVariable) Then ' InStr() built-in function will look for substring(2nd parameter) inside substring(1st parameter) returning position, in this cases any position will indicate ocurrence
        Rows(iCntr).Delete
    End If
Next

'_Creating new workbooks_
'------------------------
ActiveWorkbook.SaveAs Filename:="Falls_file.xlsx"
Workbooks.Add.SaveAs Filename:="Errors_file.xlsx"

'_Moving sheets from "errors" to 2nd workbook created_
'-----------------------------------------------------
Workbooks("Falls_file.xlsx").Sheets("Holtz Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson Behavioral Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson Memorial Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson North Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson South Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson Rehab. Hospital Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
