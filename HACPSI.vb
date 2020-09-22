Function aSubtotalMonthly(Var As String) As Boolean

'_RESETING VARIABLES FOR EACH INDICATOR(DENOMINATOR & NUMERATOR)
'--------------------------------------------------------------
nJHS_HAC1 = 0
dJHS_HAC1 = 0
nJHS_HAC3 = 0
dJHS_HAC3 = 0
nJHS_HAC5 = 0
dJHS_HAC5 = 0
nJHS_PSI3 = 0
dJHS_PSI3 = 0
nJHS_PSI4 = 0
dJHS_PSI4 = 0
nJHS_PSI6 = 0
dJHS_PSI6 = 0
nJHS_PSI9 = 0
dJHS_PSI9 = 0
nJHS_PSI10 = 0
dJHS_PSI10 = 0
nJHS_PSI11 = 0
dJHS_PSI11 = 0
nJHS_PSI12 = 0
dJHS_PSI12 = 0
nJHS_PSI13 = 0
dJHS_PSI13 = 0
nJHS_PSI14 = 0
dJHS_PSI14 = 0
nJHS_PSI15 = 0
dJHS_PSI15 = 0
'--------------
cJHS_HAC1 = 0
cJHS_HAC3 = 0
cJHS_HAC5 = 0
cJHS_PSI3 = 0
cJHS_PSI4 = 0
cJHS_PSI6 = 0
cJHS_PSI9 = 0
cJHS_PSI10 = 0
cJHS_PSI11 = 0
cJHS_PSI12 = 0
cJHS_PSI13 = 0
cJHS_PSI14 = 0
cJHS_PSI15 = 0
'------------------------


'_Finding last populated cell_
'-----------------------------
lRow = ActiveCell.SpecialCells(xlLastCell).Row

'_Looping for to increment denominator and numerators per month
'--------------------------------------------------------------
'And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")_ was added to remove that combo from month stats
'StrComp() will return False when string are equal, that is why is used below with NOT in front

For iCntr = lRow To 2 Step -1
    If Cells(iCntr, 4) = "HAC1" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then '
        ' meaning: numerator and den. counters for each specific metric on "Jackson Memorial Hospital" and also on "All hospitals"
        nJHS_HAC1 = nJHS_HAC1 + Cells(iCntr, 6) 'Numberator
        dJHS_HAC1 = dJHS_HAC1 + Cells(iCntr, 7) 'Numberator
        cJHS_HAC1 = Cells(iCntr, 9)             'Cohort always same value after getting first one
    ElseIf Cells(iCntr, 4) = "HAC3" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_HAC3 = nJHS_HAC3 + Cells(iCntr, 6)
        dJHS_HAC3 = dJHS_HAC3 + Cells(iCntr, 7)
        cJHS_HAC3 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "HAC5" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_HAC5 = nJHS_HAC5 + Cells(iCntr, 6)
        dJHS_HAC5 = dJHS_HAC5 + Cells(iCntr, 7)
        cJHS_HAC5 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 3" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI3 = nJHS_PSI3 + Cells(iCntr, 6)
        dJHS_PSI3 = dJHS_PSI3 + Cells(iCntr, 7)
        cJHS_PSI3 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 4" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI4 = nJHS_PSI4 + Cells(iCntr, 6)
        dJHS_PSI4 = dJHS_PSI4 + Cells(iCntr, 7)
        cJHS_PSI4 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 6" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI6 = nJHS_PSI6 + Cells(iCntr, 6)
        dJHS_PSI6 = dJHS_PSI6 + Cells(iCntr, 7)
        cJHS_PSI6 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 9" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI9 = nJHS_PSI9 + Cells(iCntr, 6)
        dJHS_PSI9 = dJHS_PSI9 + Cells(iCntr, 7)
        cJHS_PSI9 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 11" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI11 = nJHS_PSI11 + Cells(iCntr, 6)
        dJHS_PSI11 = dJHS_PSI11 + Cells(iCntr, 7)
        cJHS_PSI11 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 13" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI13 = nJHS_PSI13 + Cells(iCntr, 6)
        dJHS_PSI13 = dJHS_PSI13 + Cells(iCntr, 7)
        cJHS_PSI13 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 15" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI15 = nJHS_PSI15 + Cells(iCntr, 6)
        dJHS_PSI15 = dJHS_PSI15 + Cells(iCntr, 7)
        cJHS_PSI15 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 10" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI10 = nJHS_PSI10 + Cells(iCntr, 6)
        dJHS_PSI10 = dJHS_PSI10 + Cells(iCntr, 7)
        cJHS_PSI10 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 12" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI12 = nJHS_PSI12 + Cells(iCntr, 6)
        dJHS_PSI12 = dJHS_PSI12 + Cells(iCntr, 7)
        cJHS_PSI12 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 14" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI14 = nJHS_PSI14 + Cells(iCntr, 6)
        dJHS_PSI14 = dJHS_PSI14 + Cells(iCntr, 7)
        cJHS_PSI14 = Cells(iCntr, 9)
    ElseIf Cells(iCntr, 4) = "PSI 15" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And Not StrComp(Cells(iCntr, 1), "Jackson Memorial Hospital") And Not StrComp(Cells(iCntr, 2), "All hospitals") Then
        nJHS_PSI15 = nJHS_PSI15 + Cells(iCntr, 6)
        dJHS_PSI15 = dJHS_PSI15 + Cells(iCntr, 7)
        cJHS_PSI15 = Cells(iCntr, 9)
    End If
Next

'_Looping for cohort values- change to 3 line on previous look to improve performance
'-----------------------------
'For iCntr = lRow To 2 Step -1
'    If (Cells(iCntr, 4) = "HAC1" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_HAC1 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "HAC3" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_HAC3 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "HAC5" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_HAC5 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 3" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI3 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 4" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI4 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 6" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI6 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 9" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI9 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 11" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI11 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 13" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI13 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 15" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI15 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 10" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI10 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 12" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI12 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 14" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI14 = Cells(iCntr, 9)
'    ElseIf (Cells(iCntr, 4) = "PSI 15" And Left(Cells(iCntr, 3), 3) = Left(Var, 3) And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")) Then
'        cJHS_PSI15 = Cells(iCntr, 9)
'    End If
'Next


'_Finding first empty cell on A column to populate metric calculated above_
'--------------------------------------
aRange = Range("A1").End(xlDown).Offset(1, 0).Select 'Selecting
aRow = ActiveCell.Row                                'Saving row value for future


'_Setting format to text for issue with month, excel was handling incorrectly the pasting of the value on the cell on loop below
Columns(3).Select
Selection.NumberFormat = "@"


'_Looping adding metric values on current month for "Jackson Health System"_
'------------------------------------------------------------------------
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "HAC1"
Cells.Item(aRow, "E") = "Foreign object retained after surgery"
Cells.Item(aRow, "F") = nJHS_HAC1
Cells.Item(aRow, "G") = dJHS_HAC1
Cells.Item(aRow, "I") = cJHS_HAC1
Cells.Item(aRow, "J") = "Dr. Goldberg"
Cells.Item(aRow, "K") = "Group 1"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "HAC3"
Cells.Item(aRow, "E") = "Blood incompatibility"
Cells.Item(aRow, "F") = nJHS_HAC3
Cells.Item(aRow, "G") = dJHS_HAC3
Cells.Item(aRow, "I") = cJHS_HAC3
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "HAC5"
Cells.Item(aRow, "E") = "Falls and trauma"
Cells.Item(aRow, "F") = nJHS_HAC5
Cells.Item(aRow, "G") = dJHS_HAC5
Cells.Item(aRow, "I") = cJHS_HAC5
Cells.Item(aRow, "J") = "Arlene Cameron"
Cells.Item(aRow, "K") = "Group 3"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI3"
Cells.Item(aRow, "E") = "Pressure Ulcer Rate"
Cells.Item(aRow, "F") = nJHS_PSI3
Cells.Item(aRow, "G") = dJHS_PSI3
Cells.Item(aRow, "I") = cJHS_PSI3
Cells.Item(aRow, "J") = "Liz Mac"
Cells.Item(aRow, "K") = "Group 3"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 4"
Cells.Item(aRow, "E") = "Death Rate among Surgical Inpatients with Serious Treatable Conditions"
Cells.Item(aRow, "F") = nJHS_PSI4
Cells.Item(aRow, "G") = dJHS_PSI4
Cells.Item(aRow, "I") = cJHS_PSI4
Cells.Item(aRow, "J") = "Dr. Goldberg"
Cells.Item(aRow, "K") = "Group 1"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 6"
Cells.Item(aRow, "E") = "Iatrogenic Pneumothorax Rate"
Cells.Item(aRow, "F") = nJHS_PSI6
Cells.Item(aRow, "G") = dJHS_PSI6
Cells.Item(aRow, "I") = cJHS_PSI6
Cells.Item(aRow, "J") = "Dr. Silverman"
Cells.Item(aRow, "K") = "Group 2"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 9"
Cells.Item(aRow, "E") = "Perioperative Hemorrhage or Hematoma Rate"
Cells.Item(aRow, "F") = nJHS_PSI9
Cells.Item(aRow, "G") = dJHS_PSI9
Cells.Item(aRow, "I") = cJHS_PSI9
Cells.Item(aRow, "J") = "Dr. Goldberg"
Cells.Item(aRow, "K") = "Group 1"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 10"
Cells.Item(aRow, "E") = "Postoperative Acute Kidney Injury Requiring Dialysis"
Cells.Item(aRow, "F") = nJHS_PSI10
Cells.Item(aRow, "G") = dJHS_PSI10
Cells.Item(aRow, "I") = cJHS_PSI10
Cells.Item(aRow, "J") = "Dr. Silverman"
Cells.Item(aRow, "K") = "Group 2"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 11"
Cells.Item(aRow, "E") = "Postoperative Respiratory Failure Rate"
Cells.Item(aRow, "F") = nJHS_PSI11
Cells.Item(aRow, "G") = dJHS_PSI11
Cells.Item(aRow, "I") = cJHS_PSI11
Cells.Item(aRow, "J") = "Dr. Silverman"
Cells.Item(aRow, "K") = "Group 2"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 12"
Cells.Item(aRow, "E") = "Perioperative Pulmonary Embolism or Deep Vein Thrombosis Rate"
Cells.Item(aRow, "F") = nJHS_PSI12
Cells.Item(aRow, "G") = dJHS_PSI12
Cells.Item(aRow, "I") = cJHS_PSI12
Cells.Item(aRow, "J") = "Dr. Goldberg"
Cells.Item(aRow, "K") = "Group 1"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 13"
Cells.Item(aRow, "E") = "Postoperative Sepsis Rate"
Cells.Item(aRow, "F") = nJHS_PSI13
Cells.Item(aRow, "G") = dJHS_PSI13
Cells.Item(aRow, "I") = cJHS_PSI13
Cells.Item(aRow, "J") = "Dr. Silverman"
Cells.Item(aRow, "K") = "Group 2"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 14"
Cells.Item(aRow, "E") = "Postoperative Wound Dehiscence Rate"
Cells.Item(aRow, "F") = nJHS_PSI14
Cells.Item(aRow, "G") = dJHS_PSI14
Cells.Item(aRow, "I") = cJHS_PSI14
Cells.Item(aRow, "J") = "Dr. Goldberg"
Cells.Item(aRow, "K") = "Group 1"
aRow = aRow + 1
Cells.Item(aRow, "A") = "Jackson Health System"
Cells.Item(aRow, "B") = "All Hospitals"
Cells.Item(aRow, "C") = Var
Cells.Item(aRow, "D") = "PSI 15"
Cells.Item(aRow, "E") = "Unrecognized Abdominopelvic Accidental Puncture/Laceration Rate"
Cells.Item(aRow, "F") = nJHS_PSI15
Cells.Item(aRow, "G") = dJHS_PSI15
Cells.Item(aRow, "I") = cJHS_PSI15
Cells.Item(aRow, "J") = "Dr. Silverman"
Cells.Item(aRow, "K") = "Group 2"
End Function
'_Disabling worksheet recalculation, screen updating, statusbar updating_
Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    Application.DisplayAlerts = Not (isOn)
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub
Sub HACPSI()

'Last Update: 2/13/2020

'Script to clean & prep. data to excel dataset feeding Tableau Report'

Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer 'Remember time when macro starts

'_Visualization OFF for optimization_
'------------------------------------
OptimizeVBA True

'_Inserting 2 Columns on front_
'---------------------
Columns("A:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'_Inserting 1 Row for headers in custom way_
'-----------------------------
Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A1") = "Facility"
Range("B1") = "Cohort Type"
Range("D1") = "Month"
Range("E1") = "Metric"
Range("F1") = "Measure"
Range("G1") = "Metrics"
Range("H1") = "Num"
Range("I1") = "Den"
Range("J1") = "Rate"
Range("K1") = "Cohort"
Range("L1") = "Process Owner"
Range("M1") = "Group"

'_Finding last populated cell_
'-----------------------------
ActiveCell.SpecialCells(xlLastCell).Select
lRow = Selection.Row

'_Logic for data input of added column and header removal_

aFacility = "N/A"
aCohort = "N/A"
aPatType = "N/A"
For iCntr = 2 To lRow
    aCell = Cells(iCntr, 4)
    aOffset = Cells(iCntr, 4).Offset(0, 1)
    aOffsetend = Cells(iCntr, 4).Offset(4, 0)
    
    If (aCell = "Month") Then
        Rows(iCntr).Delete
        iCntr = iCntr - 1
    ElseIf IsEmpty(aCell) And Cells(iCntr, 4).Offset(1, 0) = "Page by:" Then
        Rows(iCntr).Delete
        iCntr = iCntr - 1
    ElseIf IsEmpty(aCell) And IsEmpty(aOffsetend) Then
        Rows(iCntr).Delete
        'iCntr = iCntr - 1
    ElseIf Left(aCell, 10) = "Facility: " Then
        aFacility = Right(aCell, Len(aCell) - 10)
        Rows(iCntr).Delete
        iCntr = iCntr - 1
    ElseIf Left(aCell, 9) = "Benchmark" Then
        aCohort = Right(aCell, Len(aCell) - 24)
        Rows(iCntr).Delete
        iCntr = iCntr - 1
    ElseIf Left(aCell, 7) = "Patient" Then
        aPatType = Right(aCell, Len(aCell) - 36)
        Rows(iCntr).Delete
        iCntr = iCntr - 1
    ElseIf IsEmpty(aOffset) Then
        Rows(iCntr).Delete
        iCntr = iCntr - 1
    Else
        Cells(iCntr, 3) = aFacility
        Cells(iCntr, 2) = aCohort
        Cells(iCntr, 1) = aPatType
    End If
Next

'_Original raw file change at some point some adjustment needed_
'--------------------------------------------------------------
For iCntr = lRow To 2 Step -1
    If (Cells(iCntr, 3) = "Jackson Main Campus") And (Cells(iCntr, 1) = "Jackson South") Then
        Rows(iCntr).Delete
    End If
Next

Columns(3).Delete 'No longer needed
Columns(6).Delete 'No longer needed

'_Adjustment of last row_
'------------------------
ActiveSheet.UsedRange
Range("A1").Select

'_Metric Conditionals to add Group and Owner _
'--------------------------------------------------------------
ActiveCell.SpecialCells(xlLastCell).Select
lRow = Selection.Row

For iCntr = lRow To 2 Step -1
    If (Cells(iCntr, 4) = "HAC1") Then
        Cells(iCntr, 10) = "Dr. Goldberg"
        Cells(iCntr, 11) = "Group 1"
    ElseIf (Cells(iCntr, 4) = "HAC5") Then
        Cells(iCntr, 10) = "Arlene Cameron"
        Cells(iCntr, 11) = "Group 3"
    ElseIf (Cells(iCntr, 4) = "PSI 4") Then
        Cells(iCntr, 10) = "Dr. Goldberg"
        Cells(iCntr, 11) = "Group 1"
    ElseIf (Cells(iCntr, 4) = "PSI 9") Then
        Cells(iCntr, 10) = "Dr. Goldberg"
        Cells(iCntr, 11) = "Group 1"
    ElseIf (Cells(iCntr, 4) = "PSI 12") Then
        Cells(iCntr, 10) = "Dr. Goldberg"
        Cells(iCntr, 11) = "Group 1"
    ElseIf (Cells(iCntr, 4) = "PSI 12") Then
        Cells(iCntr, 10) = "Dr. Goldberg"
        Cells(iCntr, 11) = "Group 1"
    ElseIf (Cells(iCntr, 4) = "PSI 14") Then
        Cells(iCntr, 10) = "Dr. Goldberg"
        Cells(iCntr, 11) = "Group 1"
    ElseIf (Cells(iCntr, 4) = "PSI 6") Then
        Cells(iCntr, 10) = "Dr. Silverman"
        Cells(iCntr, 11) = "Group 2"
    ElseIf (Cells(iCntr, 4) = "PSI 10") Then
        Cells(iCntr, 10) = "Dr. Silverman"
        Cells(iCntr, 11) = "Group 2"
    ElseIf (Cells(iCntr, 4) = "PSI 11") Then
        Cells(iCntr, 10) = "Dr. Silverman"
        Cells(iCntr, 11) = "Group 2"
    ElseIf (Cells(iCntr, 4) = "PSI 13") Then
        Cells(iCntr, 10) = "Dr. Silverman"
        Cells(iCntr, 11) = "Group 2"
    ElseIf (Cells(iCntr, 4) = "PSI 15") Then
        Cells(iCntr, 10) = "Dr. Silverman"
        Cells(iCntr, 11) = "Group 2"
    ElseIf (Cells(iCntr, 4) = "PSI 3") Then
        Cells(iCntr, 10) = "Liz Mac"
        Cells(iCntr, 11) = "Group 3"
    End If
Next

'_Renaming conventions_
'----------------------
'(cells(Row,1)

'Rehab                = Jackson Rehab Hospital
'JMH                  = Jackson Memorial Hospital
'Jackson North        = Jackson North Medical Center
'JMH - Holtz Children = The Women's Hospital and Children's Hospital at Jackson
'Jackson South        = Jackson South Medical Center

For iCntr = lRow To 2 Step -1
    If (Cells(iCntr, 1) = "Rehab") Then
        Cells(iCntr, 1) = "Jackson Rehab Hospital"
        Cells(iCntr, 9) = "" ' dont remember why this part
    ElseIf (Cells(iCntr, 1) = "JMH") Then
            Cells(iCntr, 1) = "Jackson Memorial Hospital"
    ElseIf (Cells(iCntr, 1) = "Jackson North") Then
            Cells(iCntr, 1) = "Jackson North Medical Center"
    ElseIf (Cells(iCntr, 1) = "JMH - Holtz Children") Then
            Cells(iCntr, 1) = "The Women's Hospital and Children's Hospital at Jackson"
    ElseIf (Cells(iCntr, 1) = "Jackson South") Then
            Cells(iCntr, 1) = "Jackson South Medical Center"
    End If
Next


'_Deleting "not wanted" rows_
'---------------------------
'If not one of the following combination will be deleted. (cells(Row,1) & cells(Row,2))

'"Jackson Memorial Hospital" / "All Teaching Hospitals"
'"Jackson North Medical Center" / "Community Hospitals"
'"The Women's Hospital and Children's Hospital at Jackson" / "Standalone Pediatric Hospitals"
'"Jackson South Medical Center" / "Trauma Level 2"
'"Jackson Memorial Hospital" / "All hospitals"

For iCntr = lRow To 2 Step -1
    If (Cells(iCntr, 1) = "Jackson Rehab Hospital") And (Cells(iCntr, 2) = "All hospitals") Then
    ElseIf (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All Teaching Hospitals") Then
    ElseIf (Cells(iCntr, 1) = "Jackson North Medical Center") And (Cells(iCntr, 2) = "Community Hospitals") Then
    ElseIf (Cells(iCntr, 1) = "The Women's Hospital and Children's Hospital at Jackson") And (Cells(iCntr, 2) = "Standalone Pediatric Hospitals") Then
    ElseIf (Cells(iCntr, 1) = "Jackson South Medical Center") And (Cells(iCntr, 2) = "Trauma Level 2") Then
    ElseIf (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals") Then ' New line added to keep this combo and delete it at the end.
    Else
        Rows(iCntr).Delete
    End If
Next


'_Adjustment of last row_
'------------------------
ActiveSheet.UsedRange
Range("A1").Select


'_Monthly summary variant as per request_
'----------------------------------------

'_Searching if month are showing on on current data

With Range("C:C")
Set aJan = .Find("Jan", LookIn:=xlValues, LookAt:=xlPart)
Set aFeb = .Find("Feb", LookIn:=xlValues, LookAt:=xlPart)
Set aMar = .Find("Mar", LookIn:=xlValues, LookAt:=xlPart)
Set aApr = .Find("Apr", LookIn:=xlValues, LookAt:=xlPart)
Set aMay = .Find("May", LookIn:=xlValues, LookAt:=xlPart)
Set aJun = .Find("Jun", LookIn:=xlValues, LookAt:=xlPart)
Set aJul = .Find("Jul", LookIn:=xlValues, LookAt:=xlPart)
Set aAug = .Find("Aug", LookIn:=xlValues, LookAt:=xlPart)
Set aSep = .Find("Sep", LookIn:=xlValues, LookAt:=xlPart)
Set aOct = .Find("Oct", LookIn:=xlValues, LookAt:=xlPart)
Set aNov = .Find("Nov", LookIn:=xlValues, LookAt:=xlPart)
Set aDec = .Find("Dec", LookIn:=xlValues, LookAt:=xlPart)
End With

'-------------------------------------------------------------------------
'_ TESTING TO REMOVE CODE ABOVE AND BELOW THIS SEGMENT_
'Dim aMonthsArr(1 To 12) As Range
'
'With Range("C:C")
'Set aMonthsArr(1) = .Find("Jan", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(2) = .Find("Feb", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(3) = .Find("Mar", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(4) = .Find("Apr", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(5) = .Find("May", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(6) = .Find("Jun", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(7) = .Find("Jul", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(8) = .Find("Aug", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(9) = .Find("Sep", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(10) = .Find("Oct", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(11) = .Find("Nov", LookIn:=xlValues, LookAt:=xlPart)
'Set aMonthsArr(12) = .Find("Dec", LookIn:=xlValues, LookAt:=xlPart)
'End With
'
'For Each Element In aMonthsArr
'        If Not aMonthsArr(Element) Is Nothing Then
'        aSubtotalMonthly (aJan)
'Next Element
'-------------------------------------------------------------------------


'_Conditional for each variable above, if month present in data then calling function"aSubtotalMonthly"...
'_ ...on month,function will do the manth and print value at the end of the current data
If Not aJan Is Nothing Then 'Or Not IsNull(aJan) Then
aSubtotalMonthly (aJan)
End If
If Not aFeb Is Nothing Then
aSubtotalMonthly (aFeb)
End If
If Not aMar Is Nothing Then 'Or Not IsNull(aMar) Then
aSubtotalMonthly (aMar)
End If
If Not aApr Is Nothing Then 'Or Not IsNull(aApr) Then
aSubtotalMonthly (aApr)
End If
If Not aMay Is Nothing Then 'Or Not IsNull(aMay) Then
aSubtotalMonthly (aMay)
End If
If Not aJun Is Nothing Then 'Or Not IsNull(aJun) Then
aSubtotalMonthly (aJun)
End If
If Not aJul Is Nothing Then 'Or Not IsNull(aJul) Then
aSubtotalMonthly (aJul)
End If
If Not aAug Is Nothing Then 'Or Not IsNull(aAug) Then
aSubtotalMonthly (aAug)
End If
If Not aSep Is Nothing Then 'Or Not IsNull(aSep) Then
aSubtotalMonthly (aSep)
End If
If Not aOct Is Nothing Then 'Or Not IsNull(aOct) Then
aSubtotalMonthly (aOct)
End If
If Not aNov Is Nothing Then 'Or Not IsNull(aNov) Then
aSubtotalMonthly (aNov)
End If
If Not aDec Is Nothing Then 'Or Not IsNull(aDec) Then
aSubtotalMonthly (aDec)
End If

'_Adjustment of last row_
'------------------------
ActiveSheet.UsedRange
Range("A1").Select

'_Removing "that" combo
'----------------------
'And (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals")_ was added to remove that combo from month stats
For iCntr = lRow To 2 Step -1
    If (Cells(iCntr, 1) = "Jackson Memorial Hospital") And (Cells(iCntr, 2) = "All hospitals") Then
        Rows(iCntr).Delete
    End If
Next


'_Cosmetics & Visualization back ON_
'-----------------------------------
ActiveSheet.Range("A1:K1").AutoFilter
ActiveWindow.DisplayGridlines = False
OptimizeVBA False

'_Adjustment of last row_
'------------------------
ActiveSheet.UsedRange
Range("A1").Select

'_Final message box_
'-------------------
SecondsElapsed = Round(Timer - StartTime, 2) 'Determine how many seconds code took to run
MsgBox ("Macro executed successfully in " & SecondsElapsed & " seconds") '--'

End Sub


