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
Sub Falls()

'Last Update: 6/22/2020
'Comments   : "Trauma 4A" removed from Jackson Rehab and now staying under Jackson Main

'Update: 5/21/2020
'Comments   : Some "Rehab Units" rebranded as "Jackson Rehab. Hospital"

'Script to clean and format raw data extracted from Quantros'

'_VALIDATING FILE BEFORE STARTING TRANSFORMATIONS_
'-------------------------------------------
If Range("A1") <> "Falls and Med Events" Then
    MsgBox ("Incorrect file! Please try again with another file.")
Exit Sub
End If

'_Creating some variables to time the macro_
'-------------------------------------------
Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer 'Remember time when macro starts


'_Disabling worksheet recalculation, screen updating, statusbar updating_
'------------------------------------------------------------------------
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False


'_GLOBAL VARIABLES IN USE_
'-------------------------
'arrays initial position is 0 on -VBA-

'ActiveCell.SpecialCells(xlLastCell).Select
lRow = ActiveCell.SpecialCells(xlLastCell).Row
Location_Array = Array("Holtz Women's and Children's", "Jackson Behavioral Health Hospital", "Jackson Memorial - Main Campus", "Jackson North Medical Center", "Jackson South Community Hospital", "Jackson South Medical Center")
event_Array = Array("Rehab-2 Annex  Pediatrics", "Rehab 2 Annex Adult", "Rehab Unit", "LRC6", "LRC7", "LRC8", "Rehab-Physical Therapy-Orthotics/Prosthetics", "Rehab-Occupational Therapy")
lMonth = Month(DateAdd("m", -1, Date)) ' Extrating position of previous month

'_REMOVING FIRST 2 EMPTY ROWS_
'-----------------------------
Rows("1:2").Select
Selection.Delete Shift:=xlUp


'_REMOVING ALL DATES NOT FROM LAST MONTH_
'----------------------------------------

'Looping deleting dates not from last month
For iCntr = lRow To 2 Step -1
    If Cells(iCntr, 5).Value2 = "'-" Then ' some records dont have a date but cell contains invalid value "'-" with this step those records get deleted
        Rows(iCntr).Delete
    ElseIf (Month(Cells(iCntr, 5))) <> lMonth Then
        Rows(iCntr).Delete
    End If
Next


'_REMOVING UNWANTED LOCATIONS_
'-----------------------------
'Any location -NOT IN- "Location_Array" will be deleted
For iCntr = lRow To 2 Step -1
    If IsInArray(Cells(iCntr, 1), Location_Array) Then ' IsInArray() funtion will compare value from cell with array declared at the beginning of script
    Else
        Rows(iCntr).Delete
    End If
Next

'Removing event where "No Person Involved"_
'------------------------------------------
aVariable = "No Person Involved"
For iCntr = lRow To 1 Step -1
    If InStr(Cells(iCntr, 9), aVariable) Then ' InStr() built-in function will look for substring(2nd parameter) inside substring(1st parameter) returning position, in this cases any position will indicate ocurrence
        Rows(iCntr).Delete
    End If
Next

'"Rehab-2 Annex  Pediatrics","Rehab 2 Annex Adult", "Rehab Unit", "Trauma 4A" events ocurrences facility update_
'---------------------------------------------------------------------------------------------------------------
bVariable = "Jackson Rehab. Hospital"
For iCntr = lRow To 1 Step -1
    ' if cell in (row,15) -IN- "event_Array" will be relocated on "Jackson Rehab. Hospital"
    If IsInArray(Cells(iCntr, 15), event_Array) Then '
        Cells(iCntr, 1).Value2 = bVariable
    ' if "Jackson South Community Hospital" or "Jackson South Medical Center" will be relocated on "Jackson South Behavioral Health"
    ElseIf Cells(iCntr, 15) = "Behavioral Health" And (Cells(iCntr, 1) = "Jackson South Medical Center" Or Cells(iCntr, 1) = "Jackson South Community Hospital") Then
        Cells(iCntr, 1).Value2 = "Jackson South Behavioral Health"

' old code left over
'    ElseIf Cells(iCntr, 15) = "Behavioral Health" And (Cells(iCntr, 1) = "Jackson South Medical Center" or Cells(iCntr, 1) = "Jackson South Community Hospital") Then
'        Cells(iCntr, 1).Value2 = "Jackson South Behavioral Health"
    End If
Next

'"Creating individual tabs_
'--------------------------
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet
Sheets.Add After:=Sheets(Sheets.Count) 'Adding extra sheet

' Renaming new sheets_
'---------------------
Sheets(2).Name = "Holtz Falls"
Sheets(3).Name = "Jackson Behavioral Falls"
Sheets(4).Name = "Jackson Memorial Falls"
Sheets(5).Name = "Jackson North Falls"
Sheets(6).Name = "Jackson South Falls"
Sheets(7).Name = "Jackson Rehab. Hospital Falls"
'-------------------------------------------------------------------
Sheets(8).Name = "Holtz Med. Errors" 'Renaming new sheet
Sheets(9).Name = "Jackson Behavioral Med. Errors" 'Renaming new sheet
Sheets(10).Name = "Jackson Memorial Med. Errors" 'Renaming new sheet
Sheets(11).Name = "Jackson North Med. Errors" 'Renaming new sheet
Sheets(12).Name = "Jackson South Med. Errors" 'Renaming new sheet
Sheets(13).Name = "Jackson Rehab. Hospital Errors" 'Renaming new sheet


'This was requested later on
'---------------------------
Sheets(14).Name = "Jackson S. Beh. Health Falls" 'Renaming new sheet
Sheets(15).Name = "Jackson S. Beh. Health Errors" 'Renaming new sheet
'---------------------------


'Copying first row on all new sheets
'-----------------------------------
Sheets(1).Range("1:1").Copy Destination:=Sheets(2).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(3).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(4).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(5).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(6).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(7).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(8).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(9).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(10).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(11).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(12).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(13).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(14).Range("1:1")
Sheets(1).Range("1:1").Copy Destination:=Sheets(15).Range("1:1")


'Creating individual counters for moving of records
'--------------------------------------------------
Dim acounter As Long
Dim bCounter As Long
Dim cCounter As Long
Dim dCounter As Long
Dim eCounter As Long
Dim fCounter As Long
Dim gCounter As Long
Dim hCounter As Long
Dim iCounter As Long
Dim jCounter As Long
Dim kCounter As Long
Dim lCounter As Long
Dim mCounter As Long
Dim nCounter As Long


'Initializing counters
'---------------------
acounter = Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row
bCounter = Sheets(8).Cells(Rows.Count, "A").End(xlUp).Row
cCounter = Sheets(3).Cells(Rows.Count, "A").End(xlUp).Row
dCounter = Sheets(9).Cells(Rows.Count, "A").End(xlUp).Row
eCounter = Sheets(4).Cells(Rows.Count, "A").End(xlUp).Row
fCounter = Sheets(10).Cells(Rows.Count, "A").End(xlUp).Row
gCounter = Sheets(5).Cells(Rows.Count, "A").End(xlUp).Row
hCounter = Sheets(11).Cells(Rows.Count, "A").End(xlUp).Row
iCounter = Sheets(6).Cells(Rows.Count, "A").End(xlUp).Row
jCounter = Sheets(12).Cells(Rows.Count, "A").End(xlUp).Row
kCounter = Sheets(7).Cells(Rows.Count, "A").End(xlUp).Row
lCounter = Sheets(13).Cells(Rows.Count, "A").End(xlUp).Row
mCounter = Sheets(14).Cells(Rows.Count, "A").End(xlUp).Row
nCounter = Sheets(15).Cells(Rows.Count, "A").End(xlUp).Row

'Setting variables for comparison
'--------------------------------
aFall = "Falls & Slips (Patient)"
aError = "Medication (Patient)"

aHoltz = "Holtz Women's and Children's"
aJBH = "Jackson Behavioral Health Hospital"
aJM = "Jackson Memorial - Main Campus"
aJN = "Jackson North Medical Center"
aJS = "Jackson South Community Hospital"
aJC = "Jackson South Medical Center"
aJRh = "Jackson Rehab. Hospital"
aJSB = "Jackson South Behavioral Health"


'Timer up to this point in case needed
'-------------------------------------
'SecondsElapsed = Round(Timer - StartTime, 2) 'Determine how many seconds code took to run
'MsgBox ("Time so far: " & SecondsElapsed & " seconds") '--'
                                                                                                                                                                                                
                                                                                                                                                                                                
'Moving records to new sheets according to classification and division
'---------------------------------------------------------------------
Sheets(1).Select
For iCntr = lRow To 1 Step -1
    Cells(iCntr, 2).Select
    If Cells(iCntr, 1).Value2 = aHoltz And Cells(iCntr, 9).Value2 = aFall Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(2).Range("A" & acounter + 1)
        acounter = Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aHoltz And Cells(iCntr, 9).Value2 = aError Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(8).Range("A" & bCounter + 1)
        bCounter = Sheets(8).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJBH And Cells(iCntr, 9).Value2 = aFall Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(3).Range("A" & cCounter + 1)
        cCounter = Sheets(3).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJBH And Cells(iCntr, 9).Value2 = aError Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(9).Range("A" & dCounter + 1)
        dCounter = Sheets(9).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJM And Cells(iCntr, 9).Value2 = aFall Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(4).Range("A" & eCounter + 1)
        eCounter = Sheets(4).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJM And Cells(iCntr, 9).Value2 = aError Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(10).Range("A" & fCounter + 1)
        fCounter = Sheets(10).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJN And Cells(iCntr, 9).Value2 = aFall Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(5).Range("A" & gCounter + 1)
        gCounter = Sheets(5).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJN And Cells(iCntr, 9).Value2 = aError Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(11).Range("A" & hCounter + 1)
        hCounter = Sheets(11).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf (Cells(iCntr, 1).Value2 = aJS Or Cells(iCntr, 1).Value2 = aJC) And Cells(iCntr, 9).Value2 = aFall Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(6).Range("A" & iCounter + 1)
        iCounter = Sheets(6).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf (Cells(iCntr, 1).Value2 = aJS Or Cells(iCntr, 1).Value2 = aJC) And Cells(iCntr, 9).Value2 = aError Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(12).Range("A" & jCounter + 1)
        jCounter = Sheets(12).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJRh And Cells(iCntr, 9).Value2 = aError Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(13).Range("A" & kCounter + 1)
        kCounter = Sheets(13).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJRh And Cells(iCntr, 9).Value2 = aFall Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(7).Range("A" & lCounter + 1)
        lCounter = Sheets(7).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJSB And Cells(iCntr, 9).Value2 = aFall Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(14).Range("A" & mCounter + 1)
        mCounter = Sheets(14).Cells(Rows.Count, "A").End(xlUp).Row
    ElseIf Cells(iCntr, 1).Value2 = aJSB And Cells(iCntr, 9).Value2 = aError Then
        Rows(ActiveCell.Row).Copy Destination:=Sheets(15).Range("A" & nCounter + 1)
        nCounter = Sheets(15).Cells(Rows.Count, "A").End(xlUp).Row
    End If
Next



'"Just a little format per sheet
'---------------------------------------------
Sheets("Holtz Falls").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Holtz Med. Errors").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson Behavioral Falls").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson Behavioral Med. Errors").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson Memorial Falls").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson Memorial Med. Errors").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson North Falls").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson North Med. Errors").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson South Falls").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson South Med. Errors").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson Rehab. Hospital Falls").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson Rehab. Hospital Errors").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson S. Beh. Health Falls").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------
Sheets("Jackson S. Beh. Health Errors").Select
ActiveWindow.DisplayGridlines = False
Cells.Select
Range("D1").Activate
Selection.RowHeight = 55
'---------------------------------------------

'Cursor back to first sheet
'--------------------------
Sheets(1).Activate


'_Creating new workbooks and saving_
'--------------------------------------

ActiveWorkbook.SaveAs Filename:="Falls_file.xlsx"
Workbooks.Add.SaveAs Filename:="Errors_file.xlsx"

'_Moving "errors" sheets to 2nd workbook created_
'-----------------------------------------------------
Workbooks("Falls_file.xlsx").Sheets("Holtz Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson Behavioral Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson Memorial Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson North Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson South Med. Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson Rehab. Hospital Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)
Workbooks("Falls_file.xlsx").Sheets("Jackson S. Beh. Health Errors").Move After:=Workbooks("Errors_file.xlsx").Sheets(Workbooks("Errors_file.xlsx").Sheets.Count)

'_Deleting original raw data sheet and 1st sheet from 2nd workbook_
'------------------------------------------------------------------
'-Workbooks("Falls_file.xlsx").Sheets(1).Delete '- Keeping original as per Afirah request
Workbooks("Errors_file.xlsx").Sheets(1).Delete  '- Deleting 1st empty sheet
Workbooks("Errors_file.xlsx").Sheets(1).Activate '- Cursor back to first sheet


'_Enabling worksheet recalculation, screen updating, statusbar updating_
'-----------------------------------------------------------------------
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

'_Renaming workbooks if needed_
'-----------------------------------------------------------------------
smonth = MonthName(lMonth)
syear = VBA.DateTime.Year(Date)
stitle1 = "Preliminary Falls & Injury Report " & smonth & " " & syear & ".xlsx"
stitle2 = "Preliminary Med.Errors Report " & smonth & " " & syear & ".xlsx"

ActiveWorkbook.SaveAs stitle2
Kill "Errors_file.xlsx" '  If you still want your old copy, then don't use this line.
Workbooks("Falls_file.xlsx").SaveAs stitle1
Kill "Falls_file.xlsx" '  If you still want your old copy, then don't use this line.

'_Macro Timer_
'-------------------------------------------
SecondsElapsed = Round(Timer - StartTime, 2) 'Determine how many seconds code took to run
MsgBox ("Macro executed in " & SecondsElapsed & " seconds") '--'
End Sub
