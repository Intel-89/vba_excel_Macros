
Sub test_123()
'
'_Setting Logical Variables
Dim aPhyID As String
Dim aPhyName As String
Dim aMeas As String
Dim aPeriod As String
Dim aLastSheet As Integer
aBLANK = ""

'Adding Sheet at Beginning of Workbook
Worksheets.Add Before:=Sheets(1)
'Saving last sheet position
aLastSheet = ActiveWorkbook.Worksheets.Count
Worksheets(1).Select
lRowSheet1 = 0
'lRowSheet1 = ActiveCell.SpecialCells(xlLastCell).Row


'_ STARTING SHEETS LOOP
For aSheetCounter = 2 To aLastSheet
'If aSheetCounter = 3 Then
'    MsgBox ("3 Page now")
'End If
Worksheets(aSheetCounter).Select

lRow = ActiveCell.SpecialCells(xlLastCell).Row
iCntr2 = 0 'This counter will prevent looping over the original amount of rows with values in case of empty rows, based on loop starting from top instead of bottom
    
    '_ STARTING INDIVIDUAL SHEET LOOP until last cell with value

    'Insert 4 Columns to the left of Column A
    Columns("A:D").Insert Shift:=xlToRight
    'Insert 1 Row for Headers
    Rows(1).Insert
    Cells(1, 1) = "Period"
    Cells(1, 2) = "Physician Name"
    Cells(1, 3) = "Physician ID"
    Cells(1, 4) = "Measure Set"
    Cells(1, 5) = "Hospital"
    Cells(1, 6) = "Measure"
    Cells(1, 7) = "Measure Name"
    Cells(1, 8) = "Provider Numerator"
    Cells(1, 9) = "Provider Denominator"
    Cells(1, 10) = "HCO Rate"
    Cells(1, 11) = "Overall Mean"
    Cells(1, 12) = "Overall Hospital Top 10%"
    Cells(1, 13) = "95% CI of Provider Rate"
    
    '------------------------ '- Starting current sheet iteration
    For iCntr = 2 To lRow
    
        If Cells(iCntr, 5) = aBLANK Then
            If Cells(iCntr, 7) = aBLANK Then
                iCntr2 = iCntr2 + 1
                If iCntr2 < lRow Then
                    Rows(iCntr).Delete
                    iCntr = iCntr - 1
                End If
            
            ElseIf Left(Cells(iCntr, 7), 15) = "Physician ID : " Then
                aPhyID = Right(Cells(iCntr, 7), (Len(Cells(iCntr, 7)) - 15))
                Rows(iCntr).Delete
                iCntr = iCntr - 1
            Else
                Rows(iCntr).Delete
                iCntr2 = iCntr2 + 1
                iCntr = iCntr - 1
            End If

'        '_EXTRACTING Physician ID
'        ElseIf Left(Cells(iCntr, 7), 15) = "Physician ID : " Then
'            aPhyID = Right(Cells(iCntr, 7), (Len(Cells(iCntr, 7)) - 15))
'            Rows(iCntr).Delete
'            iCntr = iCntr - 1
    
        '_EXTRACTING Physician Name
        ElseIf Left(Cells(iCntr, 5), 17) = "Physician Name : " Then
            aPhyName = Right(Cells(iCntr, 5), (Len(Cells(iCntr, 5)) - 17))
            Rows(iCntr).Delete
            iCntr2 = iCntr2 + 1
            iCntr = iCntr - 1
    
        '_EXTRACTING Period
        ElseIf Left(Cells(iCntr, 5), 14) = "From Period : " Then
            aPeriod = Right(Cells(iCntr, 5), (Len(Cells(iCntr, 5)) - 14))
            Rows(iCntr).Delete
            iCntr2 = iCntr2 + 1
            iCntr = iCntr - 1
    
        '_EXTRACTING Measure
        ElseIf Left(Cells(iCntr, 5), 14) = "Measure Set : " Then
            aMeas = Right(Cells(iCntr, 5), (Len(Cells(iCntr, 5)) - 14))
            Rows(iCntr).Delete
            iCntr2 = iCntr2 + 1
            iCntr = iCntr - 1
            
        '_Deleting rows with something on column 5 but no relevant other value
        ElseIf Cells(iCntr, 6) = aBLANK Then
            Rows(iCntr).Delete
            iCntr = iCntr - 1
        
        '_Deleting rows Original Header
        ElseIf Cells(iCntr, 5) = "Measure" And Cells(iCntr, 6) = "Measure Name" Then
            Rows(iCntr).Delete
            iCntr2 = iCntr2 + 1
            iCntr = iCntr - 1
        
        Else
            Cells(iCntr, 1) = aPeriod
            Cells(iCntr, 2) = aPhyName
            Cells(iCntr, 3) = aPhyID
            Cells(iCntr, 4) = aMeas
        
        End If
    Next iCntr
    '------------------------ '- End of current Sheet iteration
    
      
    '_Adjustment of last row_ '-
    '------------------------ '-
    ActiveSheet.UsedRange     '-
    Range("A1").Select        '-
    lRow = ActiveCell.SpecialCells(xlLastCell).Row '-
    '---------------------------
    
    For iCntr = lRow To 2 Step -1
        If Cells(iCntr, 5) = "Hospital" Then
            Rows(iCntr).Delete
        End If
    Next
    
    '_Adjustment of last row_ '-
    '------------------------ '-
    ActiveSheet.UsedRange     '-
    Range("A1").Select        '-
    lRow = ActiveCell.SpecialCells(xlLastCell).Row '-
    '---------------------------

    'Select ALL from active sheet, Copy and Paste on Sheet 1
    ActiveSheet.UsedRange.Copy
    Sheets(1).Select
    Dim LastRow As Long '
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row '
    Cells(LastRow, 1).Offset(1, 0).Select '
'    Cells.SpecialCells(xlCellTypeLastCell).Select
'    ActiveCell.Offset(1, 0).Select
'    'ActiveSheet.Range("A":lRowSheet1+1).Select

    '_Finding first empty cell on A column to populate metric calculated above_
    '--------------------------------------
'    Range("A1").Select        '-
'    Range("A1").End(xlDown).Offset(1, 0).Select 'Selecting
    ActiveSheet.Paste
    '------------------------- '-

Next aSheetCounter 'Increase counter for next sheet execution

'Looping deleting extra headers

lRow = ActiveCell.SpecialCells(xlLastCell).Row

For iCntr = lRow To 3 Step -1
    If Cells(iCntr, 1) = "Period" Then ' some records dont have a date but cell contains invalid value "'-" with this step those records get deleted
        Rows(iCntr).Delete
    End If
Next

Sheets(1).Name = "Compilation"
End Sub








'Cells.Replace What:=Chr(10), Replacement:=" ", LookAt:=xlPart, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False



''_Timer ON_
''----------
'Dim StartTime As Double
'Dim SecondsElapsed As Double
'StartTime = Timer 'Remember time when macro starts
'
''_Disabling worksheet recalculation, screen updating, statusbar updating_
''------------------------------------------------------------------------
'Application.Calculation = xlCalculationManual
'Application.ScreenUpdating = False
'Application.DisplayStatusBar = False
'Application.DisplayAlerts = False
'
''------------------------------------------------------------------------------------------------------------------------------------
''------------------------------------------------------------------------------------------------------------------------------------
'
''_Validating -Service- column before macro run_
''----------------------------------------------
'With Range("A:A")
'Set avalidation = .Find("Service", LookIn:=xlValues, LookAt:=xlWhole)
'End With
'
'If avalidation Is Nothing Then
'    MsgBox ("-Service- column not found! Please try again with another file.")
'Exit Sub
'End If
'
'
''_REMOVING FIRST ROWS until "Service" is found_
''----------------------------------------------
'
'iCntr = 1
'service_bool = True
'
'While service_bool
'    If Cells(1, 1) <> "Service" Then
'        Rows(1).Delete
'    Else
'        service_bool = False
'    End If
'Wend
'
'
''---------------------------
''_Adjustment of last row_ '-
''------------------------ '-
'ActiveSheet.UsedRange     '-
'Range("A1").Select        '-
'lRow = ActiveCell.SpecialCells(xlLastCell).Row '-
''---------------------------
'
'
''_Finding if ED or not_
''----------------------
'EDReport = True
'aReportType = "ED" 'Added later
'
'If Cells(2, 1) <> "Emergency Department" Then
'    EDReport = False
'    aReportType = "HCAHPS"  'Added later
''    MsgBox ("This report will NOT be ED type.")
''Else
''    MsgBox ("This report will BE ED type.")
'End If
'
'
''_DELETING/Clearing_values COLUMNS TO MATCH TABLEAU PRESET DATASOURCE_
''---------------------------------------------------------------------
'Columns(1).Delete 'Service
'Columns(5).Delete 'Survey Type
'Columns(7).Delete 'Benchmarking Option
'Columns(7).Delete 'Benchmarking Period
'Columns(8).Delete 'All PG Database N
'Columns(5).Insert 'n
'Columns(7).Cut
'Columns(6).Insert
'
'
''_Renaming Headers_
''------------------
'Range("A1") = "Facility/Unit"
'Range("B1") = "IT Discharge date"
'Range("C1") = "Unit"
'Range("D1") = "Question"
'Range("E1") = "Response"
'Range("F1") = "n"
'Range("G1") = "Raw Phone %"
'Range("H1") = "All DB %ile"
'Range("I1") = "Percentile Rank"
'Range("J1") = "Goal"
'Range("K1") = "Discharge Month"
'Range("L1") = "HCAHPS or ED"
'Range("M1") = "Raw/Mean"
'Range("N1") = "Adjusted Top Box"
'Range("O1") = "% Rank Goal"
'
'
''_Removing dash/quotation from date/unit column(s)
''-------------------------------
'Columns(2).Replace What:=" - ", Replacement:=" ", LookAt:=xlPart
'Columns(3).Replace What:="'", Replacement:="", LookAt:=xlPart
'
'
''_Formating years to 2 digits_
''-----------------------------
'Dim aArry() As String
'For iCntr = lRow To 2 Step -1
'    If (Cells(iCntr, 2) <> "Total") Then
'        aArry = Split(Cells(iCntr, 2))
'        var1 = Format(CDate(aArry(0)), "mm/dd/yy") 'Format(strDate, "DD.MM.YYYY")
'        var2 = Format(CDate(aArry(1)), "mm/dd/yy")
'        Cells(iCntr, 2) = CStr(var1) & " " & CStr(var2)
'    End If
'Next
'
'
''_Percentile Rank filled up
''--------------------------
'For iCntr = lRow To 2 Step -1
'    If (Cells(iCntr, 8) <> "N/A") And (Cells(iCntr, 8) <> "Invalid") Then
'        Cells(iCntr, 9) = Cells(iCntr, 8) / 100
'    End If
'Next
''Range("I:I").NumberFormat = "0.00%"
'
'
''_Removing Discharge Date = "Total" with "YTD'20"
''------------------------------------------------
'For iCntr = lRow To 2 Step -1
'    If (Cells(iCntr, 2) = "Total") Then
'        Cells(iCntr, 2) = "YTD'20"
'    End If
'Next
'
''_Adding values to column -HCAHPS or ED-
''---------------------------------------
'For iCntr = lRow To 2 Step -1
'    If EDReport Then
'        Cells(iCntr, 12) = "ED"
'    Else
'        Cells(iCntr, 12) = "HCAHPS"
'    End If
'Next
'
''_Adding values to column -Discharge Month-
''--------------------------------------
'For iCntr = lRow To 2 Step -1
'    If Cells(iCntr, 2) = "YTD'20" Then
'        Cells(iCntr, 11) = ""
'        'Range(iCntr & "11").ClearContents
'    Else
'        Cells(iCntr, 11) = Split(Cells(iCntr, 2), " ")(0)
'        'Range(iCntr & "11").Value = Split(Cells(iCntr, 2), " ")(0)
'    End If
'Next
'
'
''_Renaming Facilities
''--------------------
'For iCntr = lRow To 2 Step -1
'    If Cells(iCntr, 1) = "'Holtz Children's & Women's Hospital'" Then
'        Cells(iCntr, 1) = "Holtz Womens and Childrens Hospital"
'    ElseIf Cells(iCntr, 1) = "'Holtz Children's Hospital - Phone'" Then
'        Cells(iCntr, 1) = "Holtz Womens and Childrens Hospital"
'    ElseIf Cells(iCntr, 1) = "'Jackson Memorial Hospital'" Then
'        Cells(iCntr, 1) = "Jackson Memorial Hospital"
'    ElseIf Cells(iCntr, 1) = "'Jackson Memorial Hospital - Phone'" Then
'        Cells(iCntr, 1) = "Jackson Memorial Hospital"
'    ElseIf Cells(iCntr, 2) = "'Jackson North Medical Center'" Then
'        Cells(iCntr, 1) = "Jackson North Medical Center"
'    ElseIf Cells(iCntr, 1) = "'Jackson North Medical Center - Phone'" Then
'        Cells(iCntr, 1) = "Jackson North Medical Center"
'    ElseIf Cells(iCntr, 1) = "'Jackson South Community Hospital - Phone'" Then
'        Cells(iCntr, 1) = "Jackson South Medical Center"
'    ElseIf Cells(iCntr, 2) = "'Jackson South Medical Center'" Then
'        Cells(iCntr, 1) = "Jackson South Medical Center"
'    'ElseIf Cells(iCntr, 2) = "" Then
'    '    Cells(iCntr, 2) = "Jackson Heath System"
'    End If
'Next
'
''_Renaming Questions
''-------------------
'For iCntr = lRow To 2 Step -1
'    If Cells(iCntr, 4) = "*Rate hospital 0-10" Then
'        Cells(iCntr, 4) = "Overall Rating"
'    ElseIf Cells(iCntr, 4) = "*Recommend the hospital" Then
'        Cells(iCntr, 4) = "Would Recommend"
'    ElseIf Cells(iCntr, 4) = "*Comm w/ Nurses Domain Performance" Then
'        Cells(iCntr, 4) = "Communication with Nurses"
'    ElseIf Cells(iCntr, 4) = "*Response of Hosp Staff Domain Performance" Then
'        Cells(iCntr, 4) = "Responsiveness of Hospital Staff"
'    ElseIf Cells(iCntr, 4) = "*Comm w/ Doctors Domain Performance" Then
'        Cells(iCntr, 4) = "Communication with Doctors"
'    ElseIf Cells(iCntr, 4) = "*Cleanliness of hospital environment" Then
'        Cells(iCntr, 4) = "Cleanliness of Hospital Environment"
'    ElseIf Cells(iCntr, 4) = "*Quietness of hospital environment" Then
'        Cells(iCntr, 4) = "Quietness of Hospital Environment"
'    ElseIf Cells(iCntr, 4) = "*Comm About Medicines Domain Performance" Then
'        Cells(iCntr, 4) = "Communication About Medications"
'    ElseIf Cells(iCntr, 4) = "*Discharge Information Domain Performance" Then
'        Cells(iCntr, 4) = "Discharge Information"
'    End If
'Next
'
''_Removing Facility/Unit = "Total"
''---------------------------------
'For iCntr = lRow To 2 Step -1
'    If Cells(iCntr, 1) = "Total" Then
'        Rows(iCntr).Delete
'    End If
'Next
'
''---------------------------
''_Adjustment of last row_ '-
''------------------------ '-
'ActiveSheet.UsedRange     '-
'Range("A1").Select        '-
'lRow = ActiveCell.SpecialCells(xlLastCell).Row '-
''---------------------------
'
''_VLookup on reference tbl
''-------------------------
'For iCntr = lRow To 2 Step -1
'    aLookupValue = Cells(iCntr, 1) & " " & aReportType & " " & Cells(iCntr, 4)
'    'MsgBox (aLookupValue)
'    Cells(iCntr, 10) = Application.VLookup(aLookupValue, Sheets("tbl_lookup").Range("A2:F319"), 5, False)
'    Cells(iCntr, 15) = Application.VLookup(aLookupValue, Sheets("tbl_lookup").Range("A2:F319"), 6, False)
'Next
'
''------------------------------------------------------------------------------------------------------------------------------------
''------------------------------------------------------------------------------------------------------------------------------------
'
''_Timer OFF_
''-----------
'SecondsElapsed = Round(Timer - StartTime, 2) 'Determine how many seconds code took to run
'MsgBox ("Macro executed in " & SecondsElapsed & " seconds") '--'
'
''_Enabling worksheet recalculation, screen updating, statusbar updating_
''-----------------------------------------------------------------------
'Application.Calculation = xlCalculationAutomatic
'Application.ScreenUpdating = True
'Application.DisplayStatusBar = True
'Application.DisplayAlerts = True

'
'End Sub
