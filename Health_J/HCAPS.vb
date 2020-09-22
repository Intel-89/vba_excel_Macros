Sub HCAPS_pdf()

'Last Update: 2/13/2020

'Script to clean & prep. data'

'_code from here

Application.DisplayAlerts = False

'_Deleting unnecessary sheets_
Sheets("Table 3").Delete
Sheets("Table 4").Delete
Sheets("Table 8").Delete
Sheets("Table 9").Delete
Sheets("Table 10").Delete
Sheets("Table 11").Delete
Sheets("Table 12").Delete
Sheets("Table 13").Delete
Sheets("Table 14").Delete
Sheets("Table 19").Delete
Sheets("Table 20").Delete
Sheets("Table 21").Delete
Sheets("Table 22").Delete
Sheets("Table 23").Delete
Sheets("Table 24").Delete
Sheets("Table 25").Delete
Sheets("Table 26").Delete
Sheets("Table 27").Delete
Sheets("Table 28").Delete
Sheets("Table 29").Delete

'_Renaming remaining sheets_
Sheets("Table 1").Name = "Comm w Nurses"
Sheets("Table 2").Name = "Comm w Doctors"
Sheets("Table 5").Name = "Resp of Hosp Staff"
Sheets("Table 6").Name = "Comm About Mediciness"
Sheets("Table 7").Name = "Discharge Information"
Sheets("Table 15").Name = "Cleanliness"
Sheets("Table 16").Name = "Quietness"
Sheets("Table 17").Name = "Would Recommend"
Sheets("Table 18").Name = "Overall Rating"

'_Looping through all sheets_

For aCntr = 1 To 9
    Sheets(aCntr).Select
    'Modif. Headers
    Rows(1).EntireRow.Delete
    Rows(1).EntireRow.Delete

    Range("A1") = ActiveSheet.Name

    Range("A2") = "Lowest % Top Box"
    Range("B2") = "% Rank"
    Range("C2") = "Lowest % Top Box"
    Range("D2") = "% Rank"
    Range("E2") = "Lowest % Top Box"
    Range("F2") = "% Rank"
    
    Rows(1).EntireRow.Delete
    ActiveWindow.DisplayGridlines = False
Next

Application.DisplayAlerts = True

MsgBox ("Macro executed successfully.")
'
End Sub
