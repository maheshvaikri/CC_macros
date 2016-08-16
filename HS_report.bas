Attribute VB_Name = "HS_report"
Sub HS_CZ()
'copies data from Report to W:\W46_Quality_System_Management\Reporty\Lucas\....

'antiepilepticka tabulka
Application.ScreenUpdating = False

'otevri cilovy report
Workbooks.Open "W:\W46_Quality_System_Management\Reporty\Lucas\2015 Leading Indicators Country Template_CZ.xlsx"

'rename both files
Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks("2015 Leading Indicators Country Template_CZ.xlsx")

Set Source = wb1.Worksheets("BOZP")
Set Report_CZ = wb2.Worksheets("Country Template")


'----------------prekopiruj data
'NearMiss
Source.Range("V35:V46").Copy
Report_CZ.Range("B3").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'ClosedCAP
Source.Range("Z35:Z46").Copy
Report_CZ.Range("B4").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'NearMiss loni
Source.Range("V23:V34").Copy
Report_CZ.Range("B5").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'Toolbox
Source.Range("AB35:AB46").Copy
Report_CZ.Range("B9").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'WalkTheTalk
Source.Range("AD35:AD46").Copy
Report_CZ.Range("B11").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'RootCause
Source.Range("AF35:AF46").Copy
Report_CZ.Range("B13").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'LTA incidents
Source.Range("I35:I46").Copy
Report_CZ.Range("B14").PasteSpecial Transpose:=True, Paste:=xlPasteValues


'-------------------zprava na konec
MsgBox ("CZ Report pro Iaina aktualizován, nezapomeò ho poslat, Karle.")

'------------ulož oba soubory
wb1.Save
wb2.Save


'otevøi explorer W:\W46_Quality_System_Management\Reporty\Lucas
Call Shell("explorer.exe" & " " & "W:\W46_Quality_System_Management\Reporty\Lucas", vbNormalFocus)



End Sub

Sub HS_SK()
'copies data from Report to W:\W46_Quality_System_Management\Reporty\Lucas\....

'antiepilepticka tabulka
Application.ScreenUpdating = False

'otevri cilovy report
Workbooks.Open "W:\W46_Quality_System_Management\Reporty\Lucas\2015 Leading Indicators Country Template_SK.xlsx"

'rename both files
Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks("2015 Leading Indicators Country Template_SK.xlsx")

Set Source = wb1.Worksheets("BOZP")
Set Report_CZ = wb2.Worksheets("Country Template")


'----------------prekopiruj data
'NearMiss
Source.Range("W35:W46").Copy
Report_CZ.Range("B3").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'ClosedCAP
Source.Range("AA35:AA46").Copy
Report_CZ.Range("B4").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'NearMiss loni
Source.Range("W23:W34").Copy
Report_CZ.Range("B5").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'Toolbox
Source.Range("AC35:AC46").Copy
Report_CZ.Range("B9").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'WalkTheTalk
Source.Range("AE35:AE46").Copy
Report_CZ.Range("B11").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'RootCause
Source.Range("AG35:AG46").Copy
Report_CZ.Range("B13").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'LTA incidents
Source.Range("J35:J46").Copy
Report_CZ.Range("B14").PasteSpecial Transpose:=True, Paste:=xlPasteValues


'-------------------zprava na konec
MsgBox ("SK_Report pro Iaina aktualizován, nezapomeò ho poslat, Lubko.")

'------------ulož oba soubory
wb1.Save
wb2.Save

'otevøi explorer W:\W46_Quality_System_Management\Reporty\Lucas
Call Shell("explorer.exe" & " " & "W:\W46_Quality_System_Management\Reporty\Lucas", vbNormalFocus)
End Sub

Sub HS_CZ_SK()
'copies data from Report to W:\W46_Quality_System_Management\Reporty\Lucas\....

'antiepilepticka tabulka
Application.ScreenUpdating = False

'otevri cilovy report
Workbooks.Open "W:\W46_Quality_System_Management\Reporty\Lucas\2015 Leading Indicators Country Template_CZ.xlsx"

'rename both files
Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks("2015 Leading Indicators Country Template_CZ.xlsx")

Set Source = wb1.Worksheets("BOZP")
Set Report_CZ = wb2.Worksheets("Country Template")


'----------------prekopiruj data
'NearMiss
Source.Range("V35:V46").Copy
Report_CZ.Range("B3").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'ClosedCAP
Source.Range("Z35:Z46").Copy
Report_CZ.Range("B4").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'NearMiss loni
Source.Range("V23:V34").Copy
Report_CZ.Range("B5").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'Toolbox
Source.Range("AB35:AB46").Copy
Report_CZ.Range("B9").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'WalkTheTalk
Source.Range("AD35:AD46").Copy
Report_CZ.Range("B11").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'RootCause
Source.Range("AF35:AF46").Copy
Report_CZ.Range("B13").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'LTA incidents
Source.Range("I35:I46").Copy
Report_CZ.Range("B14").PasteSpecial Transpose:=True, Paste:=xlPasteValues


'------------ulož oba soubory
wb1.Save
wb2.Save


'otevri cilovy report
Workbooks.Open "W:\W46_Quality_System_Management\Reporty\Lucas\2015 Leading Indicators Country Template_SK.xlsx"

'rename both files
Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks("2015 Leading Indicators Country Template_SK.xlsx")

Set Source = wb1.Worksheets("BOZP")
Set Report_CZ = wb2.Worksheets("Country Template")


'----------------prekopiruj data
'NearMiss
Source.Range("W35:W46").Copy
Report_CZ.Range("B3").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'ClosedCAP
Source.Range("AA35:AA46").Copy
Report_CZ.Range("B4").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'NearMiss loni
Source.Range("W23:W34").Copy
Report_CZ.Range("B5").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'Toolbox
Source.Range("AC35:AC46").Copy
Report_CZ.Range("B9").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'WalkTheTalk
Source.Range("AE35:AE46").Copy
Report_CZ.Range("B11").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'RootCause
Source.Range("AG35:AG46").Copy
Report_CZ.Range("B13").PasteSpecial Transpose:=True, Paste:=xlPasteValues

'LTA incidents
Source.Range("J35:J46").Copy
Report_CZ.Range("B14").PasteSpecial Transpose:=True, Paste:=xlPasteValues


'------------ulož oba soubory
wb1.Save
wb2.Save


'-------------------zprava na konec
MsgBox ("Reporty pro CZ i SK byly aktualizovány, teï je tøeba doplnit data na OneDrive - CZ i SK")


'otevøi explorer W:\W46_Quality_System_Management\Reporty\Lucas
Call Shell("explorer.exe" & " " & "W:\W46_Quality_System_Management\Reporty\Lucas", vbNormalFocus)
End Sub
