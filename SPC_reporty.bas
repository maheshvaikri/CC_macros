Attribute VB_Name = "SPC_reporty"
Sub SPC_OneDrive_CZ()
'2016_01_20


'copies data from Report to excel, data then have to be copy pasted to OneDrive excel

'je treba definovat
    'i
    'wbname (target file)
    'Set Report = wb2.Worksheets("Czech") sheet je bud Czech nebo Slovakia
    


'antiepilepticka tabulka
Application.ScreenUpdating = False


'rename both files
Dim wb1 As Workbook
Dim wb2 As Workbook
Dim Source As Worksheet
Dim Report As Worksheet

'-------change the target file name here (once)
wbname = "CZ_2016 Cpk Reporting Template_TEST.xlsx"


'otevri cilovy report
Workbooks.Open ("W:\W46_Quality_System_Management\Reporty\Entropy\" & wbname)

Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks(wbname)
Set Source = wb1.Worksheets("SPC")
Set Report = wb2.Worksheets("Czech")



wb1.Activate
'do Source pak lze pøidávat øádky a staèí jen správnì nastavit "i" pro rùzná makra



Dim i As Integer
i = 35 'i - na kterém øádku je leden dané zemì - source
        '2016 CZ = 35, 2016 SK = 77
        
Dim a As Integer 'øádek na kterem zacina import - source
Dim b As Integer 'sloupec, na kterem zacina import - report


'----------------prekopiruj data


'PL4-RGB

       a = i
       b = 4
       
    Do While a < 47
        Source.Range(Cells(a, 6), Cells(a, 8)).Copy
        Report.Cells(4, b).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 4
    Loop
        

'PL2-PET
        a = i
        b = 4
       
    Do While a < 47
        Source.Range(Cells(a, 10), Cells(a, 13)).Copy
        Report.Cells(5, b).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 4
    Loop

'PL6-CAN
        a = i
        b = 4
       
    Do While a < 47
        Source.Range(Cells(a, 15), Cells(a, 17)).Copy
        Report.Cells(6, b).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 4
    Loop

'PL8-APET (pozor, neni CO2)
        a = i
        b = 4
       
    Do While a < 47
        Source.Range(Cells(a, 19), Cells(a, 19)).Copy
        Report.Cells(7, b).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 4
    Loop
    
    a = i
    b = 6
       
    Do While a < 47
        Source.Range(Cells(a, 20), Cells(a, 21)).Copy
        Report.Cells(7, b).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 4
    Loop
'------------TADY JSEM SKONCILA-------------------------


'otevøi explorer W:\W46_Quality_System_Management\Reporty\Entropy
Call Shell("explorer.exe" & " " & "W:\W46_Quality_System_Management\Reporty\Entropy", vbNormalFocus)

'message
wb1.Activate
MsgBox "Hurá, povedlo se, teï to vlož na OneDrive"

End Sub

Sub SPC_OneDrive_SK()
'2016_01_20


'copies data from Report to excel, data then have to be copy pasted to OneDrive excel

'je treba definovat
    'i
    'wbname (target file)
    'Set Report = wb2.Worksheets("Czech") sheet je bud Czech nebo Slovakia
    


'antiepilepticka tabulka
Application.ScreenUpdating = False


'rename both files
Dim wb1 As Workbook
Dim wb2 As Workbook
Dim Source As Worksheet
Dim Report As Worksheet

'-------change the target file name here (once)
wbname = "SK_2016 Cpk Reporting Template_TEST.xlsx"


'otevri cilovy report
Workbooks.Open ("W:\W46_Quality_System_Management\Reporty\Entropy\" & wbname)

Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks(wbname)
Set Source = wb1.Worksheets("SPC")
Set Report = wb2.Worksheets("Slovakia")



wb1.Activate
'do Source pak lze pøidávat øádky a staèí jen správnì nastavit "i" pro rùzná makra



Dim i As Integer
i = 77 'i - na kterém øádku je leden dané zemì - source
        '2016 CZ = 35, 2016 SK = 77
        
Dim a As Integer 'øádek na kterem zacina import - source
Dim b As Integer 'sloupec, na kterem zacina import - report


'----------------prekopiruj data


'RGB

       a = i
       b = 4
       
    Do While a < 89
        Source.Range(Cells(a, 6), Cells(a, 8)).Copy
        Report.Cells(4, b).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 4
    Loop
        

'PL2-PET
        a = i
        b = 4
       
    Do While a < 89
        Source.Range(Cells(a, 15), Cells(a, 18)).Copy
        Report.Cells(6, b).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 4
    Loop




'otevøi explorer W:\W46_Quality_System_Management\Reporty\Entropy
Call Shell("explorer.exe" & " " & "W:\W46_Quality_System_Management\Reporty\Entropy", vbNormalFocus)

'message
wb1.Activate
MsgBox "Hurá, povedlo se, teï to vlož na OneDrive"

End Sub
