Attribute VB_Name = "Reklamace"

Sub Reklamace_Entropy_Energy_CZ()
'2015_07_30
'copies data from Report to Entropy File

'antiepilepticka tabulka
Application.ScreenUpdating = False

'otevri cilovy report
Workbooks.Open "W:\W46_Quality_System_Management\Reporty\Entropy\Czech Complaints Template 2016.xls"


'rename both files
Dim wb1 As Workbook
Dim wb2 As Workbook
Dim Source As Worksheet
Dim Report As Worksheet


Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks("Czech Complaints Template 2016.xls") 'change CZ vs SK
Set Source = wb1.Worksheets("Reklamace")
Set Report = wb2.Worksheets("04. Quality Data Collection")

wb1.Activate
'do Source pak lze pøidávat øádky a staèí jen správnì nastavit "i" pro rùzná makra

Dim i As Integer
i = 37 'i - na kterém øádku je leden dané zemì - source
        '2015 CZ = 25, 2015 SK = 55
        '2016 CZ = 37, 2016 SK = 79


'----------------prekopiruj data
Dim a As Integer 'sloupec, na kterem zacina import - source
Dim b As Integer 'radek, na kterem zacina import - report


'Complaints - soucet vsech defektu pro jednotlive typy napoju
        a = 9
        b = 9
    Do While a < 15
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 3).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
    Loop
    
'Comments
        a = 15
        b = 9
    Do While a < 21
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 7).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
    Loop
    
'Sales
        a = 21
        b = 111
    Do While a < 27
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 3).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
    Loop

'Reklamace
        a = 28
        b = 213
    Do While a < 123
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 3).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
        If a = 43 Then
            a = a + 1
        ElseIf a = 59 Then
            a = a + 1
        ElseIf a = 75 Then
            a = a + 1
        ElseIf a = 91 Then
            a = a + 1
        ElseIf a = 107 Then
            a = a + 1
        End If
           
    Loop


'otevøi explorer W:\W46_Quality_System_Management\Reporty\Entropy
Call Shell("explorer.exe" & " " & "W:\W46_Quality_System_Management\Reporty\Entropy", vbNormalFocus)

'message
wb1.Activate
MsgBox "Hurá, hotovo!"

End Sub
Sub Reklamace_Entropy_Energy_SK()
'2015_07_30
'copies data from Report to Entropy File

'antiepilepticka tabulka
Application.ScreenUpdating = False

'otevri cilovy report
Workbooks.Open "W:\W46_Quality_System_Management\Reporty\Entropy\Slovakia Complaints Template 2016.xls"


'rename both files
Dim wb1 As Workbook
Dim wb2 As Workbook
Dim Source As Worksheet
Dim Report As Worksheet


Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks("Slovakia Complaints Template 2016.xls") 'change CZ vs SK
Set Source = wb1.Worksheets("Reklamace")
Set Report = wb2.Worksheets("04. Quality Data Collection")

wb1.Activate
'do Source pak lze pøidávat øádky a staèí jen správnì nastavit "i" pro rùzná makra

Dim i As Integer
i = 79 'i - na kterém øádku je leden dané zemì - source
        '2015 CZ = 25, 2015 SK = 55
        '2016 CZ = 37, 2016 SK = 79
        

'----------------prekopiruj data
Dim a As Integer 'sloupec, na kterem zacina import - source
Dim b As Integer 'radek, na kterem zacina import - report


'Complaints - soucet vsech defektu pro jednotlive typy napoju
        a = 9
        b = 9
    Do While a < 15
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 3).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
    Loop
    
'Comments
        a = 15
        b = 9
    Do While a < 21
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 7).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
    Loop
    
'Sales
        a = 21
        b = 111
    Do While a < 27
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 3).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
    Loop

'Reklamace
        a = 28
        b = 213
    Do While a < 123
        Source.Range(Cells(i, a), Cells((11 + i), a)).Copy
        Report.Cells(b, 3).PasteSpecial Paste:=xlPasteValues
        a = a + 1
        b = b + 17
        If a = 43 Then
            a = a + 1
        ElseIf a = 59 Then
            a = a + 1
        ElseIf a = 75 Then
            a = a + 1
        ElseIf a = 91 Then
            a = a + 1
        ElseIf a = 107 Then
            a = a + 1
        End If
           
    Loop


'otevøi explorer W:\W46_Quality_System_Management\Reporty\Entropy
Call Shell("explorer.exe" & " " & "W:\W46_Quality_System_Management\Reporty\Entropy", vbNormalFocus)

'message
wb1.Activate
MsgBox "Hurá, hotovo!"

End Sub
