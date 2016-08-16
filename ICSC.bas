Attribute VB_Name = "ICSC"
Sub ICSC_CZ()
'copies data from Report to W:\WU2_ICSC_reporty\2_Manufacturing Praha
'2016_01_21_update pro 2016
'pro dalsi roky staci jen upravit wbaddress a wbname a zkontrolovat, jestli export zacina v E6 a cilovy soubor taky v E6.
'A jestli je MsgBox ("ICSC data za Prahu pøekopírována do Míšina Manufacturing reportu") stale aktualni

'antiepilepticka tabulka
        Application.ScreenUpdating = False


 '-------change the target file name here (once)
        wbaddress = "W:\WU2_ICSC_reporty\2_Manufacturing Praha\2016\Monthly Reports\"
        wbname = "16_Region__Manufacturing_ver 0.3-Prague.xlsx"


'otevri cilovy report
        Workbooks.Open (wbaddress & wbname) 'abych to nemusela upravovat pokazde, kdyz nekdo zmeni nazev workbooku

'rename both files
        Dim wb1 As Workbook
        Dim wb2 As Workbook
        Dim Source As Worksheet
        Dim Report_CZ As Worksheet
        
        Set wb1 = Workbooks("Report.xlsm")
        Set wb2 = Workbooks(wbname)
        Set Source = wb1.Worksheets("ICSC")
        Set Report_CZ = wb2.Worksheets("data")



'-------------------zprava na uvod
        MsgBox ("zadej heslo bohemia")



'----------------prekopiruj data
        '-------MTD---------
            'complaints
            Source.Range("E6:P7").Copy
            Report_CZ.Range("E6").PasteSpecial Paste:=xlPasteValues
            
            'indexy
            Source.Range("E9:P15").Copy
            Report_CZ.Range("E9").PasteSpecial Paste:=xlPasteValues
            
            'Energy consumption
            Source.Range("E17:P18").Copy
            Report_CZ.Range("E17").PasteSpecial Paste:=xlPasteValues
            
            'Solid waste produced
            Source.Range("E20:P21").Copy
            Report_CZ.Range("E20").PasteSpecial Paste:=xlPasteValues
            
            'Solid waste recycled
            Source.Range("E23:P24").Copy
            Report_CZ.Range("E23").PasteSpecial Paste:=xlPasteValues
            
            'Near Miss
            Source.Range("E26:P29").Copy
            Report_CZ.Range("E26").PasteSpecial Paste:=xlPasteValues
            
            'CAP
            Source.Range("E32:P33").Copy
            Report_CZ.Range("E32").PasteSpecial Paste:=xlPasteValues
        
        
        '------YTD
            'indexy YTD
            Source.Range("AR9:BC11").Copy
            Report_CZ.Range("AR9").PasteSpecial Paste:=xlPasteValues
            
            'CPK YTD
            Source.Range("AR13:BC13").Copy
            Report_CZ.Range("AR13").PasteSpecial Paste:=xlPasteValues



'-------------------zprava na konec
        MsgBox ("ICSC data za Prahu pøekopírována do Míšina Manufacturing reportu")

'------------ulož oba soubory
        wb1.Save
        wb2.Save


'otevøi explorer W:\WU2_ICSC_reporty\2_Manufacturing Praha\...?
        Call Shell("explorer.exe" & " " & wbaddress, vbNormalFocus)



End Sub

Sub ICSC_SK()
'copies data from Report to W:\WU2_ICSC_reporty\3_Manufacturing Luka

'antiepilepticka tabulka
Application.ScreenUpdating = False

'otevri cilovy report
Workbooks.Open "W:\WU2_ICSC_reporty\3_Manufacturing Luka\16_Region__Manufacturing_ver 0.3 2016.xlsx"

'-------------------zprava na konec
MsgBox ("zadej heslo Bratislava")


'rename both files
Set wb1 = Workbooks("Report.xlsm")
Set wb2 = Workbooks("16_Region__Manufacturing_ver 0.3 2016.xlsx")

Set Source = wb1.Worksheets("ICSC")
Set Report_SK = wb2.Worksheets("data")


'----------------prekopiruj data
'-------MTD---------
'complaints
Source.Range("E140:P141").Copy
Report_SK.Range("E140").PasteSpecial Paste:=xlPasteValues

'indexy
Source.Range("E143:P149").Copy
Report_SK.Range("E143").PasteSpecial Paste:=xlPasteValues

'Energy consumption
Source.Range("E151:P152").Copy
Report_SK.Range("E151").PasteSpecial Paste:=xlPasteValues

'Solid waste produced
Source.Range("E154:P155").Copy
Report_SK.Range("E154").PasteSpecial Paste:=xlPasteValues

'Solid waste recycled
Source.Range("E157:P158").Copy
Report_SK.Range("E157").PasteSpecial Paste:=xlPasteValues

'Near Miss
Source.Range("E160:P163").Copy
Report_SK.Range("E160").PasteSpecial Paste:=xlPasteValues

'CAP
Source.Range("E166:P167").Copy
Report_SK.Range("E166").PasteSpecial Paste:=xlPasteValues


'------YTD
'indexy YTD
Source.Range("AR143:BC145").Copy
Report_SK.Range("AR143").PasteSpecial Paste:=xlPasteValues

'CPK YTD
Source.Range("AR147:BC147").Copy
Report_SK.Range("AR147").PasteSpecial Paste:=xlPasteValues



'-------------------zprava na konec
MsgBox ("ICSC data za Luku pøekopírována do Zdenina Manufacturing reportu")

'------------ulož oba soubory
wb1.Save
wb2.Save


'otevøi explorer W:\WU2_ICSC_reporty\2_Manufacturing Praha\Monthly Reporty
Call Shell("explorer.exe" & " " & "W:\WU2_ICSC_reporty\3_Manufacturing Luka", vbNormalFocus)



End Sub

