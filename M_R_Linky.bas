Attribute VB_Name = "M_R_Linky"

Sub Linka_PL4()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "PL4"
    ActiveCell.FormulaR1C1 = "PL4"
    ActiveCell.Offset(0, -1).Range("A1:B1").Copy
        ActiveCell.Offset(1, -1).Range("A1:A10").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'dopise typ
    ActiveCell.Offset(-1, 2).Range("A1").Value = "produkcni voda"
    ActiveCell.Offset(0, 2).Range("A1").Value = "vyplachova voda"
    ActiveCell.Offset(1, 2).Range("A1:A6").Value = "obal"
    ActiveCell.Offset(7, 2).Range("A1").Value = "vicka"
    ActiveCell.Offset(8, 2).Range("A1:A2").Value = "vzduch"
    ActiveCell.Offset(8, 3).Range("A1").Value = "plnic"
    ActiveCell.Offset(9, 3).Range("A1").Value = "mycka"
    Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(1, 5).Range("A1:A6").Value = "N/A" 'obal
    ActiveCell.Offset(7, 4).Range("A1:A3").Value = "N/A" 'CPM
    ActiveCell.Offset(9, 5).Range("A1:A1").Value = "N/A" 'coli
    ActiveCell.Offset(-1, 8).Range("A1:C11").Value = "N/A"
           
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(-1, 4).Range("A1").Select
End Sub

Sub Linka_PL2()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "PL2"
    ActiveCell.FormulaR1C1 = "PL2"
    ActiveCell.Offset(0, -1).Range("A1:B1").Copy
        ActiveCell.Offset(1, -1).Range("A1:A10").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'dopise typ
    ActiveCell.Offset(-1, 2).Range("A1").Value = "produkcni voda"
    ActiveCell.Offset(0, 2).Range("A1").Value = "vyplachova voda"
    ActiveCell.Offset(1, 2).Range("A1:A6").Value = "obal"
    ActiveCell.Offset(7, 2).Range("A1").Value = "vicka"
    ActiveCell.Offset(8, 2).Range("A1:A2").Value = "vzduch"
    ActiveCell.Offset(8, 3).Range("A1").Value = "plnic"
    ActiveCell.Offset(9, 3).Range("A1").Value = "vyplachovac"
    Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(1, 5).Range("A1:A6").Value = "N/A" 'obal
    ActiveCell.Offset(7, 4).Range("A1:A3").Value = "N/A" 'CPM
    ActiveCell.Offset(-1, 8).Range("A1:C11").Value = "N/A" 'ostatni parametry
    
           
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(-1, 4).Range("A1").Select
End Sub

Sub Linka_PL6()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "PL6"
    ActiveCell.FormulaR1C1 = "PL6"
    ActiveCell.Offset(0, -1).Range("A1:B1").Copy
        ActiveCell.Offset(1, -1).Range("A1:A10").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'dopise typ
    ActiveCell.Offset(-1, 2).Range("A1").Value = "produkcni voda"
    ActiveCell.Offset(0, 2).Range("A1").Value = "vyplachova voda"
    ActiveCell.Offset(1, 2).Range("A1:A6").Value = "obal"
    ActiveCell.Offset(7, 2).Range("A1").Value = "vicka"
    ActiveCell.Offset(8, 2).Range("A1:A2").Value = "vzduch"
    ActiveCell.Offset(8, 3).Range("A1").Value = "plnic"
    ActiveCell.Offset(9, 3).Range("A1").Value = "vyplachovac"
    Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(1, 5).Range("A1:A6").Value = "N/A" 'obal
    ActiveCell.Offset(7, 4).Range("A1:A3").Value = "N/A" 'CPM
    ActiveCell.Offset(-1, 8).Range("A1:C11").Value = "N/A" 'ostatni parametry
           
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(-1, 4).Range("A1").Select
End Sub

Sub Linka_PL5()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "PL5"
    ActiveCell.FormulaR1C1 = "PL5"
    ActiveCell.Offset(0, -1).Range("A1:B1").Copy
        ActiveCell.Offset(1, -1).Range("A1:A7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'dopise typ
    ActiveCell.Offset(0, 2).Range("A1:A6").Value = "vicko"
    ActiveCell.Offset(6, 2).Range("A1").Value = "vzduch"
    Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(0, 4).Range("A1:A7").Value = "N/A"
    ActiveCell.Offset(-1, 8).Range("A1:C8").Value = "N/A"
               
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(-1, 2).Range("A1").Select
End Sub

