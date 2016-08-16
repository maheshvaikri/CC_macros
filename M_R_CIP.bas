Attribute VB_Name = "M_R_CIP"
Sub CIP_6ventilu_PL2()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "CIP"
    ActiveCell.FormulaR1C1 = "PL2"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CIP"
    ActiveCell.Offset(0, -2).Range("A1:C1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
    'dopise komentare
    ActiveCell.Offset(-7, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "vodni cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "sirupova cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "plnici ventil"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
      Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(-7, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = "N/A"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:C1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C8"), Type:= _
        xlFillDefault
        
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(0, -4).Range("A1").Select
End Sub
Sub CIP_12ventilu_PL2()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "CIP"
    ActiveCell.FormulaR1C1 = "PL2"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CIP"
    ActiveCell.Offset(0, -2).Range("A1:C1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
    'dopise komentare
    ActiveCell.Offset(-13, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "vodni cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "sirupova cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "plnici ventil"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
      Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(-13, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = "N/A"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:C1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C14"), Type:= _
        xlFillDefault
        
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(0, -4).Range("A1").Select
End Sub
Sub CIP_12ventilu_PL4()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "CIP"
    ActiveCell.FormulaR1C1 = "PL4"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CIP"
    ActiveCell.Offset(0, -2).Range("A1:C1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
    'dopise komentare
    ActiveCell.Offset(-13, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "vodni cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "sirupova cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "plnici ventil"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
      Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(-13, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = "N/A"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:C1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C14"), Type:= _
        xlFillDefault
        
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(0, -4).Range("A1").Select
End Sub

Sub CIP_6ventilu_PL4()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "CIP"
    ActiveCell.FormulaR1C1 = "PL4"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CIP"
    ActiveCell.Offset(0, -2).Range("A1:C1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
    'dopise komentare
    ActiveCell.Offset(-7, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "vodni cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "sirupova cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "plnici ventil"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
      Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(-7, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = "N/A"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:C1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C8"), Type:= _
        xlFillDefault
        
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(0, -4).Range("A1").Select
End Sub
Sub CIP_5ventilu_PL6()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "CIP"
    ActiveCell.FormulaR1C1 = "PL6"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CIP"
    ActiveCell.Offset(0, -2).Range("A1:C1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
    
    'dopise komentare
    ActiveCell.Offset(-6, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "vodni cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "sirupova cesta"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "plnici ventil"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
      Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(-6, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = "N/A"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:C1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:C7"), Type:= _
        xlFillDefault
        
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(0, -4).Range("A1").Select
End Sub


Sub COP_PL2()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "PL2"
    ActiveCell.FormulaR1C1 = "PL2"
    ActiveCell.Offset(0, -1).Range("A1:B1").Copy
        ActiveCell.Offset(1, -1).Range("A1:A16").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'dopise typ
    ActiveCell.Offset(-1, 2).Range("A1:A6").Value = "COP - plnici ventil"
    ActiveCell.Offset(5, 2).Range("A1:A6").Value = "COP - snift ventil"
    ActiveCell.Offset(11, 2).Range("A1:A4").Value = "COP - ostatni"
    ActiveCell.Offset(15, 2).Range("A1:A1").Value = "COP - vzduch"

    Application.CutCopyMode = False
    
    'dopise komentar
     ActiveCell.Offset(11, 3).Range("A1:A1").Value = "ster pas1"
     ActiveCell.Offset(12, 3).Range("A1:A1").Value = "ster pas2"
    ActiveCell.Offset(13, 3).Range("A1:A1").Value = "ster pas3"
    ActiveCell.Offset(14, 3).Range("A1:A1").Value = "ster uzaviracka1"
      
    Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(11, 4).Range("A1:A5").Value = "N/A"
    ActiveCell.Offset(11, 6).Range("A1:E5").Value = "N/A"
    
    ActiveCell.Offset(-1, 5).Range("A1:A12").Value = "N/A"
    ActiveCell.Offset(-1, 8).Range("A1:C12").Value = "N/A"
    
       
           
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(-1, 4).Range("A1").Select
End Sub


Sub COP_PL4()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "PL4"
    ActiveCell.FormulaR1C1 = "PL4"
    ActiveCell.Offset(0, -1).Range("A1:B1").Copy
        ActiveCell.Offset(1, -1).Range("A1:A16").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'dopise typ
    ActiveCell.Offset(-1, 2).Range("A1:A6").Value = "COP - plnici ventil"
    ActiveCell.Offset(5, 2).Range("A1:A6").Value = "COP - snift ventil"
    ActiveCell.Offset(11, 2).Range("A1:A4").Value = "COP - ostatni"
    ActiveCell.Offset(15, 2).Range("A1:A1").Value = "COP - vzduch"

    Application.CutCopyMode = False
    
    'dopise komentar
     ActiveCell.Offset(11, 3).Range("A1:A1").Value = "ster pas1"
     ActiveCell.Offset(12, 3).Range("A1:A1").Value = "ster pas2"
    ActiveCell.Offset(13, 3).Range("A1:A1").Value = "ster pas3"
    ActiveCell.Offset(14, 3).Range("A1:A1").Value = "ster uzaviracka1"
      
    Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(11, 4).Range("A1:A5").Value = "N/A"
    ActiveCell.Offset(11, 6).Range("A1:E5").Value = "N/A"
    
    ActiveCell.Offset(-1, 5).Range("A1:A12").Value = "N/A"
    ActiveCell.Offset(-1, 8).Range("A1:C12").Value = "N/A"
    
       
           
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(-1, 4).Range("A1").Select
End Sub

Sub COP_PL6()
'
' Macro1 Macro
'
'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

'nakopiruje datum a "PL6"
    ActiveCell.FormulaR1C1 = "PL6"
    ActiveCell.Offset(0, -1).Range("A1:B1").Copy
        ActiveCell.Offset(1, -1).Range("A1:A9").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'dopise typ
    ActiveCell.Offset(-1, 2).Range("A1:A5").Value = "COP - plnici ventil"
    ActiveCell.Offset(4, 2).Range("A1:A4").Value = "COP - ostatni"
    ActiveCell.Offset(8, 2).Range("A1:A1").Value = "COP - vzduch"

    Application.CutCopyMode = False
    
    'dopise komentar
     ActiveCell.Offset(4, 3).Range("A1:A1").Value = "ster pas1"
     ActiveCell.Offset(5, 3).Range("A1:A1").Value = "ster pas2"
    ActiveCell.Offset(6, 3).Range("A1:A1").Value = "ster pas3"
    ActiveCell.Offset(7, 3).Range("A1:A1").Value = "ster uzaviracka1"
      
    Application.CutCopyMode = False
      
    'vyplní N/A
    ActiveCell.Offset(4, 4).Range("A1:A5").Value = "N/A"
    ActiveCell.Offset(4, 6).Range("A1:E5").Value = "N/A"
    
    ActiveCell.Offset(-1, 5).Range("A1:A5").Value = "N/A"
    ActiveCell.Offset(-1, 8).Range("A1:C5").Value = "N/A"
    
       
           
     'posune mys, aby slo rovnou policka vyplnovat
    ActiveCell.Offset(-1, 4).Range("A1").Select
End Sub
