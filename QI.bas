Attribute VB_Name = "QI"
Sub QI_import()
Attribute QI_import.VB_ProcData.VB_Invoke_Func = " \n14"
'Misa Bryxova ziska exportni excel z CC databaze, nezmenena data pouze prekopiruje do listu BryxovaIN
'ze 4 radkove tabulky se automaticky vytvori kompaktnejsi 2radkova.
' Makro Naimportuje data z sheetu BryxovaIN (z oblasti A15:N17)
' kliknout do mesice, kliknout do mesice, do ktereho se reportuje v CZ casti



'update 2016_01_19:
        ' jen posunut offset SK data o 55 radku oproti CZ


'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False

    ActiveCell.Select
    'CZ data
    Sheets("BryxovaIN").Select
    Range("C16:N16").Select
    Selection.Copy
    Sheets("QI").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'SK data
    Sheets("BryxovaIN").Select
    Range("C17:N17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("QI").Select
    ActiveCell.Offset(55, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    MsgBox ("Hotovo. Pìkný den :)")
    
End Sub
Sub ScaleAxes()
  With ActiveChart.Axes(xlCategory, xlPrimary)
  
    .CategoryType = xlTimeScale
    .TickLabels.NumberFormat = "yyyy.mm"
    .MaximumScale = ActiveSheet.Range("M3").Value
    .MinimumScale = ActiveSheet.Range("M2").Value
    '.MajorUnit = 1
    '.MajorUnitScale = xlMonths
    End With
End Sub
