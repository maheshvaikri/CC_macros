Attribute VB_Name = "M_R_Data_combine"
Sub Combine2()


    Dim J As Integer
    'aby to nebyla epilepticka tabulka
    Application.ScreenUpdating = False
    
    On Error Resume Next 'vubec nevim, co to melo znamenat
    'Sheets(1).Select
    'Worksheets.Add ' add a sheet in first place
    'Sheets(1).Name = "Data"

    'delete whole sheet
    Sheets(1).Cells.Clear 'pak zmìnit na skuteèné poøadí

    ' copy headings
    Sheets(2).Activate 'nebo jine cislo sheetu, odkud se maji brat nazvy sloupcu
       
    Range("A3:K3").Select
    Selection.Copy Destination:=Sheets(1).Range("A2")

    ' work through sheets
    For J = 2 To 8 ' from sheet 2 to last sheet: For J = 3 To Sheets.Count, For J = 3 To 8 pro konkretni listy
            Sheets(J).Visible = True
            Sheets(J).Activate ' make the sheet active
        Range("A4").Select
        Selection.CurrentRegion.Select ' select all cells in this sheets

        ' select all lines except title
        Selection.Offset(2, 0).Select
        Selection.Resize(Selection.Rows.Count - 2).Select
        

        ' copy cells selected in the new sheet on last line
        Selection.Copy
        Sheets(1).Range("A65536").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        
   ' For J = 2 To 7 ' from sheet 2 to last sheet: For J = 3 To Sheets.Count, For J = 3 To 8 pro konkretni listy
       
    Sheets(J).Range("A4").Select
    Application.CutCopyMode = False
                           
    Next

'aby se udelal fitr a data automaticky seradila podle data
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A3:AH65536")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets(1).Select
    Range("A2").Select
'naformatovani jako tabulka
    Dim tbl As ListObject
    Dim rng As Range

    Set rng = Range(Range("A2"), Range("A2").CurrentRegion)
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
    
    'prvni sloupec jako datum
  
  Columns(1).NumberFormat = "m/d/yyyy"
   
   Sheets(1).Range("A3").Select
   Application.CutCopyMode = False
      
End Sub
Sub CopyLastRow()

Application.ScreenUpdating = False

'vezme hodnotu z bunky = celkovy pocet opakovani posledniho radku
Dim i As Integer
i = Range("A2").Value

'vybere a kopiruje posledni radek
ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Select
ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Copy

'vlozi ho tolikrat pod tabulku
ActiveCell.Offset(1, 0).Range("A1:A" & i - 1).Select
   ActiveSheet.Paste
 
End Sub
Sub Dopln_NA()
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "N/A"
End Sub

Sub duplicate(n As Integer)

    Dim found As Range
    Dim i As Integer, _
        numRows As Integer

    Set found = Range("A:A").Find("*", _
                                  Range("A1"), _
                                  LookIn:=xlFormulas, _
                                  searchdirection:=xlPrevious)

    If Not found Is Nothing Then
        numRows = found.Row
        For i = numRows To 1 Step -1
            Range(Cells((found.Row - 1) * n + 1, 1), _
                  Cells(found.Row * n, 1)).Value = found.Value
            If found.Row > 1 Then
                Set found = found.Offset(-1, 0)
            End If
        Next i
    End If

End Sub
