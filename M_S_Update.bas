Attribute VB_Name = "M_S_Update"

Sub CopyData()

'copies data from ...Result to ...Stats

'antiepilepticka tabulka
Application.ScreenUpdating = False

'vypnout automaticke prepocitavani
Application.Calculation = xlManual

' ## 1. open ...Results
Workbooks.Open "W:\W92_Laboratory_CZ\4. MIKROBIOLOGIE\Monitoring_MIBI\CZ_MiBi_denik_results.xlsm"

' ## 2. rename both files
Set wb1 = Workbooks("CZ_MiBi_denik_results.xlsm")
Set wb2 = Workbooks("CZ_MiBi_denik_stats.xlsm")

Set Results = wb1.Worksheets("Data")
Set Stats = wb2.Worksheets("Data")


' ## 3. clean up a bit the stats's Data sheet
'remove filters in the old table
If Stats.AutoFilterMode Then Stats.ShowAllData
'delete everything from row 5 downwards
Stats.Rows("5:" & Rows.Count).Delete
' remove all conditional formatting
Stats.Cells.FormatConditions.Delete

' make Results.Data sheet visible
Results.Visible = True


' ## 4. refresh the results (source)
Application.Run ("'CZ_MiBi_denik_results.xlsm'!Combine2")
Application.CutCopyMode = False

' ## 5. copy the results to stats workbook
Results.Range("A3").CurrentRegion.Select
Selection.Offset(1, 0).Select
Selection.Resize(Selection.Rows.Count - 1).Select
Selection.Resize(, Selection.Columns.Count - 1).Select 'to je nova lajna
Selection.Copy
Stats.Range("A3").PasteSpecial Paste:=xlPasteValues

Application.CutCopyMode = False
Results.Visible = False

' make Results.Data sheet hidden

' closes the Results workbook
wb1.Close SaveChanges:=True



' ## 6.Finish the table

' extend the array formulas

LastRow = Stats.Range("A" & Rows.Count).End(xlUp).Row
Stats.Range("L4:AH4").Copy _
Destination:=Stats.Range("L5:L" & LastRow)


' formatting as a number (no decimals)
Stats.Range("D3:K" & LastRow).Select
Selection.NumberFormat = "0"


' add conditional formatting
 With Stats.Range("E3:K" & LastRow).FormatConditions _
    .Add(xlCellValue, xlGreater, "=N3")
  With .Interior
  .Color = 255
  End With
  End With
  
     
' recalculate and turn off auto calculations
Application.Calculation = xlAutomatic
Application.Calculation = xlManual


' refresh pivot table
Sheets("Indexy_podle linek").Select
ActiveSheet.PivotTables("PivotTable10").PivotCache.Refresh

MsgBox ("Data ze souboru Results byla pøekopírována sem.")

End Sub
