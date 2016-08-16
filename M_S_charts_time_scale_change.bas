Attribute VB_Name = "M_S_charts_time_scale_change"
Sub SortByWeek()

Dim PvtTbl As PivotTable
Dim rngGroup As Range
Set PvtTbl = Worksheets("Indexy_podle linek").PivotTables("PivotTable10")

'set range of dates to be grouped
Set rngGroup = PvtTbl.PivotFields("Datum").DataRange

'rngGroup.Cells(1) indicates the first cell in the range of rngGroup - remember that the RangeObject in the Group Method should only be a single cell otherwise the method will fail.
rngGroup.Cells(1).Group By:=7, Periods:=Array(False, False, False, True, False, False, False)

'to ungroup:
'rngGroup.Cells(1).Ungroup

End Sub
Sub SortByMonth()


Dim PvtTbl As PivotTable
Dim rngGroup As Range
Set PvtTbl = Worksheets("Indexy_podle linek").PivotTables("PivotTable10")

'set range of dates to be grouped
Set rngGroup = PvtTbl.PivotFields("Datum").DataRange

'rngGroup.Cells(1) indicates the first cell in the range of rngGroup - remember that the RangeObject in the Group Method should only be a single cell otherwise the method will fail.
rngGroup.Cells(1).Group By:=7, Periods:=Array(False, False, False, False, True, False, True)

'to ungroup:
'rngGroup.Cells(1).Ungroup

End Sub
