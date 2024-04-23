Attribute VB_Name = "Utils"
Sub ReplaceCommasWithDots()
Attribute ReplaceCommasWithDots.VB_ProcData.VB_Invoke_Func = "R\n14"
    Cells.Replace What:=",", Replacement:=".", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Sub CurrentColumnToText()
    ActiveCell.EntireColumn.TextToColumns destination:=ActiveCell.EntireColumn, DataType:=xlDelimited, FieldInfo:=Array(1, 2), TrailingMinusNumbers:=True
End Sub
Sub ConvertNumbers()
    Dim sourceData As range
    Set sourceData = ActiveSheet.range("A1").CurrentRegion
    For Each cell In sourceData.Cells()
        If IsNumeric(cell.Value) And cell.NumberFormat = "@" And Len(cell.Value) < 15 Then ' "@" is the format for text and 15 is the length of IMEIs
            cell.Value = Val(cell.Value) ' Convert text to number
        End If
    Next cell
End Sub
Sub SetupPivotTable(pivot As PivotTable)
    With pivot
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With pivot.PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
End Sub
Function FindLastCellInColumn() As Variant
    FindLastCellInColumn = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print ("The last row with a value is " & FindLastCellInColumn)
End Function

Sub CreatePivotTable(ByRef srcData As Variant, ByVal destination As String, ByVal pivotTblName As String)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourceData:=srcData, Version:=8).CreatePivotTable TableDestination:=destination, TableName:=pivotTblName, DefaultVersion:=8
End Sub
