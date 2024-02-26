Attribute VB_Name = "AmgenUtils"
Option Private Module

Sub ReplaceUnwantedValues(ByVal reportName As String)
    
    ActiveSheet.range("A1").currentRegion.Select
    
    Selection.Replace What:="-> ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="UEM Device", Replacement:="Airwatch", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="$", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EST", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="edt", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="Line Bill ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="Account Number", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:=" +00:00", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

        
    For Each cell In Selection.Rows(1).Cells
        If cell.value = "Total Charges" Then
           cell.value = "Total Charges Dollar"
        ElseIf cell.value = "Active?" Then
            cell.value = "Person Active?"
        ElseIf cell.value = "Inactive At" Then
            cell.value = "Person Inactive At"
        ElseIf cell.value = "Created" Then
            If reportName = "SeedstockDevices" Or reportName = "PendingDestructionDevices" Or reportName = "TangoeVsAirwatch" Or reportName = "DEPReport" Then
                cell.value = "Device Created"
            Else
                cell.value = "Line Created"
            End If
        End If
        cell.value = Trim(cell.value)
    Next cell
End Sub

Sub FormatPivotTable(ByVal reportName As String, ByVal pivotTableName As String)
'
' AmgenFormatPivotTable Macro
' Formats the pivot table in Amgen's standards
'
'
    Dim pivot As PivotTable
    Dim pivotSheet As Worksheet
    
    Set pivot = ActiveSheet.PivotTables(pivotTableName)
    
    ' Show in Tabular form
    pivot.RowAxisLayout xlTabularRow
    
    ' Do not show subtotals
    For Each Field In pivot.PivotFields()
        Field.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next Field
    
    Call Amgen_FormatDate
    
    Cells.EntireRow.AutoFit
    Cells.EntireColumn.AutoFit
    
    If reportName = "UsersWithMultipleDevices" Then
        pivot.RepeatAllLabels xlRepeatLabels
    Else
        pivot.PivotFields(1).ShowDetail = False
    End If
End Sub

Sub Amgen_FormatDate()
'
' AmgenFormatDate Macro
' format the column as date (d/mmm/yy)
'
'
    Dim rng As range
    Dim cell As range
    Dim col As range

    Set rng = ActiveSheet.UsedRange

    Application.ScreenUpdating = False

    For Each col In rng.Columns
        Set cell = col.Cells(4, 1)
        ' Check if the cell is formatted as a date or can be interpreted as a date
        If IsDate(cell.value) Then
            col.EntireColumn.NumberFormat = "d-mmm-yy"
        End If
    Next col

    Application.ScreenUpdating = True
End Sub
