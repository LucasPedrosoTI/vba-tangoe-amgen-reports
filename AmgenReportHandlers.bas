Attribute VB_Name = "AmgenReportHandlers"
Option Private Module
Sub DEP_Report_Handler()
    
    Dim depReport As Workbook
    Dim wb As Workbook
    Dim lastCell As Long: lastCell = Utils.FindLastCellInColumn()
    
    Set depReport = Workbooks.Open(Application.GetOpenFilename())
    Debug.Print (depReport.Name)
    
    For Each wb In Workbooks
        If wb.Name <> depReport.Name And wb.Name Like "*In Tangoe Not In DEP*" Then
            wb.Activate
            Exit For
        End If
    Next wb
    
    ' Add columns
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    range("E1").FormulaR1C1 = "Match in DEP report"
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    range("G1").FormulaR1C1 = "Match in DEP report 2"
    
    range("E2").NumberFormat = "General"
    
    range("E2").FormulaR1C1 = "=VLOOKUP(RC[-1],'[" & depReport.Name & "]Sheet1'!C1,1,0)"
    
    range("E2").AutoFill destination:=range("E2:E" & lastCell)
    
    range("E2").Copy range("G2")
    Application.CutCopyMode = False
    
    range("G2").AutoFill destination:=range("G2:G" & lastCell)
    
    range("A1").CurrentRegion.Copy
    range("A1").CurrentRegion.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    With range("A1").CurrentRegion.Rows(1)
        .AutoFilter Field:=5, Criteria1:="#N/D", Operator:=xlAnd
        .AutoFilter Field:=7, Criteria1:="#N/D"
    End With
    
    depReport.Close
End Sub

Sub OpenActivitiesReportHandler()
    Dim lastCell As Long: lastCell = Utils.FindLastCellInColumn()
    
    ActiveSheet.range("J1").Value = "Today"
    ActiveSheet.range("K1").Value = "Days Open"
    
    ActiveSheet.range("J2").FormulaR1C1 = "=TODAY()"
    ActiveSheet.range("K2").FormulaR1C1 = "=DAYS(RC[-1],RC[-3])"
    
    If lastCell > 2 Then
        ActiveSheet.range("J2").AutoFill destination:=range("J2:J" & lastCell)
        ActiveSheet.range("K2").AutoFill destination:=range("K2:K" & lastCell)
    End If
End Sub

Sub OpenSupportRequestsReportHandler()
    Dim lastCell As Long: lastCell = Utils.FindLastCellInColumn()
    
    ActiveSheet.range("M1").Value = "Today"
    ActiveSheet.range("N1").Value = "Days Open"
    
    ActiveSheet.range("M2").FormulaR1C1 = "=TODAY()"
    ActiveSheet.range("N2").FormulaR1C1 = "=DAYS(RC[-1],RC[-3])"
    
    If lastCell > 2 Then
        ActiveSheet.range("M2").AutoFill destination:=range("M2:M" & lastCell)
        ActiveSheet.range("N2").AutoFill destination:=range("N2:N" & lastCell)
    End If
End Sub

Sub TangoeVsAirwatchHandler()
    Dim rg As range
    Dim rawDataName As String: rawDataName = "Raw Data All Devices"
    Dim nonAWName As String: nonAWName = "Non Seedstock & Not in AW"
    ActiveSheet.Name = rawDataName
    Set rg = ActiveSheet.range("A1").CurrentRegion.Rows(1)
    With rg
        .AutoFilter Field:=4, Criteria1:="<>*seedstock*", Operator:=xlAnd
        .AutoFilter Field:=10, Criteria1:="="
    End With
    Sheets.Add.Name = nonAWName
    Sheets(rawDataName).range("A1").CurrentRegion.Copy destination:=Sheets(nonAWName).range("A1")
End Sub

Sub ReturnDeviceActivitiesHandler()
    Dim i As Long
    
    ' Deletes all rows that has Created after 28 days befores
     For i = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If IsDate(ActiveSheet.Cells(i, 6).Value) Then
            If ActiveSheet.Cells(i, 6).Value >= DateAdd("d", -28, Date) Then
                ActiveSheet.Rows(i).Delete
            End If
        End If
    Next i
End Sub

Sub UsersWithMultipleDevicesHandler()
    Dim sourceData As range
    Dim dest As range
    Dim multiDevicesPivot As Worksheet
    Dim lastCell As Long: lastCell = Utils.FindLastCellInColumn()
    Dim pivotTableName As String: pivotTableName = "Multi Device Users Pivot Table"
    Dim sheetName As String: sheetName = "Multi Device Users Pivot"
    
    Sheets.Add.Name = "Comparison"
    Sheets("Raw Data Pivot").range("A4").CurrentRegion.Copy
    
    Sheets("Comparison").range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("Comparison").Activate
    
    ActiveSheet.Rows(1).Delete
    
    Cells.EntireRow.AutoFit
    Cells.EntireColumn.AutoFit
    
    range("A1").CurrentRegion.Rows(1).AutoFilter Field:=2, Criteria1:=""
    
    For i = lastCell To 2 Step -1
        If Not ActiveSheet.Cells(i, 1).EntireRow.Hidden Then
            ActiveSheet.Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
    ActiveSheet.AutoFilterMode = False
      
    range("E1").Value = "Concatanate"
    range("E2").FormulaR1C1 = "=CONCAT(RC[-1],""^"",RC[-2])"
    range("E2").AutoFill destination:=range("E2:E" & Utils.FindLastCellInColumn())

    range("A1").CurrentRegion.Copy
    Sheets.Add.Name = "Multi Device Users"
    range("A1").PasteSpecial xlPasteValues
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
    range("A1").CurrentRegion.Rows(1).AutoFilter Field:=5, Criteria1:=Array("^1", "1^", "1^1"), Operator:=xlFilterValues
    
    For i = Utils.FindLastCellInColumn To 2 Step -1
        If Not ActiveSheet.Cells(i, 1).EntireRow.Hidden Then
            ActiveSheet.Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
    ActiveSheet.AutoFilterMode = False
    
    ' Create Pivot Table
    Sheets.Add.Name = sheetName
    Set multiDevicesPivot = Sheets(sheetName)
    
    Set sourceData = Sheets("Multi Device Users").range("A1").CurrentRegion
    Set dest = multiDevicesPivot.range("A3")
    dest.Activate
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourceData:=sourceData, Version:=8).CreatePivotTable TableDestination:=dest, TableName:=pivotTableName, DefaultVersion:=8
    
    With ActiveSheet.PivotTables(pivotTableName)
        .AddDataField .PivotFields("Group"), "Count of Group", xlCount
        .PivotFields("Group").Orientation = xlRowField
        .PivotFields("Person Hr Data Amgen Workforce Login Name").Orientation = xlRowField
    End With
    
    FormatPivotTable "Other", pivotTableName
    
End Sub
