Attribute VB_Name = "AmgenFormatReport"
Sub FormatReport(ByVal reportName As String)
Attribute FormatReport.VB_Description = "Formats the reports to the Amgen's standards"
Attribute FormatReport.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AmgenFormatReport Macro
' Formats the reports to the Amgen's standards
'

'
    Dim sourceData As range
    Dim ws As Worksheet
    Dim pivot As PivotTable
    Dim pivotSheet As Worksheet
    Dim range As range
    Dim wb As Workbook
    
    Set wb = Workbooks.Open(Application.GetOpenFilename())
    wb.Activate
    
    Set ws = ActiveSheet
    ' Change Sheet's Name
    If ws.Name <> "Raw Data" Then
        ws.Name = "Raw Data"
    End If
        
    Set sourceData = ws.range("A1").CurrentRegion
    
    If sourceData.Rows.Count < 2 Then
        MsgBox "Looks like the sheet doesn't have enough data"
        Unload AmgenReportsForm
        Exit Sub
    End If
    
    ' Replace unwanted values
    ReplaceUnwantedValues reportName

    ' Convert String Numbers to Numbers
    Call ConvertNumbers
    
    If reportName = "OpenActivities" Then
        Call OpenActivitiesReportHandler
        Set sourceData = ws.range("A1").CurrentRegion
    ElseIf reportName = "OpenSupportRequests" Then
        Call OpenSupportRequestsReportHandler
        Set sourceData = ws.range("A1").CurrentRegion
    End If
        

    If reportName = "TangoeVsAirwatch" Then
        Call TangoeVsAirwatchHandler
        Set sourceData = Sheets("Non Seedstock & Not in AW").range("A1").CurrentRegion
    End If
    
    If reportName = "DEPReport" Then
        Call DEP_Report_Handler
        Set sourceData = ws.range("A1").CurrentRegion
    End If
    
    Cells.EntireColumn.AutoFit
  
    ' Create Pivot Table
    Sheets.Add.Name = "Pivot"
    Set pivotSheet = Sheets("Pivot")
    Utils.CreatePivotTable sourceData, "Pivot!R3C1", "Pivot Table"
    Set pivot = pivotSheet.PivotTables("Pivot Table")
    
    Select Case reportName
        Case "LinesToInactiveUsers"
            Call AmgenFieldsHandler.Amgen_Add_LinesToInactiveUsers_Fields
        Case "LinesWithNoOwner"
            Call AmgenFieldsHandler.Amgen_Add_LinesWithNoOwner_Fields
        Case "ZeroUsageLines"
            Call AmgenFieldsHandler.Amgen_Add_ZeroUsageLines_Fields
        Case "DevicesToInactiveUsers"
            Call AmgenFieldsHandler.Amgen_Add_DevicesToInactiveUsers_Fields
        Case "TangoeVsAirwatch"
            Call AmgenFieldsHandler.Amgen_Add_TangoeVsAirwatch_Fields
        Case "DEPReport"
            Call AmgenFieldsHandler.Amgen_Add_DEPReport_Fields
        Case "OpenActivities"
            Call AmgenFieldsHandler.Amgen_Add_OpenActivities_Field
        Case "OpenSupportRequests"
            Call AmgenFieldsHandler.Amgen_Add_OpenSupportRequests_Field
        Case "SeedstockDevices"
            Call AmgenFieldsHandler.Amgen_Add_SeedstockDevices_Fields
        Case "PendingDestructionDevices"
            Call AmgenFieldsHandler.Amgen_Add_PendingDestructionDevices_Fields
        Case "UsersWithMultipleDevices"
            AmgenFieldsHandler.Amgen_Add_UsersWithMultipleDevices_Fields
        Case Else
            MsgBox "Report Name Not Found, Not adding any fields to table"
    End Select
        
    ' Format Pivot Table
    If reportName <> "OpenActivities" Then
        FormatPivotTable reportName, "Pivot Table"
    End If
    
    If reportName = "UsersWithMultipleDevices" Then
        UsersWithMultipleDevicesHandler
    End If
    
    Unload AmgenReportsForm
    
End Sub

Sub FormatAirwatchVsTangoeReport(ByVal region As String)
'
'
'
    Dim rg As range
    Dim rng As range
    Dim lastCell As Long
    Dim wb As Workbook
    
    Set wb = Workbooks.Open(Application.GetOpenFilename())
    wb.Activate
    
    wb.Sheets("Sheet1").Activate
    
    ActiveSheet.AutoFilterMode = False
    
    Set rg = ActiveSheet.range("A1").CurrentRegion.Rows(1)
    
    ' Custom filters per region
    With rg
        .AutoFilter Field:=111, Criteria1:="#N/D"
        .AutoFilter Field:=25, Criteria1:="Enrolled"
        .AutoFilter Field:=2, Criteria1:="Amgen Corporate"
        .AutoFilter Field:=70, Criteria1:=Array("Consultant", "Staff", "Temp"), Operator:=xlFilterValues
        If region = "LATAM" Then
            .AutoFilter Field:=77, Criteria1:="LATAM"
        ElseIf region = "NA" Then
            .AutoFilter Field:=77, Criteria1:=Array("Canada", "Puerto Rico", "United States"), Operator:=xlFilterValues
        Else
            .AutoFilter Field:=77, Criteria1:="=JAPAC", Operator:=xlOr, Criteria2:="=SG"
        End If
    End With
    
       
    ' Check the number of visible rows
    'On Error Resume Next ' In case there are no visible cells
    'Set rng = range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
    'On Error GoTo 0
    
    'If rng.Rows.Count < 2 Then
    '    MsgBox "Not enough data"
    '    Unload AmgenReportsForm
    '    Exit Sub
    'End If
        
    
    ' Copy the data
    range("J:CD").Copy
    
    ' Create a new workbook
    Workbooks.Add
    ActiveSheet.Name = "Raw Data"
    
    ' Paste the data there
    range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Cells.EntireColumn.AutoFit
   
    ' Delete unwanted columns
    Columns("F:F").Delete Shift:=xlToLeft
    Columns("G:I").Delete Shift:=xlToLeft
    Columns("H:I").Delete Shift:=xlToLeft
    Columns("I:L").Delete Shift:=xlToLeft
    Columns("J:V").Delete Shift:=xlToLeft
    Columns("L:AW").Delete Shift:=xlToLeft
    
    lastCell = Utils.FindLastCellInColumn()

    ' Create and format pivot table
    Sheets.Add.Name = "Pivot"
    
    Utils.CreatePivotTable "Raw Data!R1C1:R" & lastCell & "C12", "Pivot!R3C1", "Pivot Table"

    With ActiveSheet.PivotTables("Pivot Table")
        .AddDataField .PivotFields("Serial Number"), "Count Of Serial Number", xlCount
        .PivotFields("Country (39)").Orientation = xlRowField
        .PivotFields("Display Name").Orientation = xlRowField
        .PivotFields("Device Model").Orientation = xlRowField
        .PivotFields("Serial Number").Orientation = xlRowField
        .PivotFields("Enrollment Date").Orientation = xlRowField
        .PivotFields("Last Seen").Orientation = xlRowField
    End With
    
    FormatPivotTable "other", "Pivot Table"
    
    wb.Close
    
    Unload AmgenReportsForm
    
End Sub

Sub FormatGlobalSeedstockReport()
    With ActiveSheet.PivotTables("Pivot Table").PivotFields("Person")
        .PivotItems("Horizon Verizon Seedstock").Visible = False
        .PivotItems("Horizon AT&T Seedstock").Visible = False
        .PivotItems("Mobility Test Device Seedstock").Visible = False
        .PivotItems("Promotional Seedstock").Visible = False
        .PivotItems("Puerto Rico Seedstock").Visible = False
        .PivotItems("CCXI Seedstock  Sweet").Visible = False
        .PivotItems("Teze iPad Seedstock").Visible = False
        .PivotItems("Shared ADL Operations Seedstock").Visible = False
        .PivotItems("Shared AML Operations Seedstock").Visible = False
        .PivotItems("Shared ARI Operations Seedstock").Visible = False
        .PivotItems("Shared ATO Operations Seedstock").Visible = False
        .PivotItems("Submission and Launch Project Seedstock").Visible = False
        .PivotItems("US EVIP Seedstock").Visible = False
        .PivotItems("US Replacement Seedstock").Visible = False
        .PivotItems("US Sales Seedstock").Visible = False
        .PivotItems("US Used Seedstock").Visible = False
        .PivotItems("USTO Seedstock").Visible = False
    End With
    Unload AmgenReportsForm
End Sub
