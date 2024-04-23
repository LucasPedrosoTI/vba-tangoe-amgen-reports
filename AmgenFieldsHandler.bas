Attribute VB_Name = "AmgenFieldsHandler"
Option Private Module

Sub Amgen_Add_LinesToInactiveUsers_Fields()
    With ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
        .AddDataField ActiveSheet.PivotTables("Pivot Table").PivotFields("International Number"), "Count of International Number", xlCount
        .PivotFields("Carrier").Orientation = xlRowField
        .PivotFields("Owner").Orientation = xlRowField
        .PivotFields("Total Charges Dollar").Orientation = xlRowField
        .PivotFields("Total Data Usage (GBs)").Orientation = xlRowField
        .PivotFields("Total Messaging Usage").Orientation = xlRowField
        .PivotFields("Total Voice Usage").Orientation = xlRowField
        .PivotFields("Airwatch Person").Orientation = xlRowField
        .PivotFields("Owner Inactive At").Orientation = xlRowField
     End With
End Sub

Sub Amgen_Add_LinesWithNoOwner_Fields()
    With ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
        .AddDataField ActiveSheet.PivotTables("Pivot Table").PivotFields("International Number"), "Count of International Number", xlCount
        .PivotFields("Carrier").Orientation = xlRowField
        .PivotFields("Carrier Account").Orientation = xlRowField
        .PivotFields("Total Charges Dollar").Orientation = xlRowField
        .PivotFields("Total Data Usage (GBs)").Orientation = xlRowField
        .PivotFields("Total Messaging Usage").Orientation = xlRowField
        .PivotFields("Total Voice Usage").Orientation = xlRowField
        .PivotFields("Line Created").Orientation = xlRowField
        .PivotFields("Airwatch Person").Orientation = xlRowField
     End With
End Sub

Sub Amgen_Add_DevicesToInactiveUsers_Fields()
    With ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
        .PivotFields("Group").Orientation = xlRowField
        .AddDataField ActiveSheet.PivotTables("Pivot Table").PivotFields("Identifier"), "Count of Identifier", xlCount
        .PivotFields("Identifier").Orientation = xlRowField
        .PivotFields("Airwatch Person").Orientation = xlRowField
        .PivotFields("Airwatch Enrollment Status").Orientation = xlRowField
        .PivotFields("Airwatch Enrollment Date").Orientation = xlRowField
        .PivotFields("Airwatch Last Seen Date").Orientation = xlRowField
     End With
End Sub

Sub Amgen_Add_TangoeVsAirwatch_Fields()
    With ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
        .AddDataField ActiveSheet.PivotTables("Pivot Table").PivotFields("Group"), "Count of Group", xlCount
        .PivotFields("Group").Orientation = xlRowField
        .PivotFields("Person").Orientation = xlRowField
        .PivotFields("Person Active?").Orientation = xlRowField
        .PivotFields("Person Inactive At").Orientation = xlRowField
        .PivotFields("Model").Orientation = xlRowField
        .PivotFields("Device Created").Orientation = xlRowField
     End With
End Sub

Sub Amgen_Add_ZeroUsageLines_Fields()
    With ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
        .PivotFields("Billing Period").Orientation = xlRowField
        .PivotFields("Carrier").Orientation = xlRowField
        .AddDataField .PivotFields("International Number"), "Count Of Number", xlCount
        .PivotFields("International Number").Orientation = xlRowField
        .PivotFields("Total Rebilled Charges").Orientation = xlRowField
        .PivotFields("Person").Orientation = xlRowField
        .PivotFields("Person Active?").Orientation = xlRowField
        .PivotFields("Person Inactive At").Orientation = xlRowField
        .PivotFields("Group").Orientation = xlRowField
     End With
End Sub

Sub Amgen_Add_DEPReport_Fields()
    Dim pivot As PivotTable: Set pivot = ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
    With pivot
        .AddDataField pivot.PivotFields("ID"), "Count of ID", xlCount
        .PivotFields("Group").Orientation = xlRowField
        .PivotFields("Person").Orientation = xlRowField
        .PivotFields("Status").Orientation = xlRowField
        .PivotFields("Device Created").Orientation = xlRowField
        .PivotFields("Model").Orientation = xlRowField
        .PivotFields("IMEI").Orientation = xlRowField
        .PivotFields("Serial Number").Orientation = xlRowField
        .PivotFields("AppleCare Serial Number").Orientation = xlRowField
        .PivotFields("Match in DEP report").Orientation = xlPageField
        .PivotFields("Match in DEP report").ClearAllFilters
        .PivotFields("Match in DEP report").CurrentPage = "#N/A"
        .PivotFields("Match in DEP report 2").Orientation = xlPageField
        .PivotFields("Match in DEP report 2").ClearAllFilters
        .PivotFields("Match in DEP report 2").CurrentPage = "#N/A"
     End With
End Sub

Sub Amgen_Add_OpenActivities_Field()
    Dim pivot As PivotTable
    Set pivot = ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
    With pivot
        .AddDataField pivot.PivotFields("ID"), "Count of ID", xlCount
        .PivotFields("Person Activity For Group").Orientation = xlRowField
    End With
    
    range("A3").CurrentRegion.Copy range("A" & Utils.FindLastCellInColumn + 3)
    
    With Sheets("Pivot").PivotTables(1)
        .PivotFields("Person Activity For Group").Orientation = xlHidden
        .PivotFields("Count of ID").Orientation = xlHidden
        
        .PivotFields("ID").Orientation = xlRowField
        .AddDataField .PivotFields("Days Open"), "Sum of Days Open", xlSum
    End With
End Sub

Sub Amgen_Add_OpenSupportRequests_Field()
    Dim pivot As PivotTable
    Set pivot = ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
    With pivot
        .AddDataField pivot.PivotFields("ID"), "Count of ID", xlCount
        .PivotFields("On Behalf Of Group").Orientation = xlRowField
    End With
    
    range("A3").CurrentRegion.Copy range("A" & Utils.FindLastCellInColumn + 3)
    
    With Sheets("Pivot").PivotTables(1)
        .PivotFields("On Behalf Of Group").Orientation = xlHidden
        .PivotFields("Count of ID").Orientation = xlHidden
        
        .PivotFields("ID").Orientation = xlRowField
        .AddDataField .PivotFields("Days Open"), "Sum of Days Open", xlSum
    End With
End Sub

Sub Amgen_Add_SeedstockDevices_Fields()
    Dim pivot As PivotTable
    Set pivot = ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
    With pivot
        .AddDataField pivot.PivotFields("Person"), "Count of Person", xlCount
        .PivotFields("Person").Orientation = xlRowField
        .PivotFields("Device Category").Orientation = xlColumnField
        .ColumnGrand = False
        .RowGrand = False
    End With
    Call FilterDeviceCategory
End Sub

Sub Amgen_Add_PendingDestructionDevices_Fields()
    Dim pivot As PivotTable
    Set pivot = ActiveWorkbook.Sheets("Pivot").PivotTables("Pivot Table")
    With pivot
        .AddDataField pivot.PivotFields("Person"), "Count of Person", xlCount
        .PivotFields("Person").Orientation = xlRowField
        .PivotFields("Device Category").Orientation = xlColumnField
        .PivotFields("Status").Orientation = xlPageField
        .PivotFields("Status").ClearAllFilters
        .PivotFields("Status").CurrentPage = "Pending Destruction"
    End With
    Call FilterDeviceCategory
    Columns("B:C").EntireColumn.Hidden = True
End Sub


Sub Amgen_Add_UsersWithMultipleDevices_Fields()
    ActiveSheet.Name = "Raw Data Pivot"

    With ActiveSheet.PivotTables("Pivot Table")
        .AddDataField .PivotFields("Device Category"), "Count of Device Category", xlCount
        .PivotFields("Device Category").Orientation = xlColumnField
        .PivotFields("Group").Orientation = xlRowField
        .PivotFields("Person Hr Data Amgen Workforce Login Name").Orientation = xlRowField
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    Call FilterDeviceCategory
End Sub

Private Sub FilterDeviceCategory()
    Dim pivotItem As pivotItem
    For Each pivotItem In ActiveSheet.PivotTables("Pivot Table").PivotFields("Device Category").PivotItems
        Select Case pivotItem.Name
            Case "Data Card", "Phone", "Router", "Signal Booster"
                pivotItem.Visible = False
        End Select
    Next pivotItem
End Sub
