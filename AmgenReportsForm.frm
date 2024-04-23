VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AmgenReportsForm 
   Caption         =   "Reports"
   ClientHeight    =   10005
   ClientLeft      =   -150
   ClientTop       =   -510
   ClientWidth     =   9945.001
   OleObjectBlob   =   "AmgenReportsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AmgenReportsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ToggleButton_AirwatchVsTangoe_JAPAC_Click()
    AmgenFormatReport.FormatAirwatchVsTangoeReport "JAPAC"
End Sub

Private Sub ToggleButton_AirwatchVsTangoe_LATAM_Click()
    AmgenFormatReport.FormatAirwatchVsTangoeReport "LATAM"
End Sub

Private Sub ToggleButton_AirwatchVsTangoe_NA_Click()
    AmgenFormatReport.FormatAirwatchVsTangoeReport "NA"
End Sub

Private Sub ToggleButton_Cancel_Click()
    Unload Me
End Sub

Private Sub ToggleButton_FormatGlobalSeedstockDevices_Click()
    AmgenFormatReport.FormatGlobalSeedstockReport
End Sub

Private Sub ToggleButton_LineToInactiveUsers_Click()
    AmgenFormatReport.FormatReport "LinesToInactiveUsers"
End Sub

Private Sub ToggleButton_LinesWithNoOwner_Click()
    AmgenFormatReport.FormatReport "LinesWithNoOwner"
End Sub

Private Sub ToggleButton_OpenActivities_Click()
    AmgenFormatReport.FormatReport "OpenActivities"
End Sub

Private Sub ToggleButton_OpenSupportRequests_Click()
    AmgenFormatReport.FormatReport "OpenSupportRequests"
End Sub

Private Sub ToggleButton_PendingDestructionDevices_Click()
    AmgenFormatReport.FormatReport "PendingDestructionDevices"
End Sub

Private Sub ToggleButton_SeedstockDevices_Click()
    AmgenFormatReport.FormatReport "SeedstockDevices"
End Sub

Private Sub ToggleButton_ZeroUsageLines_Click()
    AmgenFormatReport.FormatReport "ZeroUsageLines"
End Sub

Private Sub ToggleButton_DevicesToInactiveUsers_Click()
    AmgenFormatReport.FormatReport "DevicesToInactiveUsers"
End Sub

Private Sub ToggleButton_TangoeVsAirwatch_Click()
    AmgenFormatReport.FormatReport "TangoeVsAirwatch"
End Sub

Private Sub ToggleButton_DEPReport_Click()
    AmgenFormatReport.FormatReport "DEPReport"
End Sub

Private Sub ToggleButton_UsersWithMultipleDevices_Click()
    AmgenFormatReport.FormatReport "UsersWithMultipleDevices"
End Sub
