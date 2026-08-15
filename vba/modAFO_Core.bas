Attribute VB_Name = "modAFO_Core"
Option Explicit

' =============================================================================
' COPY-READY SAP ANALYSIS FOR OFFICE VBA
' =============================================================================
' These are the small macros to copy into another AFO workbook.

Public Sub Refresh_AFO()
    Dim lResult As Long
    Dim dataSourceAlias As String
    Dim sapClient As String
    Dim userName As String
    Dim password As String

    dataSourceAlias = "DS_1"
    sapClient = "500"
    userName = "SAP_USERNAME"
    password = "SAP_PASSWORD"

    lResult = Application.Run("SAPLogon", dataSourceAlias, sapClient, userName, password)
    DoEvents

    If lResult <> 1 Then
        MsgBox "SAPLogon failed for " & dataSourceAlias & ".", vbCritical
        Exit Sub
    End If

    lResult = Application.Run("SAPExecuteCommand", "RefreshData")
    DoEvents

    If lResult <> 1 Then
        MsgBox "SAP Analysis for Office refresh failed.", vbCritical
    End If
End Sub

Public Sub Set_AFO_Filter()
    Dim lResult As Long
    Dim dataSourceAlias As String
    Dim variableName As String
    Dim promptValue As String

    dataSourceAlias = "DS_1"
    variableName = "P_COMP_CODE"
    promptValue = "1000"

    lResult = Application.Run("SAPSetFilter", dataSourceAlias, variableName, _
                              promptValue, "INPUT_STRING")
    DoEvents

    If lResult <> 1 Then
        MsgBox "SAPSetFilter failed for " & variableName & ".", vbCritical
    End If
End Sub

' Useful AFO worksheet formula:
' =@SAPGetSourceInfo("DS_1", "DataSourceName")
