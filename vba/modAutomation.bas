Attribute VB_Name = "modAutomation"
Option Explicit

' =============================================================================
' SAP AFO AUTOMATION WORKFLOW
' =============================================================================
' Start with RefreshSelectedReport. It shows the complete workflow in order:
' read Configuration, log on, set prompts, refresh, update the queue and log it.
'
' Table values are always read by their visible column names. There are no hidden
' column-number constants to look up.

' -----------------------------------------------------------------------------
' RUN SELECTED
' -----------------------------------------------------------------------------

Public Sub RefreshSelectedReport()
    Dim wsSummary As Worksheet
    Dim wsConfiguration As Worksheet
    Dim wsReport As Worksheet
    Dim tblReports As ListObject
    Dim tblPrompts As ListObject
    Dim reportNumber As Long
    Dim selectedReportNumber As Long
    Dim promptNumber As Long
    Dim selectedReportID As String
    Dim reportID As String
    Dim reportName As String
    Dim reportSheetName As String
    Dim runMode As String
    Dim promptReportID As String
    Dim promptDataSource As String
    Dim afoVariableName As String
    Dim promptRule As String
    Dim promptValue As String
    Dim startedAt As Single
    Dim durationSeconds As Double
    Dim resultStatus As String
    Dim resultDetail As String
    Dim errorText As String
    Dim oldEvents As Boolean

    On Error GoTo Failed
    startedAt = Timer
    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    Application.StatusBar = "Refreshing the selected AFO report"

    Set wsSummary = ThisWorkbook.Worksheets("Summary")
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    Set tblReports = wsConfiguration.ListObjects("tblReports")
    Set tblPrompts = wsConfiguration.ListObjects("tblPrompts")

    ' 1. Find the report selected in cell P5 on Summary.
    selectedReportID = wsSummary.Range("P5").Value

    For reportNumber = 1 To tblReports.ListRows.Count
        If tblReports.ListColumns("Report ID").DataBodyRange.Cells(reportNumber, 1).Value = selectedReportID Then
            selectedReportNumber = reportNumber
            Exit For
        End If
    Next reportNumber

    If selectedReportNumber = 0 Then
        Err.Raise vbObjectError + 710, , "The selected report was not found in tblReports."
    End If

    Dim dataSourceAlias As String

    reportID = tblReports.ListColumns("Report ID").DataBodyRange.Cells(selectedReportNumber, 1).Value
    reportName = tblReports.ListColumns("Report Name").DataBodyRange.Cells(selectedReportNumber, 1).Value
    dataSourceAlias = tblReports.ListColumns("Data Source").DataBodyRange.Cells(selectedReportNumber, 1).Value
    reportSheetName = tblReports.ListColumns("Target Sheet").DataBodyRange.Cells(selectedReportNumber, 1).Value
    runMode = wsConfiguration.Range("C11").Value

    If runMode <> "SIMULATION" And runMode <> "LIVE" Then
        Err.Raise vbObjectError + 711, , "Mode must be SIMULATION or LIVE."
    End If
    If dataSourceAlias = "" Then
        Err.Raise vbObjectError + 712, , "The selected report has no data-source alias."
    End If
    If reportSheetName = "" Then
        Err.Raise vbObjectError + 713, , "The selected report has no target sheet."
    End If

    tblReports.ListColumns("Status").DataBodyRange.Cells(selectedReportNumber, 1).Value = "Running"
    tblReports.ListColumns("Progress").DataBodyRange.Cells(selectedReportNumber, 1).Value = 0.15
    UpdateSummaryStatus "RUNNING", reportName, 0.15, "Preparing prompts"

    ' 2. LIVE only: log on to the AFO data source selected above.
    If runMode = "LIVE" Then
        If wsConfiguration.Range("C12").Value = "" Then
            Err.Raise vbObjectError + 714, , "SAP Client is blank."
        End If
        If wsConfiguration.Range("C16").Value = "" Or _
           wsConfiguration.Range("C17").Value = "" Or _
           wsConfiguration.Range("C16").Value = "SAP_USERNAME" Or _
           wsConfiguration.Range("C17").Value = "SAP_PASSWORD" Then
            Err.Raise vbObjectError + 715, , _
                "Enter the SAP Username and SAP Password on Configuration."
        End If

        Dim lResult As Long
        Dim sapClient As String
        Dim userName As String
        Dim password As String

        dataSourceAlias = tblReports.ListColumns("Data Source").DataBodyRange.Cells(selectedReportNumber, 1).Value
        sapClient = wsConfiguration.Range("C12").Value
        userName = wsConfiguration.Range("C16").Value
        password = wsConfiguration.Range("C17").Value

        ' REAL AFO CALL 1 OF 3.
        lResult = Application.Run("SAPLogon", dataSourceAlias, sapClient, userName, password)

        If lResult <> 1 Then
            Err.Raise vbObjectError + 716, , "SAPLogon failed for " & dataSourceAlias & "."
        End If

        userName = ""
        password = ""
    End If

    ' 3. Resolve each active prompt for this report.
    For promptNumber = 1 To tblPrompts.ListRows.Count
        If tblPrompts.ListColumns("Active").DataBodyRange.Cells(promptNumber, 1).Value = True Then
            promptReportID = tblPrompts.ListColumns("Report ID").DataBodyRange.Cells(promptNumber, 1).Value

            If promptReportID = "*" Or promptReportID = reportID Then
                promptRule = tblPrompts.ListColumns("Rule").DataBodyRange.Cells(promptNumber, 1).Value

                Select Case promptRule
                    Case "VALUE"
                        promptValue = tblPrompts.ListColumns("Value").DataBodyRange.Cells(promptNumber, 1).Text
                    Case "LAST_13_WEEKS"
                        promptValue = PreviousCompleteWeeksDaily(13)
                    Case "WEEK_STARTS"
                        promptValue = PreviousWeekWorkingDays(13)
                    Case "FY_DAILY"
                        promptValue = CurrentFiscalYearDates(Date, 4)
                    Case "FY_MONTH_STARTS"
                        promptValue = FiscalMonthFirstWorkingDays(Date, 4)
                    Case Else
                        Err.Raise vbObjectError + 717, , "Unknown prompt rule: " & promptRule
                End Select

                afoVariableName = tblPrompts.ListColumns("AFO Variable").DataBodyRange.Cells(promptNumber, 1).Value
                promptDataSource = tblPrompts.ListColumns("Data Source").DataBodyRange.Cells(promptNumber, 1).Value
                If promptDataSource = "" Then promptDataSource = dataSourceAlias

                If tblPrompts.ListColumns("Required").DataBodyRange.Cells(promptNumber, 1).Value = True _
                   And promptValue = "" Then
                    Err.Raise vbObjectError + 718, , _
                        "Required prompt is blank: " & _
                        tblPrompts.ListColumns("Parameter").DataBodyRange.Cells(promptNumber, 1).Value
                End If

                ' Text format preserves values such as fiscal period 005.
                tblPrompts.ListColumns("Preview").DataBodyRange.Cells(promptNumber, 1).NumberFormat = "@"
                If Len(promptValue) > 70 Then
                    tblPrompts.ListColumns("Preview").DataBodyRange.Cells(promptNumber, 1).Value = _
                        Left(promptValue, 67) & "..."
                Else
                    tblPrompts.ListColumns("Preview").DataBodyRange.Cells(promptNumber, 1).Value = promptValue
                End If

                ' REAL AFO CALL 2 OF 3: set the technical variable.
                If runMode = "LIVE" And promptValue <> "" Then
                    lResult = Application.Run("SAPSetFilter", promptDataSource, _
                                              afoVariableName, promptValue, "INPUT_STRING")

                    If lResult <> 1 Then
                        Err.Raise vbObjectError + 719, , _
                            "SAPSetFilter failed for " & afoVariableName & "."
                    End If
                End If
            End If
        End If
    Next promptNumber

    ' 4. Refresh AFO. SIMULATION replaces only the unavailable SAP call.
    UpdateSummaryStatus "RUNNING", reportName, 0.6, "Refreshing " & dataSourceAlias

    If runMode = "LIVE" Then
        ' REAL AFO CALL 3 OF 3: refresh the selected data source.
        lResult = Application.Run("SAPExecuteCommand", "RefreshData", dataSourceAlias)

        If lResult <> 1 Then
            Err.Raise vbObjectError + 720, , "RefreshData failed for " & dataSourceAlias & "."
        End If

        resultStatus = "Success"
        resultDetail = "AFO refresh completed"
    Else
        resultStatus = tblReports.ListColumns("Simulation Result").DataBodyRange.Cells(selectedReportNumber, 1).Value
        resultDetail = "Refresh completed"

        If resultStatus = "Warning" Then resultDetail = "Prompt returned no rows"
        If resultStatus = "Failed" Then
            Err.Raise vbObjectError + 721, , "Simulated AFO data-source failure"
        End If

        Set wsReport = ThisWorkbook.Worksheets(reportSheetName)
        wsReport.Range("B2").Value = UCase(reportName)
        wsReport.Range("K3").Value = dataSourceAlias
        wsReport.Range("B31").Value = "Generated " & Format(Now, "dd-mmm-yyyy hh:mm:ss")
    End If

    ' 5. Update the report queue and audit history.
    durationSeconds = Timer - startedAt
    If durationSeconds < 0 Then durationSeconds = durationSeconds + 86400

    tblReports.ListColumns("Status").DataBodyRange.Cells(selectedReportNumber, 1).Value = resultStatus
    tblReports.ListColumns("Progress").DataBodyRange.Cells(selectedReportNumber, 1).Value = 1
    tblReports.ListColumns("Last Run").DataBodyRange.Cells(selectedReportNumber, 1).Value = Now
    tblReports.ListColumns("Last Run").DataBodyRange.Cells(selectedReportNumber, 1).NumberFormat = "dd-mmm hh:mm"
    tblReports.ListColumns("Result Detail").DataBodyRange.Cells(selectedReportNumber, 1).Value = resultDetail

    AddRunHistoryEntry reportID, reportName, "Refresh", resultStatus, _
                       durationSeconds, "", resultDetail
    UpdateSummaryStatus UCase(resultStatus), reportName, 1, resultDetail

CleanExit:
    userName = ""
    password = ""
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    CopyLatestHistoryToSummary
    Exit Sub

Failed:
    errorText = Err.Description
    durationSeconds = Timer - startedAt
    If durationSeconds < 0 Then durationSeconds = durationSeconds + 86400

    On Error Resume Next
    If selectedReportNumber > 0 Then
        tblReports.ListColumns("Status").DataBodyRange.Cells(selectedReportNumber, 1).Value = "Failed"
        tblReports.ListColumns("Progress").DataBodyRange.Cells(selectedReportNumber, 1).Value = 1
        tblReports.ListColumns("Last Run").DataBodyRange.Cells(selectedReportNumber, 1).Value = Now
        tblReports.ListColumns("Last Run").DataBodyRange.Cells(selectedReportNumber, 1).NumberFormat = "dd-mmm hh:mm"
        tblReports.ListColumns("Result Detail").DataBodyRange.Cells(selectedReportNumber, 1).Value = errorText
    End If
    AddRunHistoryEntry reportID, reportName, "Refresh", "Failed", _
                       durationSeconds, "", errorText
    UpdateSummaryStatus "FAILED", reportName, 0, errorText
    On Error GoTo 0
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------
' RUN ACTIVE BATCH
' -----------------------------------------------------------------------------

Public Sub RunActiveBatch()
    Dim wsSummary As Worksheet
    Dim wsConfiguration As Worksheet
    Dim tblReports As ListObject
    Dim reportNumber As Long
    Dim originalSelectedReport As String
    Dim reportName As String
    Dim resultStatus As String
    Dim continueOnError As Boolean
    Dim totalReports As Long
    Dim completedReports As Long
    Dim successfulReports As Long
    Dim warningReports As Long
    Dim failedReports As Long
    Dim startedAt As Single
    Dim durationSeconds As Double
    Dim batchStatus As String
    Dim batchDetail As String
    Dim errorText As String

    On Error GoTo BatchFailed
    startedAt = Timer

    Set wsSummary = ThisWorkbook.Worksheets("Summary")
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    Set tblReports = wsConfiguration.ListObjects("tblReports")

    originalSelectedReport = wsSummary.Range("P5").Value
    continueOnError = wsConfiguration.Range("C15").Value

    ValidateTemplate
    If wsSummary.Range("G10").Value = "NEEDS SETUP" Then Exit Sub

    For reportNumber = 1 To tblReports.ListRows.Count
        If tblReports.ListColumns("Active").DataBodyRange.Cells(reportNumber, 1).Value = True Then
            totalReports = totalReports + 1
        End If
    Next reportNumber

    If totalReports = 0 Then
        Err.Raise vbObjectError + 730, , "There are no active reports."
    End If

    UpdateSummaryStatus "RUNNING", "Preparing batch", 0, totalReports & " active reports"

    ' Reuse the same two workflows as the Summary buttons.
    For reportNumber = 1 To tblReports.ListRows.Count
        If tblReports.ListColumns("Active").DataBodyRange.Cells(reportNumber, 1).Value = True Then
            wsSummary.Range("P5").Value = _
                tblReports.ListColumns("Report ID").DataBodyRange.Cells(reportNumber, 1).Value
            reportName = tblReports.ListColumns("Report Name").DataBodyRange.Cells(reportNumber, 1).Value

            RefreshSelectedReport
            resultStatus = tblReports.ListColumns("Status").DataBodyRange.Cells(reportNumber, 1).Value

            If resultStatus <> "Failed" Then
                ExportSelectedReport
                resultStatus = tblReports.ListColumns("Status").DataBodyRange.Cells(reportNumber, 1).Value
            End If

            completedReports = completedReports + 1
            Select Case resultStatus
                Case "Success"
                    successfulReports = successfulReports + 1
                Case "Warning"
                    warningReports = warningReports + 1
                Case Else
                    failedReports = failedReports + 1
            End Select

            UpdateSummaryStatus "RUNNING", reportName, completedReports / totalReports, _
                                completedReports & " of " & totalReports & " completed"

            If resultStatus = "Failed" And Not continueOnError Then Exit For
        End If
    Next reportNumber

    wsSummary.Range("P5").Value = originalSelectedReport
    durationSeconds = Timer - startedAt
    If durationSeconds < 0 Then durationSeconds = durationSeconds + 86400

    If failedReports > 0 Or warningReports > 0 Then
        batchStatus = "Warning"
    Else
        batchStatus = "Success"
    End If

    batchDetail = successfulReports & " successful / " & _
                  warningReports & " warning / " & failedReports & " failed"

    AddRunHistoryEntry "", "", "Batch", batchStatus, durationSeconds, "", batchDetail
    UpdateSummaryStatus "COMPLETED", "Active batch finished", 1, batchDetail
    CopyLatestHistoryToSummary
    Exit Sub

BatchFailed:
    errorText = Err.Description
    On Error Resume Next
    wsSummary.Range("P5").Value = originalSelectedReport
    AddRunHistoryEntry "", "", "Batch", "Failed", 0, "", errorText
    UpdateSummaryStatus "FAILED", "Batch stopped", 0, errorText
    CopyLatestHistoryToSummary
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' VALIDATE TEMPLATE
' -----------------------------------------------------------------------------

Public Sub ValidateTemplate()
    Dim wsSummary As Worksheet
    Dim wsConfiguration As Worksheet
    Dim wsReport As Worksheet
    Dim tblReports As ListObject
    Dim tblPrompts As ListObject
    Dim reportNumber As Long
    Dim promptNumber As Long
    Dim reportID As String
    Dim reportSheetName As String
    Dim selectedReportID As String
    Dim runMode As String
    Dim exportFolder As String
    Dim activeReports As Long
    Dim selectedReportFound As Boolean
    Dim issues As String

    On Error GoTo BrokenWorkbook

    Set wsSummary = ThisWorkbook.Worksheets("Summary")
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    Set tblReports = wsConfiguration.ListObjects("tblReports")
    Set tblPrompts = wsConfiguration.ListObjects("tblPrompts")

    runMode = wsConfiguration.Range("C11").Value
    exportFolder = wsConfiguration.Range("C14").Value
    selectedReportID = wsSummary.Range("P5").Value

    If runMode <> "SIMULATION" And runMode <> "LIVE" Then
        issues = issues & "- Mode must be SIMULATION or LIVE" & vbLf
    End If
    If wsConfiguration.Range("C12").Value = "" Then
        issues = issues & "- SAP Client is blank" & vbLf
    End If
    If exportFolder = "" Or InStr(exportFolder, ":") > 0 Or _
       Left(exportFolder, 1) = "\" Or Left(exportFolder, 1) = "/" Or _
       InStr(exportFolder, "..") > 0 Then
        issues = issues & "- Export Folder must be a relative folder" & vbLf
    End If
    If ThisWorkbook.Path = "" Then
        issues = issues & "- Save the workbook before exporting" & vbLf
    End If

    For reportNumber = 1 To tblReports.ListRows.Count
        reportID = tblReports.ListColumns("Report ID").DataBodyRange.Cells(reportNumber, 1).Value

        If reportID = selectedReportID Then selectedReportFound = True

        If tblReports.ListColumns("Active").DataBodyRange.Cells(reportNumber, 1).Value = True Then
            activeReports = activeReports + 1

            If reportID = "" Then
                issues = issues & "- An active report has no Report ID" & vbLf
            End If
            If tblReports.ListColumns("Data Source").DataBodyRange.Cells(reportNumber, 1).Value = "" Then
                issues = issues & "- " & reportID & " has no data-source alias" & vbLf
            End If

            reportSheetName = tblReports.ListColumns("Target Sheet").DataBodyRange.Cells(reportNumber, 1).Value
            If reportSheetName = "" Then
                issues = issues & "- " & reportID & " has no target sheet" & vbLf
            Else
                Set wsReport = Nothing
                On Error Resume Next
                Set wsReport = ThisWorkbook.Worksheets(reportSheetName)
                On Error GoTo BrokenWorkbook

                If wsReport Is Nothing Then
                    issues = issues & "- " & reportID & " target sheet was not found" & vbLf
                End If
            End If
        End If
    Next reportNumber

    If activeReports = 0 Then issues = issues & "- There are no active reports" & vbLf
    If Not selectedReportFound Then issues = issues & "- The selected report was not found" & vbLf

    For promptNumber = 1 To tblPrompts.ListRows.Count
        If tblPrompts.ListColumns("Active").DataBodyRange.Cells(promptNumber, 1).Value = True Then
            If tblPrompts.ListColumns("Report ID").DataBodyRange.Cells(promptNumber, 1).Value = "" Then
                issues = issues & "- An active prompt has no Report ID" & vbLf
            End If
            If tblPrompts.ListColumns("AFO Variable").DataBodyRange.Cells(promptNumber, 1).Value = "" Then
                issues = issues & "- An active prompt has no AFO variable" & vbLf
            End If
            If tblPrompts.ListColumns("Required").DataBodyRange.Cells(promptNumber, 1).Value = True _
               And tblPrompts.ListColumns("Rule").DataBodyRange.Cells(promptNumber, 1).Value = "VALUE" _
               And tblPrompts.ListColumns("Value").DataBodyRange.Cells(promptNumber, 1).Text = "" Then
                issues = issues & "- A required VALUE prompt is blank" & vbLf
            End If
        End If
    Next promptNumber

    If runMode = "LIVE" Then
        If wsConfiguration.Range("C16").Value = "" Or _
           wsConfiguration.Range("C17").Value = "" Or _
           wsConfiguration.Range("C16").Value = "SAP_USERNAME" Or _
           wsConfiguration.Range("C17").Value = "SAP_PASSWORD" Then
            issues = issues & "- Enter the SAP Username and SAP Password" & vbLf
        End If
    End If

PublishResult:
    If issues = "" Then
        UpdateSummaryStatus "READY", "Configuration validated", 0, _
                            "Aliases, variables and relative output settings are valid"
        AddRunHistoryEntry "", "", "Validate", "Success", 0, "", "Configuration passed"
    Else
        UpdateSummaryStatus "NEEDS SETUP", "Configuration requires attention", 0, issues
        AddRunHistoryEntry "", "", "Validate", "Failed", 0, "", Replace(issues, vbLf, " | ")
    End If

    CopyLatestHistoryToSummary
    Exit Sub

BrokenWorkbook:
    issues = issues & "- Workbook structure error: " & Err.Description & vbLf
    Resume PublishResult
End Sub

' -----------------------------------------------------------------------------
' EXPORT SELECTED
' -----------------------------------------------------------------------------

Public Sub ExportSelectedReport()
    Dim wsSummary As Worksheet
    Dim wsConfiguration As Worksheet
    Dim wsReport As Worksheet
    Dim tblReports As ListObject
    Dim reportNumber As Long
    Dim selectedReportNumber As Long
    Dim selectedReportID As String
    Dim reportID As String
    Dim reportName As String
    Dim reportSheetName As String
    Dim exportFolder As String
    Dim exportPath As String
    Dim baseFileName As String
    Dim outputText As String
    Dim errorText As String
    Dim exportFormats As Variant
    Dim exportFormat As Variant
    Dim folderPart As Variant
    Dim invalidCharacter As Variant
    Dim exportWorkbook As Workbook
    Dim oldDisplayAlerts As Boolean

    On Error GoTo Failed
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.StatusBar = "Exporting the selected report"

    Set wsSummary = ThisWorkbook.Worksheets("Summary")
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    Set tblReports = wsConfiguration.ListObjects("tblReports")

    selectedReportID = wsSummary.Range("P5").Value

    For reportNumber = 1 To tblReports.ListRows.Count
        If tblReports.ListColumns("Report ID").DataBodyRange.Cells(reportNumber, 1).Value = selectedReportID Then
            selectedReportNumber = reportNumber
            Exit For
        End If
    Next reportNumber

    If selectedReportNumber = 0 Then
        Err.Raise vbObjectError + 740, , "The selected report was not found."
    End If

    reportID = tblReports.ListColumns("Report ID").DataBodyRange.Cells(selectedReportNumber, 1).Value
    reportName = tblReports.ListColumns("Report Name").DataBodyRange.Cells(selectedReportNumber, 1).Value
    reportSheetName = tblReports.ListColumns("Target Sheet").DataBodyRange.Cells(selectedReportNumber, 1).Value
    Set wsReport = ThisWorkbook.Worksheets(reportSheetName)

    ' Build the output path from the workbook folder and the relative setting.
    exportFolder = wsConfiguration.Range("C14").Value
    If exportFolder = "" Or InStr(exportFolder, ":") > 0 Or _
       Left(exportFolder, 1) = "\" Or Left(exportFolder, 1) = "/" Or _
       InStr(exportFolder, "..") > 0 Then
        Err.Raise vbObjectError + 741, , "Export Folder must be relative to this workbook."
    End If
    If ThisWorkbook.Path = "" Then
        Err.Raise vbObjectError + 742, , "Save the workbook before exporting."
    End If

    exportPath = ThisWorkbook.Path
    For Each folderPart In Split(Replace(exportFolder, "/", Application.PathSeparator), Application.PathSeparator)
        If folderPart <> "" Then
            exportPath = exportPath & Application.PathSeparator & folderPart
            If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
        End If
    Next folderPart

    baseFileName = reportID & "_" & Format(Now, "yyyy-mm-dd_HHmmss")
    For Each invalidCharacter In Array("<", ">", ":", Chr(34), "/", "\", "|", "?", "*")
        baseFileName = Replace(baseFileName, invalidCharacter, "-")
    Next invalidCharacter

    exportFormats = Split( _
        tblReports.ListColumns("Export Formats").DataBodyRange.Cells(selectedReportNumber, 1).Value, _
        ",")

    For Each exportFormat In exportFormats
        Select Case exportFormat
            Case "PDF"
                wsReport.ExportAsFixedFormat xlTypePDF, _
                    exportPath & Application.PathSeparator & baseFileName & ".pdf", _
                    xlQualityStandard, True, False
                If outputText <> "" Then outputText = outputText & " / "
                outputText = outputText & "PDF"

            Case "XLSX", "CSV"
                wsReport.Copy
                Set exportWorkbook = ActiveWorkbook

                If exportFormat = "XLSX" Then
                    exportWorkbook.SaveAs _
                        exportPath & Application.PathSeparator & baseFileName & ".xlsx", _
                        xlOpenXMLWorkbook
                Else
                    exportWorkbook.SaveAs _
                        exportPath & Application.PathSeparator & baseFileName & ".csv", _
                        xlCSVUTF8
                End If

                exportWorkbook.Close False
                Set exportWorkbook = Nothing
                If outputText <> "" Then outputText = outputText & " / "
                outputText = outputText & exportFormat
        End Select
    Next exportFormat

    If outputText = "" Then
        Err.Raise vbObjectError + 743, , "No supported export format was configured."
    End If

    AddRunHistoryEntry reportID, reportName, "Export", "Success", 0, outputText, "Export completed"
    UpdateSummaryStatus "EXPORTED", reportName, 1, _
                        outputText & " created in the relative export folder"

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.StatusBar = False
    CopyLatestHistoryToSummary
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not exportWorkbook Is Nothing Then exportWorkbook.Close False
    If selectedReportNumber > 0 Then
        tblReports.ListColumns("Status").DataBodyRange.Cells(selectedReportNumber, 1).Value = "Failed"
        tblReports.ListColumns("Result Detail").DataBodyRange.Cells(selectedReportNumber, 1).Value = _
            "Export failed: " & errorText
    End If
    AddRunHistoryEntry reportID, reportName, "Export", "Failed", 0, "", errorText
    UpdateSummaryStatus "FAILED", reportName, 0, errorText
    On Error GoTo 0
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------
' SMALL BUTTONS
' -----------------------------------------------------------------------------

Public Sub OpenExportFolder()
    Dim wsConfiguration As Worksheet
    Dim exportFolder As String
    Dim exportPath As String
    Dim folderPart As Variant

    On Error GoTo Failed

    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    exportFolder = wsConfiguration.Range("C14").Value

    If exportFolder = "" Or InStr(exportFolder, ":") > 0 Or _
       Left(exportFolder, 1) = "\" Or Left(exportFolder, 1) = "/" Or _
       InStr(exportFolder, "..") > 0 Then
        Err.Raise vbObjectError + 750, , "Export Folder must be relative to this workbook."
    End If
    If ThisWorkbook.Path = "" Then
        Err.Raise vbObjectError + 751, , "Save the workbook first."
    End If

    exportPath = ThisWorkbook.Path
    For Each folderPart In Split(Replace(exportFolder, "/", Application.PathSeparator), Application.PathSeparator)
        If folderPart <> "" Then
            exportPath = exportPath & Application.PathSeparator & folderPart
            If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
        End If
    Next folderPart

    Shell "explorer.exe " & Chr(34) & exportPath & Chr(34), vbNormalFocus
    Exit Sub

Failed:
    UpdateSummaryStatus "FAILED", "", 0, Err.Description
End Sub

Public Sub ResetWorkspace()
    Dim wsSummary As Worksheet
    Dim wsConfiguration As Worksheet
    Dim tblReports As ListObject
    Dim tblRunHistory As ListObject
    Dim reportNumber As Long
    Dim oldEvents As Boolean
    Dim errorText As String

    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Application.EnableEvents = False

    Set wsSummary = ThisWorkbook.Worksheets("Summary")
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    Set tblReports = wsConfiguration.ListObjects("tblReports")
    Set tblRunHistory = ThisWorkbook.Worksheets("Run History").ListObjects("tblRunHistory")

    wsConfiguration.Range("C11").Value = "SIMULATION"
    wsSummary.Range("P5").Value = "RPT001"

    For reportNumber = 1 To tblReports.ListRows.Count
        If tblReports.ListColumns("Active").DataBodyRange.Cells(reportNumber, 1).Value = True Then
            tblReports.ListColumns("Status").DataBodyRange.Cells(reportNumber, 1).Value = "Ready"
            tblReports.ListColumns("Result Detail").DataBodyRange.Cells(reportNumber, 1).Value = "Not run"
        Else
            tblReports.ListColumns("Status").DataBodyRange.Cells(reportNumber, 1).Value = "Not Run"
            tblReports.ListColumns("Result Detail").DataBodyRange.Cells(reportNumber, 1).Value = "Inactive"
        End If

        tblReports.ListColumns("Progress").DataBodyRange.Cells(reportNumber, 1).Value = 0
        tblReports.ListColumns("Last Run").DataBodyRange.Cells(reportNumber, 1).ClearContents
    Next reportNumber

    If Not tblRunHistory.DataBodyRange Is Nothing Then tblRunHistory.DataBodyRange.Delete

    AddRunHistoryEntry "", "", "Reset", "Success", 0, "", "Workspace reset"
    UpdateSummaryStatus "READY", "No batch is running", 0, _
                        "Choose a report or run the active batch"

CleanExit:
    Application.EnableEvents = oldEvents
    CopyLatestHistoryToSummary
    Exit Sub

Failed:
    errorText = Err.Description
    UpdateSummaryStatus "FAILED", "", 0, errorText
    Resume CleanExit
End Sub

Public Sub OpenCodeLibrary()
    ThisWorkbook.Worksheets("AFO Code Library").Activate
    ActiveWindow.ScrollRow = 1
End Sub

Public Sub InitialiseTemplate()
    Dim wsSummary As Worksheet
    Dim wsConfiguration As Worksheet

    On Error Resume Next
    Set wsSummary = ThisWorkbook.Worksheets("Summary")
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")

    wsSummary.Activate
    wsSummary.Range("N3").Value = wsConfiguration.Range("C11").Value
    CopyLatestHistoryToSummary
    ActiveWindow.Zoom = 80
    On Error GoTo 0
End Sub

' =============================================================================
' THREE REPEATED WORKBOOK TASKS
' =============================================================================

Private Sub AddRunHistoryEntry(ByVal reportID As String, ByVal reportName As String, _
                               ByVal actionText As String, ByVal statusText As String, _
                               ByVal durationSeconds As Double, ByVal outputText As String, _
                               ByVal detailText As String)
    Dim tblRunHistory As ListObject
    Dim historyRowNumber As Long

    Set tblRunHistory = ThisWorkbook.Worksheets("Run History").ListObjects("tblRunHistory")
    tblRunHistory.ListRows.Add
    historyRowNumber = tblRunHistory.ListRows.Count

    tblRunHistory.ListColumns("Timestamp").DataBodyRange.Cells(historyRowNumber, 1).Value = Now
    tblRunHistory.ListColumns("Report ID").DataBodyRange.Cells(historyRowNumber, 1).Value = reportID
    tblRunHistory.ListColumns("Report Name").DataBodyRange.Cells(historyRowNumber, 1).Value = reportName
    tblRunHistory.ListColumns("Action").DataBodyRange.Cells(historyRowNumber, 1).Value = actionText
    tblRunHistory.ListColumns("Status").DataBodyRange.Cells(historyRowNumber, 1).Value = statusText
    tblRunHistory.ListColumns("Duration (s)").DataBodyRange.Cells(historyRowNumber, 1).Value = _
        Round(durationSeconds, 1)
    tblRunHistory.ListColumns("Output").DataBodyRange.Cells(historyRowNumber, 1).Value = outputText
    tblRunHistory.ListColumns("Detail").DataBodyRange.Cells(historyRowNumber, 1).Value = detailText
End Sub

Private Sub CopyLatestHistoryToSummary()
    Dim wsSummary As Worksheet
    Dim wsConfiguration As Worksheet
    Dim tblRunHistory As ListObject
    Dim historyRowNumber As Long
    Dim summaryRowNumber As Long

    Set wsSummary = ThisWorkbook.Worksheets("Summary")
    Set wsConfiguration = ThisWorkbook.Worksheets("Configuration")
    Set tblRunHistory = ThisWorkbook.Worksheets("Run History").ListObjects("tblRunHistory")

    wsSummary.Range("N3").Value = wsConfiguration.Range("C11").Value
    wsSummary.Range("G35:M38").ClearContents
    wsSummary.Range("G35:G38").NumberFormat = "hh:mm"

    summaryRowNumber = 35
    For historyRowNumber = tblRunHistory.ListRows.Count To 1 Step -1
        wsSummary.Cells(summaryRowNumber, "G").Value = _
            tblRunHistory.ListColumns("Timestamp").DataBodyRange.Cells(historyRowNumber, 1).Value
        wsSummary.Cells(summaryRowNumber, "H").Value = _
            tblRunHistory.ListColumns("Report ID").DataBodyRange.Cells(historyRowNumber, 1).Value
        wsSummary.Cells(summaryRowNumber, "I").Value = _
            tblRunHistory.ListColumns("Action").DataBodyRange.Cells(historyRowNumber, 1).Value
        wsSummary.Cells(summaryRowNumber, "J").Value = _
            tblRunHistory.ListColumns("Status").DataBodyRange.Cells(historyRowNumber, 1).Value
        wsSummary.Cells(summaryRowNumber, "K").Value = _
            tblRunHistory.ListColumns("Duration (s)").DataBodyRange.Cells(historyRowNumber, 1).Value
        wsSummary.Cells(summaryRowNumber, "L").Value = _
            tblRunHistory.ListColumns("Output").DataBodyRange.Cells(historyRowNumber, 1).Value
        wsSummary.Cells(summaryRowNumber, "M").Value = _
            tblRunHistory.ListColumns("Detail").DataBodyRange.Cells(historyRowNumber, 1).Value

        summaryRowNumber = summaryRowNumber + 1
        If summaryRowNumber > 38 Then Exit For
    Next historyRowNumber

    wsSummary.Calculate
End Sub

Private Sub UpdateSummaryStatus(ByVal statusText As String, ByVal currentReport As String, _
                                ByVal progressValue As Double, ByVal messageText As String)
    Dim wsSummary As Worksheet
    Dim progressColors As Variant
    Dim completedSegments As Long
    Dim segmentNumber As Long

    Set wsSummary = ThisWorkbook.Worksheets("Summary")

    If progressValue < 0 Then progressValue = 0
    If progressValue > 1 Then progressValue = 1

    wsSummary.Range("G10").Value = statusText
    wsSummary.Range("G13").Value = currentReport
    wsSummary.Range("Q9").Value = progressValue
    wsSummary.Range("G16").Value = messageText

    Select Case statusText
        Case "FAILED", "NEEDS SETUP"
            wsSummary.Range("G10").Font.Color = RGB(214, 69, 69)
        Case "RUNNING"
            wsSummary.Range("G10").Font.Color = RGB(10, 110, 209)
        Case "WARNING"
            wsSummary.Range("G10").Font.Color = RGB(201, 137, 0)
        Case Else
            wsSummary.Range("G10").Font.Color = RGB(31, 157, 85)
    End Select

    progressColors = Array( _
        RGB(207, 234, 251), RGB(184, 225, 247), RGB(159, 216, 243), _
        RGB(132, 207, 238), RGB(103, 196, 234), RGB(73, 184, 228), _
        RGB(37, 172, 222), RGB(10, 158, 215), RGB(7, 140, 201), _
        RGB(8, 123, 188), RGB(9, 106, 173), RGB(10, 89, 159))

    completedSegments = WorksheetFunction.RoundUp(progressValue * 12, 0)

    For segmentNumber = 0 To 11
        If segmentNumber < completedSegments Then
            wsSummary.Range("G14").Offset(0, segmentNumber).Resize(2, 1).Interior.Color = _
                progressColors(segmentNumber)
        Else
            wsSummary.Range("G14").Offset(0, segmentNumber).Resize(2, 1).Interior.Color = _
                RGB(231, 238, 244)
        End If
    Next segmentNumber

    wsSummary.Calculate
    DoEvents
End Sub
