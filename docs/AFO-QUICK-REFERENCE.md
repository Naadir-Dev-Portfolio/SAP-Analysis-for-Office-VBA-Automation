# AFO Quick Reference

This is the short explanation of the project. The complete template adds configuration, progress, logging and exports around these calls.

## 1. Refresh AFO

```vb
Sub Refresh_AFO()
    Dim lResult As Long

    lResult = Application.Run("SAPExecuteCommand", "RefreshData")

    If lResult <> 1 Then
        Err.Raise vbObjectError + 700, "Refresh_AFO", _
                  "SAP Analysis for Office refresh failed."
    End If
End Sub
```

Add a data-source alias as the third argument to refresh only one source:

```vb
lResult = Application.Run("SAPExecuteCommand", "RefreshData", "DS_1")
```

## 2. Log on

```vb
Dim userName As String
Dim password As String

userName = ThisWorkbook.Worksheets("Configuration").Range("C16").Value
password = ThisWorkbook.Worksheets("Configuration").Range("C17").Value

lResult = Application.Run("SAPLogon", "DS_1", "500", userName, password)
```

- `DS_1` is the data-source alias in the AFO Design Panel.
- `500` is the SAP client and must be configured for the target system.
- The template reads the username and password from **Configuration** immediately above the call.
- The packaged workbook contains placeholders; enter live credentials only in a private working copy.
- A successful call returns `1`.

## 3. Set a prompt

```vb
lResult = Application.Run( _
    "SAPSetFilter", _
    "DS_1", _
    "P_COMP_CODE", _
    "1000", _
    "INPUT_STRING")
```

For multiple values, pass a semicolon-separated string:

```vb
promptValue = "1000; 2000; 3000"
```

The template reads each active row in `tblPrompts` and turns it into one `SAPSetFilter` call.

## 4. Read source information

Use this AFO formula in a worksheet cell:

```excel
=@SAPGetSourceInfo("DS_1", "DataSourceName")
```

Replace `DS_1` with the alias used by that workbook.

For a workbook that must also open cleanly on a machine without AFO, use a display fallback:

```excel
=IFERROR(@SAPGetSourceInfo("DS_1", "DataSourceName"), "AFO ADD-IN NOT DETECTED")
```

The portfolio Summary adds a SIMULATION-mode check around this so it can state clearly that SAP calls are disabled instead of showing an Excel formula error.

## Date prompt recipes

```vb
PreviousCompleteWeeksDaily(13)
PreviousWeekWorkingDays(13)
CurrentFiscalYearDates(Date, 4)
FiscalMonthFirstWorkingDays(Date, 4)
```

Each function returns an AFO-ready `INPUT_STRING` value in `dd.mm.yyyy` format. Weekends and dates in `tblHolidays` are skipped by the working-day routines.

## Batch flow

`RefreshSelectedReport` contains the complete one-report workflow:

1. Read the selected row from `tblReports`.
2. Log on with `SAPLogon` in LIVE mode.
3. Resolve each active prompt and apply it with `SAPSetFilter` in LIVE mode.
4. Refresh with `SAPExecuteCommand`, `RefreshData`; SIMULATION mode simulates only this SAP boundary.
5. Update the report row and append one audit entry to `tblRunHistory`.

`RunActiveBatch` loops through the active report rows and reuses `RefreshSelectedReport` followed by `ExportSelectedReport`. That means the selected and batch buttons do not contain two different implementations of the refresh logic.

## Interview explanation

> The workbook uses table-driven report and prompt configuration. The three AFO calls are log on, set filters and refresh, and they are visible together inside `RefreshSelectedReport`. Every call must return 1. The batch macro loops through active report rows and reuses the same refresh and export macros. Offline simulation mode skips only the unavailable SAP calls, so the rest of the workflow can still be reviewed without an SAP system.

## Module map

- `modAFO_Core`: actual AFO calls; start here.
- `modDatePrompts`: reusable date-selection functions.
- `modAutomation`: one readable public macro per Summary button, plus three small display/history procedures at the bottom.
- `ThisWorkbook`: startup event only.
