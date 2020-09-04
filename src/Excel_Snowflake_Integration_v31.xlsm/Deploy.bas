Attribute VB_Name = "Deploy"

Sub createAddin()
    Dim wb As Workbook
    'Delete all worksheets except the config one
    CleanupWorksheets
    ' Remove ranges that are invalid and set others to empty
    CleanupRanges
    ' Application.DisplayAlerts = False
    Set wb = Workbooks(ActiveWorkbook.name)
    wb.SaveAs fileName:=ThisWorkbook.Path & "\" & "SnowFlowExcelAddin.xlam", FileFormat:=xlOpenXMLAddIn, CreateBackup:=False
End Sub

Sub CleanupRanges()
    For Each n In ActiveWorkbook.Names
        If InStr(n.value, "#REF!") > 0 Then
            n.Delete
        End If
    Next
    For Each n In ActiveWorkbook.Names
        If n.name = sgRangeDefaultDatabase Or n.name = sgRangeUserID Or n.name = sgRangeRole Or n.name = sgRangeDefaultSchema _
                Or n.name = sgRangeServer Or n.name = sgRangeStage Or n.name = sgRangeWarehouse Or n.name = sgRangePassword _
                Or n.name = sgRangeResultsWorksheet Or n.name = sgRangeUploadWorksheet Then
            range(n) = ""
        End If
    Next n
End Sub

Sub setRangeDefaultValues()
    Utils.CustomRange(sgRangeSnowflakeDriver) = "{SnowflakeDSIIDriver}"
    Utils.CustomRange(sgRangeAuthType) = "User & Pass"
    Utils.CustomRange(sgRangeLogWorksheet) = "Log"
    Utils.CustomRange(sgRangeWindowsTempDirectory) = "C:\temp"
    Utils.CustomRange(sgRangeDateInputFormat) = "Auto"
    Utils.CustomRange(sgRangeTimestampInputFormat) = "Auto"
    Utils.CustomRange(sgRangeTimeInputFormat) = "Auto"
End Sub

Sub CleanupWorksheets()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(gsSnowflakeConfigWorksheetName).Visible = True
    For Each ws In Worksheets
        If ws.name <> gsSnowflakeConfigWorksheetName Then
            ws.Delete
        End If
    Next
End Sub






