Attribute VB_Name = "Setup"
Dim arrSnowflakeConfigRanges() As String


Sub ConfigSnowflakeAddIn()
    If ActiveWorkbook Is Nothing Or Not ThisWorkbook.IsAddin Then
        Exit Sub
    End If
    
    For Each n In ThisWorkbook.Names
        On Error GoTo createName
    Next n

End Sub


'Public Const sgRangeWorksheetVersionNumber As String = "sfWorksheetVersionNumber"
'
''Connection
'Public Const sgRangeSnowflakeDriver As String = "sfSnowflakeDriver"
'Public Const sgRangeAuthType As String = "sfAuthType"
'Public Const sgRangeDefaultDatabase As String = "sfDatabase"
'Public Const sgRangeUserID As String = "sfUserID"
'Public Const sgRangeRole As String = "sfRole"
'Public Const sgRangeDefaultSchema As String = "sfSchema"
'Public Const sgRangeServer As String = "sfServer"
'Public Const sgRangeStage As String = "sfStage"
'Public Const sgRangeWarehouse As String = "sfWarehouse"
'Public Const sgRangePassword As String = "sfPassword"
'
'Public Const sgRangeResultsWorksheet As String = "sfResultsWorksheet"
'Public Const sgRangeUploadWorksheet As String = "sfUploadWorksheet"
'Public Const sgRangeLogWorksheet As String = "sfLogWorksheet"
'
'Public Const sgRangeDateInputFormat As String = "sfDateInputFormat"
'Public Const sgRangeTimestampInputFormat As String = "sfTimestampInputFormat"
'Public Const sgRangeTimeInputFormat As String = "sfTimeInputFormat"
'
'Public Const sgRangeWindowsTempDirectory As String = "sfWindowsTempDirectory"
'
'Sub setRangeDefaultValues()
'    Utils.CustomRange(sgRangeSnowflakeDriver) = "{SnowflakeDSIIDriver}"
'    Utils.CustomRange(sgRangeAuthType) = "User & Pass"
'    Utils.CustomRange(sgRangeLogWorksheet) = "Log"
'    Utils.CustomRange(sgRangeWindowsTempDirectory) = "C:\temp"
'    Utils.CustomRange(sgRangeDateInputFormat) = "Auto"
'    Utils.CustomRange(sgRangeTimestampInputFormat) = "Auto"
'    Utils.CustomRange(sgRangeTimeInputFormat) = "Auto"
'End Sub
