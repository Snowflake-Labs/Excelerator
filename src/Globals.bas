Attribute VB_Name = "Globals"
' When adding a new named range to the Config sheet, you should update the Build.setRangeDefaultValues sub to set the defaults
'from Utils module
Public Const gWorkbookVers As String = "1.04"                                ' This is just for information
Public Const gWorkbookSPCompatibilityVers As Integer = 2                    ' Used to ensure the proper Stored Procs are used
Public Const gsQueryResultsCell As String = "A1"
Public Const giStartingRowForUpload As Integer = 1
Public Const sgLastCellOnLogWS As String = "A10000"
Public Const sgParameterFileName As String = "SnowflakeExcelAddin.ini"
Public Const sgParameterFileDirectory As String = "C:/temp"

'Checks versioning of worksheet. If active worksheet has an older one then it deletes it and copies a new one
Public Const sgRangeWorksheetVersionNumber As String = "sfWorksheetVersionNumber"

'Workbook specific named ranges that are dynamic
Public Const igAllSingleValueRages_ColNumber As Integer = 1
Public Const sgSavedSQL_SelectedIndex_RangePrefix As String = "sfSavedSQLSelectedIndex"
Public Const sgSavedSQL_LastExecutedSQL_RangePrefix As String = "sfSavedSQLLastExecutedSQL"

Public Const sgSavedSQL_Name_RangePrefix As String = "sfSavedSQLName"
Public Const sgSavedSQL_SQL_RangePrefix As String = "sfSavedSQLSQL"

' Last selected Database, Schema and tables
Public Const sgDBObj_LastSelectedDB_RangePrefix As String = "sfDBObjLastDB"
Public Const sgDBObj_LastSelectedSchema_RangePrefix As String = "sfDBObjLastSchema"
Public Const sgDBObj_LastSelectedTable_RangePrefix As String = "sfDBObjLastTable"

'For Table level optimistic locking
Public Const sgLockedDownloadTableDateTime_RangePrefix As String = "sfLockedDownloadTableDate"

'Rollback
Public Const sgRollbackSQL_RangePrefix As String = "sfRollbackSQL"
Public Const sgRollbackUploadDate_RangePrefix As String = "sfRollbackUploadDateTime"
Public Const sgRollbackUploadTableName_RangePrefix As String = "sfRollbackUploadTableName"

'Events
Public Const giCancelEvent As Integer = 2000
Public Const giSQLErrorEvent As Integer = 2001
Public Const giSuppressErrorMessage As Integer = 2002
Public Const giUndefinedError As Integer = 2002

'Connection
Public Const sgRangeSnowflakeDriver As String = "sfSnowflakeDriver"
Public Const sgRangeAuthType As String = "sfAuthType"
Public Const sgRangeDefaultDatabase As String = "sfDatabase"
Public Const sgRangeUserID As String = "sfUserID"
Public Const sgRangeRole As String = "sfRole"
Public Const sgRangeDefaultSchema As String = "sfSchema"
Public Const sgRangeServer As String = "sfServer"
Public Const sgRangeStage As String = "sfStage"
Public Const sgRangeWarehouse As String = "sfWarehouse"
Public Const sgRangePassword As String = "sfPassword"
Public Const sgRangeDSN As String = "sfDSN"  ' doesnt exist yet

'Enable button for read write
Public Const sgRangeReadOnly As String = "sfReadOnly"

'Worksheets
Public Const sgRangeResultsWorksheet As String = "sfResultsWorksheet"
Public Const sgRangeUploadWorksheet As String = "sfUploadWorksheet"
Public Const sgRangeLogWorksheet As String = "sfLogWorksheet"
Public Const gsSnowflakeWorkbookParamWorksheetName As String = "SnowflakeWorkbookParams"
Public Const gsSnowflakeConfigWorksheetName As String = "SnowflakeConfig"

'SQL & Load
'There is one of these per sheet
Public Const sgUploadMergeKeys_RangePrefix As String = "sfUploadMergeKeys"
Public Const sgUploadType_RangePrefix As String = "sfUploadType" 'merge, truncate, append, recreate

Public rgUploadMergeKeysRange As range
Public Const sgUploadMergeKeysByLetters_RangePrefix As String = "sfUploadMergeKeysByLetters"

Public Const sgRangeWindowsTempDirectory As String = "sfWindowsTempDirectory"

Public Const sgRangeDateInputFormat As String = "sfDateInputFormat"
Public Const sgRangeTimestampInputFormat As String = "sfTimestampInputFormat"
Public Const sgRangeTimeInputFormat As String = "sfTimeInputFormat"

'Show upload warning
Public sgShowVerifyMsg As String

'Ribbon
Public gRibbon As IRibbonUI

'Snowflake meta data
Public gArrDatabases() As String

'list of data types
Public Const sgDatatypes = "Text,Integer,Date,Timestamp,Double,Number,Number(p s),Varchar(n),Float,Boolean,Time,Variant,Object,Array"

'list of words that signify some kind of modify statement is written
Public Const sgSQLUpdateWords = "DROP ,UPDATE , DELETE, INSERT, TRUNCATE , MERGE , ALTER , CREATE "
