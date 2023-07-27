Attribute VB_Name = "Query"
' Not used yet
Dim arrSQLUpdateWords As Variant

Sub ExecuteSQL(sqlString As String)
    If sqlString = "" Then
        MsgBox "SQL string is empty"
    Else
        StatusForm.Update_Status ("Initializing...")
        Dim connString As String
        Dim queryResultWorksheet As Worksheet
        'get worksheet to store query results. Create if it doesn't exist
        'If the config is blank then use current active sheet
        If CustomRange(sgRangeResultsWorksheet) = "" Then
            Set queryResultWorksheet = ActiveSheet
        Else
            Set queryResultWorksheet = Utils.getWorksheet(CustomRange(sgRangeResultsWorksheet))
        End If
        queryResultWorksheet.Activate
        queryResultWorksheet.Cells.Clear
        'get the SQL to execute from the named cell = "SQL"

        Call Utils.RemoveQueryTables(queryResultWorksheet)
        
        sqlString = TrimTrailing(sqlString)
        'remove last char if it's a ;
        If (Right(sqlString, 1) = ";") Then
            sqlString = Left(sqlString, Len(sqlString) - 1)
        End If
        
        On Error GoTo ErrorHandlerExecSQL
        Dim arrQueries() As String
        arrQueries = Split(sqlString, ";")
        For i = LBound(arrQueries) To UBound(arrQueries) - 1
            arrQueries(i) = TrimTrailing(arrQueries(i))
            If arrQueries(i) <> "" Then
                StatusForm.Update_Status ("Executing query #" & i + 1 & "...")
                Utils.execSQLFireAndForget (arrQueries(i))
            End If
        Next i
        sqlString = TrimTrailing(arrQueries(i))
        If sqlString <> "" Then
            StatusForm.Update_Status ("Executing query...")
            
            Call Utils.ExecSQL(queryResultWorksheet, gsQueryResultsCell, sqlString)
            ' set download datetime
            Call Query.setDownloadDateTime
            'On Error GoTo ErrorHandlerCreateExcelTable
            StatusForm.Update_Status ("Creating Excel table...")
        Call Utils.createTableForAllDataOnWorksheet(queryResultWorksheet)
        End If
        StatusForm.Hide
        RibbonactivateHomeTab
    End If
    Exit Sub
ErrorHandlerExecSQL:
    If err.Number <> giCancelEvent Then
        'The error msg is handled in the query sub so I removed it from here
        'Call Utils.handleError("Error trying to execute query. ", err)
        StatusForm.Hide
    End If
    Exit Sub
ErrorHandlerCreateExcelTable:
    Call Utils.handleError("Error trying to format final table on worksheet. ", err)
    Exit Sub
End Sub
Function TrimTrailing(str1 As Variant) As String

Dim i As Integer
TrimTrailing = str1
For i = Len(str1) To 1 Step -1
    If Application.Clean(Trim(Mid(str1, i, 1))) = "" Then
        'skip
    Else 'stop at first non-empty character
        TrimTrailing = Mid(str1, 1, i)
        Exit For
    End If
Next i
If i = 0 Then 'if it hits the end
    TrimTrailing = ""
End If
End Function

Sub ExecuteSQLFromNamedCell(sql As String)
    ExecuteSQL (sql)
End Sub

Sub ExecuteSelectAllFromUploadTable()
    Dim table As String
    database = FormCommon.getDatabase()
    schema = FormCommon.getSchema()
    table = FormCommon.getTable()
    If database = "" Or schema = "" Or table = "" Then
        StatusForm.Hide
        MsgBox ("No valid SQL to execute.")
        Exit Sub
    End If
    sqlString = "select * from """ + database + """.""" + schema + """.""" + (table) + """"
    mergeKeys = FormCommon.getMergeKeys()
    If mergeKeys <> "" Then
        sqlString = sqlString & " order by (" & mergeKeys & ")"
    End If
    ExecuteSQL (sqlString)
    Call Query.setDownloadDateTime
End Sub

Sub OpenSQLForm()
    Set SQLForm = Nothing
    SQLForm.ShowForm
End Sub

Sub setDownloadDateTime()
    'Sets the date the sheet downloaded data. Used to check if data has changed when uploading
    'Dim lockRangeTableDate As range
    Dim currentTimeSQL As String
    On Error GoTo ErrorHandlerIgnoreError
    'Initialize Table locking Date range
    'Set lockRangeTableDate = FormCommon.initializeRange("LockTableDate")
    currentTimeSQL = "select to_char(current_Timestamp,'YYYY-MM-DD HH24:MI:SS.FF')"
    'lockRangeTableDate = Format(Utils.execSQLSingleValueOnly(currentTimeSQL), "YYYY-MM-DD HH24:Mmm:SS")
    FormCommon.initializeRange("LockTableDate") = Format(Utils.execSQLSingleValueOnly(currentTimeSQL), "YYYY-MM-DD HH24:Mmm:SS")
ErrorHandlerIgnoreError:
    'Do nothing
End Sub

Function getDownloadDateTime()
    getDownloadDateTime = FormCommon.initializeRange("LockTableDate")
End Function

' Started this function for catching if a user is trying to update data or DDL in the SQL textbox. Should not be allowed for Read only users. Not implementing yet
Function getArrSQLUpdateWords()
    If IsEmpty(arrSQLUpdateWords) Then
        arrSQLUpdateWords = Split(sgSQLUpdateWords, ",")
    End If
    getArrSQLUpdateWords = arrSQLUpdateWords
End Function
