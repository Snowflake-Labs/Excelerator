Attribute VB_Name = "Load"
'vba code to convert excel to csv
'to set ODBC log levels and directory update /opt/snowflake/snowflakeodbc/lib/universal/simba.snowflake.ini
Private Type fileNameWithPath
    save As String
    put As String
End Type

Dim tableName As String
Dim mergeKeys As String
Dim fileName As String
Dim connString As String
Dim dataWorksheet As Worksheet
Dim statusWorksheet As Worksheet
Dim FullFileNameWithPath As fileNameWithPath
Dim columnHeadersProvided As Boolean
Dim iHeaderRow As Integer
Dim clonedTable As String
Dim stageName As String
Dim sqlFileFormatSkipHeader As String

Sub OpenUploadDataForm()
    Set UploadDataForm = Nothing
    If Utils.login Then
        StatusForm.Update_Status ("Preparing Upload Form...")
        Call UploadDataForm.ShowForm
    End If
End Sub

Sub uploadData(copyType As String, inTableName As String, inMergeKeys As String)
    StatusForm.Update_Status ("Initializing...")
    Dim windowsTempDirectory As String 'Only for windows. Folder where csv file is stored
    Dim sqlString As String
    Dim stageCreated As Boolean
    Dim authType As String
    'These 3 are used to check if the stored procs are up to date or if this is an older workbook
    Dim spVers As String
    Dim spMinVers As Integer
    Dim spMaxVers As Integer
    Dim msg As String
    Dim bDataTypeRowExists As Boolean

    'The final message, success or failure
    Dim completionStatus As String
    completionStatus = "Upload Failed!"

    'reseting the clonedTable
    clonedTable = ""

    '************ Gets the Worksheets and create them if needed ***********
    On Error GoTo ErrorHandlerUploadWorksheetDoesNotExist
    ' DATA - remove all empty rows in data worksheet
    Set dataWorksheet = getDataWorksheet()
    On Error GoTo 0
    Call DeleteAllEmptyRows(dataWorksheet)
    'LOG - Clear
    Set statusWorksheet = Utils.getWorksheet(CustomRange(sgRangeLogWorksheet))
    statusWorksheet.Cells.Clear

    '************ Checking for Data Type row ************
    Dim firstCellValue As String
    firstCellValue = dataWorksheet.Cells(giStartingRowForUpload, 1).value
    ' Check if first row has the data types.
    iHeaderRow = giStartingRowForUpload
    Dim arrDatatypes() As String
    arrDatatypes = Split(sgDatatypes, ",")
    If IsInArray(firstCellValue, arrDatatypes) Or firstCellValue = "" Or InStr(firstCellValue, "(") Then
        iHeaderRow = iHeaderRow + 1
        bDataTypeRowExists = True
    End If

    '************ Checking to make sure the first column name has a value ************
    If dataWorksheet.range("A" & iHeaderRow) = "" Then
        MsgBox ("Worksheet '" & dataWorksheet.name & "' is empty or the first cell is empty. Nothing to upload.")
        StatusForm.Hide
        Exit Sub
    End If

    mergeKeys = inMergeKeys
    tableName = inTableName  ' this is the fully qualified table name
    fileName = UCase(tableName & "TEMP.CSV")
    fileName = Replace(fileName, Chr(34), vbNullString, 1, 6)
    windowsTempDirectory = CustomRange(sgRangeWindowsTempDirectory)

    'authType = CustomRange(sgRangeAuthType)
    StatusForm.Update_Status ("Exporting data to csv file ...")
    On Error GoTo Cleanup
    'Create csv file with the name of the table
    FullFileNameWithPath = fullFileName(fileName, windowsTempDirectory)
    'Create the stage
    stageCreated = createStage(stageName)
    'Save CSV file and execute 'put'
    saveAndUploadFile

    On Error GoTo ErrorHandlerSimpleUpdate
    Select Case True
        Case copyType = "MergeLocal" Or copyType = "AppendLocal"
            addColumns
            copyDataLocal (copyType)
        Case copyType = "TruncateLocal"
            addColumns
            truncateTable
            copyDataLocal (copyType)
        Case copyType = "RecreateLocal" Or copyType = "CreateLocal"
            If copyType = "CreateLocal" Then On Error GoTo ErrorHandlerSimpleUpdateNoRollback
            createTableLocal
            copyDataLocal (copyType)
            Call FormCommon.dropDBObjectsTableCache
        Case Else
            On Error GoTo ErrorHandleruploadNeedsStoredProcs
            clonedTable = createClone(tableName)
            uploadNeedsStoredProcs (copyType)
            If copyType = "RecreateTable" Then
                Call FormCommon.dropDBObjectsTableCache
            End If
    End Select

    ' Set SQL, table and Date ranges for Rollback later
    Load.setRollbackSQL
    Load.setUploadDateTime
    Load.setUploadFullyQualifiedTableName

    If clonedTable <> "NoTableToClone" Then
        ' Drop the clone table if it was created
        Call droptable(clonedTable)
    End If


    completionStatus = "Upload Success!"
    StatusForm.Hide
    answer = MsgBox("Upload Success!" & vbNewLine & vbNewLine & "Refresh data from table?", vbYesNo + vbQuestion, "Upload Status")

    If answer = vbYes Then
        Call StatusForm.execMethod("Query", "ExecuteSelectAllFromUploadTable")
        'Call Query.ExecuteSelectAllFromUploadTable
    Else
        ' Delete Data Type row if it exists
        If bDataTypeRowExists Then
            dataWorksheet.Rows(1).Delete
        End If
    End If
    GoTo Cleanup

ErrorHandlerUploadWorksheetDoesNotExist:
    MsgBox ("The source worksheet: '" & CustomRange(sgRangeUploadWorksheet) & "' does not exist.")
    GoTo Cleanup
ErrorHandlerSimpleUpdate:
    StatusForm.Hide
    If err.Number <> giSuppressErrorMessage And err.Number <> 0 Then
        Call Utils.handleError("Error:", err)
    End If
    'this needs to come after the err.Number check beacuse it resets the err number
    Call rollback(tableName, clonedTable)
    Exit Sub
ErrorHandlerSimpleUpdateNoRollback:
    StatusForm.Hide
    Exit Sub
ErrorHandleruploadNeedsStoredProcs:
    If err.Number = giSuppressErrorMessage Then
        StatusForm.Hide
    Else
        GoTo Cleanup
    End If
    Exit Sub
Cleanup:
    If stageCreated Then
        StatusForm.Update_Status ("Cleaning up, dropping stage...")
        Call Utils.execSQLFireAndForget("drop stage " & stageName)
    End If
    If err.Number = giSuppressErrorMessage Then
        StatusForm.Hide
    Else
        StatusForm.Update_Status (completionStatus)
    End If
End Sub

Function getDataWorksheet() As Worksheet
    If CustomRange(sgRangeUploadWorksheet) <> "" Then
        Set getDataWorksheet = ActiveWorkbook.Sheets(CustomRange(sgRangeUploadWorksheet).value)
    Else
        Set getDataWorksheet = ActiveSheet
    End If
End Function

Function createStage(ByRef stageName As String)
    Dim sqlString As String
    stageName = CustomRange(sgRangeStage)
    If stageName = "" Then
        On Error GoTo ErrorHandlerCreateStage
        StatusForm.Update_Status ("Creating Stage...")
        'Generate random number for stagename
        randomNumber = Int(1 + Rnd * (1000)) * Second(Now())
        stageName = "ExcelStage_" & Replace(fileName, ".", "_") & randomNumber
        sqlString = "create or replace stage " & stageName & " file_format = (type=csv, SKIP_HEADER=0,FIELD_OPTIONALLY_ENCLOSED_BY = '""')"
        Call Utils.execSQLFireAndForget(sqlString)
        createStage = True
    Else
        createStage = False
        On Error GoTo ErrorHandlerRemoveFileFromStage
        'delete the file in the stage
        Call Utils.execSQLFireAndForget("remove @" & stageName & "/" & fileName & "")
    End If
    Exit Function
ErrorHandlerCreateStage:
    Call Utils.handleError("Unable to create stage. Check to see if role has proper privilege. ", err)
    err.Raise giSuppressErrorMessage
ErrorHandlerRemoveFileFromStage:
    Call Utils.handleError("Error occured while attempting to delete file from stage '" & stageName & "'.", err)
    err.Raise giSuppressErrorMessage
End Function

Sub copyDataLocal(copyType As String)
    Dim sqlString As String
    If copyType = "MergeLocal" Then
        On Error GoTo ErrorHandlerAlterStage
        'Changing the skip _header is only needed on the merge becuse it is built into the 'copy' sql
        sqlFileFormatSkipHeader = "alter stage " & stageName & " set FILE_FORMAT = (FIELD_OPTIONALLY_ENCLOSED_BY = '""', SKIP_HEADER=" & iHeaderRow & ")"
        Call Utils.execSQLFireAndForget(sqlFileFormatSkipHeader)
        sqlString = createMergeSQL
    Else
        sqlString = createCopySQL
    End If

    'load the table
    StatusForm.Update_Status ("Loading data...")
    On Error GoTo ErrorHandlerCopy
    Call Utils.execSQLFireAndForget(sqlString)
    Exit Sub
ErrorHandlerAlterStage:
    Call Utils.handleError("Error occured executing: " & sqlFileFormatSkipHeader, err)
    err.Raise giSuppressErrorMessage
ErrorHandlerCopy:
    handleLoadError (err.Description)
    err.Raise giSuppressErrorMessage
End Sub
Sub handleLoadError(errDescription)
    Dim table As String
    Select Case True
        Case InStr(errDescription, "invalid identifier")
            'Find Column name in error message
            i = InStr(1, errDescription, "'", vbTextCompare) + 1
            j = InStr(i, errDescription, "'", vbTextCompare)
            foundString = Mid(errDescription, i, j - 1)
            MsgBox ("Column not found in table. Column name: '" & foundString & vbNewLine & vbNewLine & _
                 "If this is a new column you want to add to the table, please set the data type by selecting " & _
                 "the 'Define Data Types' button in the Excel Ribbon.")
        Case InStr(errDescription, "is not recognized")
            MsgBox ("Data type mismatch Error:" & vbNewLine & Left(err.Description, InStr(err.Description, "is not recognized") + 17) & _
            vbNewLine & "You can change the default data types in the Config button in the Excel Ribbon")
        Case InStr(err.Description, "does not exist or not authorized")
            MsgBox ("Table " & tableName & " does not exist." & _
            vbNewLine & "To create the table please select the 'Create / Recreate table' option.")
        Case Else
            Call Utils.handleError("Error in upload", err)
    End Select
End Sub
Sub uploadNeedsStoredProcs(copyType)
    'COPYTYPE: 'RecreateTable','Truncate', 'Upsert', 'Append', 'CreateTableOnly', 'Merge'
    Dim sqlString As String
    '*********************** Check stored proc version compatibility ***********************************

    If Not Utils.checkStoredProcCompatibility(statusWorksheet) Then
        err.Raise giSuppressErrorMessage
    End If

    'load the table
    On Error GoTo ErrorHandlerUpload
    StatusForm.Update_Status ("Loading data...")
    Dim numberOfColumns As Integer
    numberOfColumns = dataWorksheet.UsedRange.columns.Count
    sqlString = "call create_table_from_file_and_load('" & tableName & "','" & stageName & "','" & fileName & "','" & copyType & "','" & mergeKeys & "'," & numberOfColumns & ");"
    Call Utils.ExecSQL(statusWorksheet, nextStatusCellToLoad, sqlString)
    Exit Sub
ErrorHandlerUpload:
    Call Utils.handleError("Error uploading file. ", err)
    err.Raise 1000
ErrorHandlerPutFile:
    Call Utils.handleError("Unable to 'PUT' files to stage. ", err)
    err.Raise 1000
End Sub
Sub saveAndUploadFile()
    Dim sqlString As String
    'save csv to temp directory
    On Error GoTo ErrorHandlerSaveFile
    Call SaveWorksheet(dataWorksheet, FullFileNameWithPath.save)
    'put the csv into stage
    StatusForm.Update_Status ("Uploading File...")
    On Error GoTo ErrorHandlerPutFile
    sqlString = "put 'file://" & FullFileNameWithPath.put & "' @" & stageName & ";"
    Call Utils.ExecSQL(statusWorksheet, nextStatusCellToLoad, sqlString)
    Exit Sub

ErrorHandlerSaveFile:
    Call Utils.handleError("Error Saving file.", err)
    err.Raise 1000
ErrorHandlerPutFile:
    Call Utils.handleError("Unable to 'PUT' files to stage. ", err)
    err.Raise 1000
End Sub
Function CreateFolderinMacOffice2016(NameFolder As String) As String
    Dim OfficeFolder As String
    Dim PathToFolder As String

    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
        "Library/Group Containers/UBF8T346G9.Office/"

    PathToFolder = OfficeFolder & NameFolder
    dirExists = Dir(PathToFolder & "*", vbDirectory)
    If dirExists = vbNullString Then
        MkDir PathToFolder
    End If
    CreateFolderinMacOffice2016 = PathToFolder
End Function


Sub DeleteAllEmptyRows(ws As Worksheet)
    Dim LastRowIndex As Long
    Dim RowIndex As Long
    Dim UsedRng As range

    Set UsedRng = ws.UsedRange
    LastRowIndex = UsedRng.row - 1 + UsedRng.Rows.Count
    Application.ScreenUpdating = False

    For RowIndex = LastRowIndex To 1 Step -1
        If Application.CountA(ws.Rows(RowIndex)) = 0 Then
            ws.Rows(RowIndex).Delete
        End If
    Next RowIndex

    Application.ScreenUpdating = True
End Sub

Sub DeleteAllEmptyColumns(ws As Worksheet)
    Dim LastColIndex As Long
    Dim ColIndex As Long
    Dim UsedRng As range

    Set UsedRng = ws.UsedRange
    LastColIndex = UsedRng.Column - 1 + UsedRng.Column.Count
    Application.ScreenUpdating = False

    For ColIndex = LastRowIndex To 1 Step -1
        If Application.CountA(ws.columns(ColIndex)) = 0 Then
            ws.columns(ColIndex).Delete
        End If
    Next ColIndex

    Application.ScreenUpdating = True
End Sub

Function fullFileName(fileName As String, windowsDirectory As String) As fileNameWithPath
    Dim windowsTempDirectory As String
    Dim snowflakeDirectory As String

    #If Mac Then
    ' FullFileName.save = "/excel" & "/" & fileName
    fullFileName.save = CreateFolderinMacOffice2016(NameFolder:="snowflake_put") & "/" & fileName
    fullFileName.put = fullFileName.save
    #Else
    On Error GoTo ErrorHandlerGetDirectory
    windowsTempDirectory = windowsDirectory
    snowflakeDirectory = windowsDirectory & "\Snowflake"
    dirExists = vba.FileSystem.Dir(windowsTempDirectory, vbDirectory)
    If dirExists = vba.Constants.vbNullString Then
        MkDir windowsTempDirectory
    End If
    dirExists2 = vba.FileSystem.Dir(snowflakeDirectory, vbDirectory)
    If dirExists2 = vba.Constants.vbNullString Then
        MkDir snowflakeDirectory
    End If
    'FullFileName = directory & "/" & fileName
    fullFileName.save = snowflakeDirectory & "\" & fileName
    fullFileName.put = Replace(fullFileName.save, "\", "\\")
    #End If
    Debug.Print "Saving File to = " & fullFileName.save
    Debug.Print "put from location = " & fullFileName.put
    Exit Function
ErrorHandlerGetDirectory:
    Call Utils.handleError("Error trying to get or create directory '" & windowsDirectory & "\Snowflake'.", err)
    err.Raise giSuppressErrorMessage
End Function

Sub SaveWorksheet(ws As Worksheet, fileName As String)
    Dim wb As Workbook
    ws.Copy
    Set wb = Workbooks(ActiveWorkbook.name)
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=fileName, FileFormat:=xlCSV, CreateBackup:=False
    Debug.Print ("Active Workbook:" & ActiveWorkbook.name)

    wb.Close
    Application.DisplayAlerts = True

End Sub

Sub AddDataTypeDropDowns()
    Dim rRange As range
    Dim t: t = Null

    Set dataWorksheet = getDataWorksheet()
    ' need to activate this because this Cells(giStartingRowForUpload, 1), will get the value of the active cell
    dataWorksheet.Activate
    Set UsedRng = dataWorksheet.UsedRange
    LastColIndex = UsedRng.columns.Count  'UsedRng.Rows.Count
    ' If there isn't data then bail
    If LastColIndex > 0 Then
        'Check if the first cell has a dropdown already. If it does than it means that we should update not insert the row
        On Error Resume Next
        t = dataWorksheet.Cells(giStartingRowForUpload, 1).Validation.Type
        On Error GoTo 0
        If IsNull(t) Then 'There is no dropdown so Insert
            dataWorksheet.Rows(giStartingRowForUpload).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If

        Set rRange = dataWorksheet.range(dataWorksheet.Cells(giStartingRowForUpload, 1), dataWorksheet.Cells(giStartingRowForUpload, LastColIndex))

        With rRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
      xlBetween, Formula1:=sgDatatypes
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = False
        End With
    End If
End Sub

Function createMergeSQL()
    Dim numberOfColumns As Integer
    Dim tableColumnsCSV As String
    Dim fileColumnsCSV As String
    Dim matchByClause As String
    Dim keys() As String
    Dim colName As String
    Dim updateClause As String
    Dim matchByClauseIfChanged As String
    'Todo need to get rid of this since it's a global
    quote = """"
    fileColumnsCSV = ""
    tableColumnsCSV = ""
    matchByClause = ""
    updateClause = ""
    matchByClauseIfChanged = ""

    'create array of keys
    keys = Split(mergeKeys, ",")

    numberOfColumns = dataWorksheet.UsedRange.columns.Count
    For i = 1 To numberOfColumns
        colName = dataWorksheet.Cells(iHeaderRow, i)
        If colName <> "" Then
            colName = quote & colName & quote
            fileColumnsCSV = fileColumnsCSV & ",$" & i
            tableColumnsCSV = tableColumnsCSV & "," & colName
            If IsInArray(i, keys) Then
                matchByClause = matchByClause & " and " & colName & " = fileSource.$" & i
            Else
                updateClause = updateClause & ", " & colName & " = fileSource.$" & i
                matchByClauseIfChanged = matchByClauseIfChanged & " and " & colName & " = fileSource.$" & i & " and "
                matchByClauseIfChanged = matchByClauseIfChanged & " not(" & colName & " is Null and fileSource.$" & i & " is not null ) and "
                matchByClauseIfChanged = matchByClauseIfChanged & " not(" & colName & " is not Null and fileSource.$" & i & " is null ) "
            End If
        End If
    Next i

    'remove leading chars
    fileColumnsCSV = Right(fileColumnsCSV, Len(fileColumnsCSV) - 1)
    tableColumnsCSV = Right(tableColumnsCSV, Len(tableColumnsCSV) - 1)
    matchByClause = Right(matchByClause, Len(matchByClause) - 5)

    createMergeSQL = "merge into " & tableName & " using (select " & fileColumnsCSV & " from @" & stageName & "/" & fileName & ") as fileSource on " & matchByClause
    If updateClause <> "" Then
        'remove leading chars
        matchByClauseIfChanged = Right(matchByClauseIfChanged, Len(matchByClauseIfChanged) - 5)
        updateClause = Right(updateClause, Len(updateClause) - 1)
        createMergeSQL = createMergeSQL & " WHEN MATCHED AND NOT(" & matchByClauseIfChanged & ") THEN UPDATE SET " & updateClause
        'Test with updating everything'  -  createMergeSQL = createMergeSQL & " WHEN MATCHED  THEN UPDATE SET " & updateClause
    End If
    createMergeSQL = createMergeSQL & " WHEN NOT MATCHED THEN INSERT (" & tableColumnsCSV & ") VALUES (" & fileColumnsCSV & ")"
End Function

Function createCopySQL()
    Dim numberOfColumns As Integer
    Dim tableColumnsCSV As String
    Dim fileColumnsCSV As String
    Dim matchByClause As String
    Dim keys() As String
    Dim colName As String
    Dim updateClause As String
    Dim matchByClauseIfChanged As String

    'These will be set later when we add more functionality
    columnHeadersProvided = True
    quote = """"
    fileColumnsCSV = ""
    tableColumnsCSV = ""
    matchByClause = ""
    updateClause = ""
    matchByClauseIfChanged = ""

    If columnHeadersProvided Then
        Set dataWorksheet = getDataWorksheet()
        numberOfColumns = dataWorksheet.UsedRange.columns.Count
        For i = 1 To numberOfColumns
            colName = dataWorksheet.Cells(iHeaderRow, i)
            If colName <> "" Then
                fileColumnsCSV = fileColumnsCSV & ",$" & i
                tableColumnsCSV = tableColumnsCSV & "," & quote & colName & quote
            End If
        Next i
        'Remove leading comma
        fileColumnsCSV = Right(fileColumnsCSV, Len(fileColumnsCSV) - 1)
        tableColumnsCSV = Right(tableColumnsCSV, Len(tableColumnsCSV) - 1)
        createCopySQL = "copy into " & tableName & " (" & tableColumnsCSV & _
        ") from (select " & fileColumnsCSV & "from @" & stageName & "/" & fileName & _
        ")  FILE_FORMAT = (FIELD_OPTIONALLY_ENCLOSED_BY = '""', SKIP_HEADER=" & iHeaderRow & ")"
    Else
        createCopySQL = "copy into " & tableName & " from  @" & stageName & "/" & fileName
    End If

End Function


Public Function IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
    Dim i
    stringToBeFound = CStr(stringToBeFound)
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Sub truncateTable()
    clonedTable = createClone(tableName)
    Call Utils.execSQLFireAndForget("truncate table " & tableName)
End Sub

Sub createTableLocal()
    'Assume that the check for the data type row is done prior
    Dim datatype As String
    Dim colName As String
    Dim colDefinition As String
    Dim sqlString As String
    If iHeaderRow = 1 Then
        MsgBox ("Please set the data types by clicking the 'Define Data Types' button in the ribbon.")
        err.Raise giSuppressErrorMessage
    End If
    On Error GoTo ErrorHandlerCreateTable
    'Create clone table
    clonedTable = createClone(tableName)

    numberOfColumns = dataWorksheet.UsedRange.columns.Count
    For i = 1 To numberOfColumns
        colName = dataWorksheet.Cells(iHeaderRow, i)
        If colName <> "" Then
            datatype = dataWorksheet.Cells(iHeaderRow - 1, i)
            If datatype = "" Then
                MsgBox ("Error: Data type for " & colName & " is not defined. Please set a data type for all columns")
                err.Raise giSuppressErrorMessage
            End If
            colDefinition = colDefinition & "," & colName & " " & datatype
        End If
    Next i
    'remove leading chars
    colDefinition = Right(colDefinition, Len(colDefinition) - 1)
    sqlString = "create or replace table " & tableName & " (" & colDefinition & ")"

    Utils.execSQLFireAndForget (sqlString)
    Exit Sub

ErrorHandlerCreateTable:
    If err.Number = giSuppressErrorMessage Then
        err.Raise giSuppressErrorMessage
    End If
    Call Utils.handleError("Problem creating table", err)
    err.Raise giUndefinedError
End Sub

Sub addColumns()
    Dim datatype As String
    Dim colName As String
    Dim colDefinition As String
    Dim sqlString As String

    'If no header row then bail
    If iHeaderRow = 1 Then
        Exit Sub
    End If

    'Create clone table
    clonedTable = createClone(tableName)

    On Error GoTo ErrorHandlerAddColumns
    numberOfColumns = dataWorksheet.UsedRange.columns.Count
    For i = 1 To numberOfColumns
        colName = dataWorksheet.Cells(iHeaderRow, i)
        datatype = dataWorksheet.Cells(iHeaderRow - 1, i)
        If datatype <> "" Then
            alterSQL = "alter table " & tableName & " add column """ & colName & """ " & datatype
            Utils.execSQLFireAndForget (alterSQL)
            FormCommon.dropDBObjectsTableCache
        End If
    Next i
    Exit Sub

ErrorHandlerAddColumns:
    Select Case True
        Case InStr(err.Description, "does not exist or not authorized")
            MsgBox ("Error: Table " & tableName & " does not exist." & _
            vbNewLine & "To create the table please select the 'Create / Recreate table' option.")
        Case InStr(err.Description, "already exists")
            MsgBox ("Error: Column '" & colName & "' already exists." & vbNewLine & _
            "You should not define the data type for a column that already exists.")
        Case InStr(err.Description, "Unsupported data type")
            MsgBox ("Error: Unsupported data type '" & datatype & "'." & vbNewLine & _
            "Please select a valid data type.")
        Case Else
            Call Utils.handleError("Problem adding column " & colName, err)
    End Select
    err.Raise giSuppressErrorMessage
End Sub

Function createClone(origTable As String)
    'returns name of cloned table
    Dim cloneSQL As String
    Dim generatedClonedTableName As String

    generatedClonedTableName = createClonedTableName(origTable)
    cloneSQL = "create or replace table " & generatedClonedTableName & " clone " & origTable
    On Error GoTo ErrorHandlerCloneNotCreated
    Call Utils.execSQLFireAndForget(cloneSQL)
    createClone = generatedClonedTableName
    Exit Function
ErrorHandlerCloneNotCreated:
    createClone = "NoTableToClone"
End Function

Function createClonedTableName(origTable As String)
    'the table will be fully qualified with quotes so we need to get rid of the quotes then just take the table
    Dim tempTable As String
    'Remove the quotes
    tempTable = Replace(origTable, Chr(34), vbNullString, 1, 6)
    Dim tableArr() As String
    tableArr = Split(tempTable, ".")
    tempTable = tableArr(UBound(tableArr))

    createClonedTableName = tempTable & (Int(1 + Rnd * (1000)) * Second(Now())) & "_BackupForExcel"

End Function
Sub rollback(origTable As String, clonedTable As String)
    Dim swapSQL As String
    On Error GoTo ErrorHandlerSQL
    If clonedTable <> "" Then
        swapSQL = "alter table " & origTable & " swap with " & clonedTable
        Utils.execSQLFireAndForget (swapSQL)
    End If
    Exit Sub
ErrorHandlerSQL:
    Call Utils.handleError("Rolling back to cloned table", err)
End Sub

Sub droptable(table As String)
    On Error GoTo ErrorHandlerSQL
    If table <> "" Then
        Utils.execSQLFireAndForget ("drop table " & table)
    End If
    Exit Sub
ErrorHandlerSQL:
    Call Utils.handleError("Rolling back to cloned table", err)
End Sub

Function grantAllPrivsToClonedTableSQL(origTable As String, clonedTable As String)
    Dim sql As String

    Call Utils.execSQLFireAndForget("show grants on " & origTable)
    sql = "with ShowGrants (created_on,privilege,granted_on,name,granted_to,grantee_name,grant_option,granted_by) as " & _
    "(select * from  table(result_scan(last_query_id()))) " & _
    "select 'grant '|| privilege|| ' on table ' || '" & clonedTable & "' || ' to '|| granted_to ||' ' || grantee_name || " & _
    "IFF( grant_option='true' , ' with grant option' ,'' ) " & _
    "from ShowGrants where privilege <>'OWNERSHIP' order by created_on desc"

    grantAllPrivsToClonedTableSQL = Utils.execSQLReturnConcatResults(sql, "~")
End Function

Sub RollbackLastUpdateWithCheck()
    If Load.getRollbackSQL = "" Then
        MsgBox ("Nothing to rollback.")
    Else
        If Load.checkIfTableHasBeenAltered(Load.getUploadDateTime, Load.getUploadFullyQualifiedTableName) = True Then
            If MsgBox("Table " & Load.getUploadFullyQualifiedTableName & " has been updated by someone else since your last upload." & _
                "If you Rollback, their changes will be lost." & _
                        vbNewLine & "Continue to Rollback?", vbOKCancel, "Rollback Conflict") = vbCancel Then
                Exit Sub
            End If
        Else
            MsgBoxVerifyRollback = MsgBox("This will rollback table:   " & Load.getUploadFullyQualifiedTableName & vbNewLine & _
                "to the state prior to the last Upload." & vbNewLine & "Do you want to continue?", vbOKCancel)
            If MsgBoxVerifyRollback = vbCancel Then
                Exit Sub
            End If
        End If
        Call StatusForm.execMethod("Load", "RollbackLastUpdate")
        'clear out SQL, tablename and upload date so it can't be used again.
        Call removeRollbackData
    End If
End Sub
Sub RollbackLastUpdate()
    ' this takes a CSV string and executes each SQL. The SQL should copy the privs from the original table, rollback to the Clone using swap,and drop the clone
    On Error GoTo ErrorHandlerRollback

    StatusForm.Update_Status ("Rollingback...")
    ' execute all the sql in the string array Load.getRollbackSQL
    Call execSQLFromStringArray(Load.getRollbackSQL)

    ' Get the data
    If UCase(Left(Load.getRollbackSQL, 4)) = "DROP" Then
        StatusForm.Hide
        MsgBox ("Rollback Complete. Table dropped.")
    Else
        On Error GoTo ErrorHandlerGetData
        'Call Query.ExecuteSelectAllFromUploadTable
        StatusForm.Hide
        Call StatusForm.execMethod("Query", "ExecuteSelectAllFromUploadTable")
    End If
    Exit Sub
ErrorHandlerRollback:
    If InStr(err.Description, "Insufficient privileges") Then
        Dim alterSQL As String
        Dim sqlArr() As String
        sqlArr = Split(Load.getRollbackSQL, "~")
        alterSQL = sqlArr(UBound(sqlArr) - 1)
        MsgBox ("Insufficient privileges to perform Rollback. The role you are logged in with must have ownership privileges on the table." & _
        vbNewLine & "To rollback, have the table owner execute this SQL Statement:" & vbNewLine & vbNewLine & _
        alterSQL)
    Else
        Call Utils.handleError("rollback Failed", err)
    End If
    StatusForm.Hide
    Exit Sub
ErrorHandlerGetData:
    Call Utils.handleError("Rollback succeeded but retrieving the data failed.", err)
    StatusForm.Hide
    Exit Sub
End Sub

Sub setUploadDateTime()
    'Sets the date the sheet downloaded data. Used to check if data has changed when uploading
    Dim currentTimeSQL As String
    'get the current time from snowflake
    currentTimeSQL = "select to_char(current_Timestamp,'YYYY-MM-DD HH24:MI:SS.FF')"
    FormCommon.initializeRange("RollbackUploadDateTime") = Format(Utils.execSQLSingleValueOnly(currentTimeSQL), "YYYY-MM-DD HH24:Mmm:SS")
End Sub
Function getUploadDateTime()
    getUploadDateTime = FormCommon.initializeRange("RollbackUploadDateTime")
End Function

Sub setUploadFullyQualifiedTableName()
    FormCommon.initializeRange("RollbackUploadTableName") = tableName
End Sub

Function getUploadFullyQualifiedTableName()
    getUploadFullyQualifiedTableName = FormCommon.initializeRange("RollbackUploadTableName")
End Function

Function setRollbackSQL()
    Dim tempClone As String

    ' In case where the table is new, there will not be a cloned table so just drop the table that was created
    If clonedTable = "NoTableToClone" Then
        RollbackSQL = "Drop table " & tableName
        Exit Function
    End If
    lastQueryID = execSQLReturnSingleValueWithErrorMsgOption(statusWorksheet, Utils.nextStatusCellToLoad, "select last_query_id();", True)
    tempClone = clonedTable
    If tempClone = "" Then
        tempClone = createClonedTableName(tableName)
        RollbackSQL = "create or replace table " & tempClone & " clone " & tableName & _
                    " before (statement =>  '" & lastQueryID & "')~"
    Else
        RollbackSQL = "undrop table " & clonedTable & "~"
    End If
    RollbackSQL = RollbackSQL + grantAllPrivsToClonedTableSQL(tableName, tempClone)
    ' Need to Drop then rename instead of using Swap because swap loses all history - show tables history like <table>
    RollbackSQL = RollbackSQL & "~" & "Drop table " & tableName
    RollbackSQL = RollbackSQL & "~" & "alter table " & tempClone & " rename to " & tableName
    FormCommon.initializeRange("RollbackSQL") = RollbackSQL
End Function

Function getRollbackSQL()
    getRollbackSQL = FormCommon.initializeRange("RollbackSQL")
End Function
Sub removeRollbackData()
    FormCommon.initializeRange("RollbackSQL") = ""
    FormCommon.initializeRange("RollbackUploadTableName") = ""
    FormCommon.initializeRange("RollbackUploadDateTime") = ""
End Sub
Function checkIfTableHasBeenAltered(compareDate As String, fullyQualifiedTable As String)
    Dim lastAlteredSQL As String
    Dim arr() As String
    On Error GoTo ErrorHandlerGeneral
    ' If the date does not exist then return false so upload will continue
    If compareDate = "" Then
        checkIfTableHasBeenAltered = "False"
        Exit Function
    End If
    fullyQualifiedTable = Replace(fullyQualifiedTable, Chr(34), vbNullString, 1, 6)
    arr = Split(fullyQualifiedTable, ".")
    Db = arr(0)
    schema = arr(1)
    table = arr(2)
    'Check if the last_altered data of the table is later than the download date
    lastAlteredSQL = "Select IFF( last_altered > '" & Format(compareDate, "YYYY-MM-DD HH:mm:SS") & "' , 'TRUE' , 'FALSE' ) From " & _
    arr(0) & ".information_schema.tables where table_schema = '" & _
    arr(1) & "' and table_name = '" & arr(2) & "'"
    checkIfTableHasBeenAltered = Utils.execSQLSingleValueOnly(lastAlteredSQL)
    Exit Function
ErrorHandlerGeneral:
    Call Utils.handleError("Error ckeching if table has been altered", err)
    checkIfTableHasBeenAltered = "False"
End Function

Sub detectDateFormat()
    ' Snowflake date formats https://docs.snowflake.com/en/user-guide/date-time-input-output.html#date-formats
    ' Decided not to use this because Excel treats timestamps as a date. It's better to have the user decide the formats
    Dim cellToCheck As range
    Dim dateFormat As String
    Dim dateInputFormat As String

    dateFormat = ""
    numberOfColumns = dataWorksheet.UsedRange.columns.Count
    For i = 1 To numberOfColumns
        Set cellToCheck = dataWorksheet.Cells(iHeaderRow + 1, i)
        If vba.IsDate(cellToCheck) Then
            dateFormat = cellToCheck.NumberFormat
            GoTo dateFound
        End If
    Next i
dateFound:
    If dateFormat <> "" Then

        Select Case True
            Case dateFormat = "m/d/yyyy" Or dateFormat = "mm/dd/yyyy"
                dateInputFormat = "MM/DD/YYYY"
            Case dateFormat = "m/d/yy" Or dateFormat = "mm/dd/yy"
                dateInputFormat = "MM/DD/YY"
            Case dateFormat = "m-d-yyyy" Or dateFormat = "mm-dd-yyyy"
                dateInputFormat = "MM-DD-YYYY"
            Case dateFormat = "yyyy-mm-dd" Or dateFormat = "yyyy-m-d"
                dateInputFormat = "YYYY-MM-DD"
            Case dateFormat = "yy-mm-dd" Or dateFormat = "yy-m-d"
                dateInputFormat = "YY-MM-DD"
            Case dateFormat = "dd-mm-yyyy" Or dateFormat = "d-m-yyyy"
                dateInputFormat = "DD-MM-YYYY"
            Case InStr(dateFormat, "d-mmm-yy")
                dateInputFormat = "DD-MON-YY"
            Case Else
                dateInputFormat = Utils.CustomRange(sgRangeDateInputFormat)
        End Select

        Call Utils.SetDateInputFormat
    End If

End Sub
