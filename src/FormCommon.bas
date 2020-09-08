Attribute VB_Name = "FormCommon"
'DB Objects Ranges
Dim dbObjRangeSelectedDB As range
Dim dbObjRangeSelectedSchema As range
Dim dbObjRangeSelectedTable As range

'worksheet for parameters
Dim wsWorkbookParams As Worksheet
'Holds cached DB, Schema and table and columns
Dim dictTables As New Scripting.Dictionary ' holds all tables per db.schema
Dim dictColumns As New Scripting.Dictionary ' holds all columns per db.schema
Dim dictSchemas As New Scripting.Dictionary ' holds all schema per db
Dim arrDatabases As Variant
Dim bInitializing As Boolean

Sub dropDBObjectsCache()
    arrDatabases = Empty
    dictSchemas.RemoveAll
    dictTables.RemoveAll
    dictColumns.RemoveAll
End Sub
Sub dropDBObjectsTableCache()
    dictTables.RemoveAll
    dictColumns.RemoveAll
End Sub
Sub dropDBObjectsColumnCache()
    dictColumns.RemoveAll
End Sub

Sub initializeDBObjectsComboBoxes(ByRef cbDatabases As comboBox, ByRef cbSchemas As comboBox, ByRef cbTables As comboBox)
    bInitializing = True
    'Get the CodeName for the worksheet - Need to do this craziness because the codename is not available when not in the debugger until you execute the below lines
    'Also I have to trap for errors because if someone renames the sheet before running a query it will through an error, but the codeName will then be available.
    sheetCode = ActiveSheet.CodeName
    If sheetCode = "" Then
        On Error Resume Next
        sheetCode = ActiveWorkbook.VBProject.VBComponents(ActiveSheet.name).Properties("Codename")
        sheetCode = ActiveSheet.CodeName
        On Error GoTo 0
    End If

    'Get the worksheet that holds the params
    Set wsWorkbookParams = Utils.getWorksheet(gsSnowflakeWorkbookParamWorksheetName)

    '********* DB Object Comboboxes **************
    Set dbObjRangeSelectedDB = FormCommon.initializeRange("DBCombobox")
    Set dbObjRangeSelectedSchema = FormCommon.initializeRange("SchemaCombobox")
    Set dbObjRangeSelectedTable = FormCommon.initializeRange("TableCombobox")

    ' *********** databases Combobox  ****************
    Call FormCommon.getDatabasesCombobox(cbDatabases)
    If cbDatabases.ListCount = 0 Then
        MsgBox ("User does not have access to any databases")
        Exit Sub
    End If
    'If there isn't one saved then take the default DB
    If dbObjRangeSelectedDB = "" Then
        dbObjRangeSelectedDB = CustomRange(sgRangeDefaultDatabase)
    End If
    'If the DB is in the list then set the index, else set it to the first one
    index = FormCommon.indexOfValueInList(cbDatabases, dbObjRangeSelectedDB.value)
    If index > -1 Then
        cbDatabases.ListIndex = index
    Else
        cbDatabases.ListIndex = 0
    End If

    ' *********** Schema Combobox  ****************

    Call FormCommon.getSchemasCombobox(cbSchemas, cbDatabases.value)
    'If not save already then take the default
    If dbObjRangeSelectedSchema = "" Then
        dbObjRangeSelectedSchema = CustomRange(sgRangeDefaultSchema) 'UCase(CustomRange(sgRangeDefaultSchema))
    End If
    ' Select schema from list
    index = FormCommon.indexOfValueInList(cbSchemas, dbObjRangeSelectedSchema.value)
    If index > -1 Then
        cbSchemas.ListIndex = index
    Else
        If cbSchemas.ListCount > 0 Then
            cbSchemas.ListIndex = 0
        End If
    End If

    ' *********** Table Combobox  ****************
    Call FormCommon.getTablesCombobox(cbTables, cbDatabases.value, cbSchemas.value)
    Call FormCommon.setCombboxValue(cbTables, dbObjRangeSelectedTable.value)
    bInitializing = False
End Sub

Sub saveDBObjectsValues(ByRef cbDatabases As comboBox, ByRef cbSchemas As comboBox, ByRef cbTables As comboBox)
    dbObjRangeSelectedDB = cbDatabases.value
    dbObjRangeSelectedSchema = cbSchemas.value
    dbObjRangeSelectedTable = cbTables.value
End Sub
Function getDatabase()
    Dim rangeName As String
    rangeName = sgDBObj_LastSelectedDB_RangePrefix & ActiveSheet.CodeName
    getDatabase = Utils.CustomRange(rangeName)
End Function
Function getSchema()
    Dim rangeName As String
    rangeName = sgDBObj_LastSelectedSchema_RangePrefix & ActiveSheet.CodeName
    getSchema = Utils.CustomRange(rangeName)
End Function

Function getTable()
    Dim rangeName As String
    rangeName = sgDBObj_LastSelectedTable_RangePrefix & ActiveSheet.CodeName
    getTable = Utils.CustomRange(rangeName)
End Function

Function getMergeKeys()
    Dim rangeName As String
    rangeName = sgUploadMergeKeys_RangePrefix & ActiveSheet.CodeName
    On Error GoTo NoMergeKey
    getMergeKeys = Utils.CustomRange(rangeName)
    Exit Function
NoMergeKey:
    getMergeKeys = ""
End Function

Function getFullyQualifiedTable()
    getFullyQualifiedTable = """" + getDatabase + """.""" + getSchema + """.""" + getTable + """"
End Function

Function initializeRange(field As String)
    Dim uploadTypeRange_Name As String
    Select Case field
        Case "UploadType"
            Prefix = sgUploadType_RangePrefix
        Case "MergeKeysNumbers"
            Prefix = sgUploadMergeKeys_RangePrefix
        Case "MergeKeysLetters"
            Prefix = sgUploadMergeKeysByLetters_RangePrefix
        Case "DBCombobox"
            Prefix = sgDBObj_LastSelectedDB_RangePrefix
        Case "SchemaCombobox"
            Prefix = sgDBObj_LastSelectedSchema_RangePrefix
        Case "TableCombobox"
            Prefix = sgDBObj_LastSelectedTable_RangePrefix
        Case "LockTableDate"
            Prefix = sgLockedDownloadTableDateTime_RangePrefix
        Case "RollbackUploadDateTime"
            Prefix = sgRollbackUploadDate_RangePrefix
        Case "RollbackUploadTableName"
            Prefix = sgRollbackUploadTableName_RangePrefix
        Case "RollbackSQL"
            Prefix = sgRollbackSQL_RangePrefix
    End Select
    Set initializeRange = Utils.getOrCreateRange(wsWorkbookParams, Prefix & ActiveSheet.CodeName, igAllSingleValueRages_ColNumber)
End Function

Sub getTablesCombobox(ByRef cbTables As comboBox, database As String, schema As String)
    Dim sql As String
    Dim arrTables As Variant
    Dim key As String
    Call StatusForm.Update_Status("Getting Tables...")
    key = database + "-" + schema
    cbTables.Clear
    'Get the table arrays for the key - DB + schema. If it doesn't exsist get it
    On Error GoTo ErrorHandlerDone
    If dictTables.Exists(key) Then
        arrTables = dictTables(key)
    Else
        sql = "select table_name from """ & database & """.information_schema.tables where table_schema = '" & schema & "'"
        arrTables = Utils.execSQLToArray(sql)
        dictTables.Add Item:=arrTables, key:=key
    End If

    On Error Resume Next ' Doing this because the array could be empty
    'arrTables is a 2 dimensional array, with Columns being the first and rows the second
    On Error GoTo ErrorHandlerDone
    For i = LBound(arrTables) To UBound(arrTables, 2)
        cbTables.AddItem (arrTables(0, i))
    Next i
    cbTables.ListIndex = 0
ErrorHandlerDone:
    If Not bInitializing Then
        StatusForm.Hide
    End If
End Sub

Sub getDatabasesCombobox(ByRef cbDatabases As comboBox)
    Dim sql As String
    Call StatusForm.Update_Status("Getting Databases...")
    If IsEmpty(arrDatabases) Then
        sql = "show databases"
        Utils.execSQLFireAndForget (sql)
        sql = "WITH databases (a,name,b,c,d,e,f,g,h) as (select * from table(result_scan(last_query_id()))) " & _
                "select name from databases"
        arrDatabases = Utils.execSQLToArray(sql)
    End If
    For i = LBound(arrDatabases) To UBound(arrDatabases, 2)
        cbDatabases.AddItem (arrDatabases(0, i))
    Next i

End Sub

Sub getSchemasCombobox(ByRef cbSchemas As comboBox, database As String)
    Dim sql As String
    Dim arrSchemas As Variant
    Dim key As String

    Call StatusForm.Update_Status("Getting Schemas...")
    key = database
    cbSchemas.Clear
    'Get the table arrays for the key - DB + schema. If it doesn't exsist get it
    If dictSchemas.Exists(key) Then
        arrSchemas = dictSchemas(key)
    Else
        On Error Resume Next
        sql = "select schema_name from """ & database & """.information_schema.schemata "
        arrSchemas = Utils.execSQLToArray(sql)
        dictSchemas.Add Item:=arrSchemas, key:=key
        On Error GoTo 0
    End If

    'arrTables is a 2 dimensional array, with Columns being the first and rows the second
    If Not IsEmpty(arrSchemas) Then
        For i = LBound(arrSchemas) To UBound(arrSchemas, 2)
            cbSchemas.AddItem (arrSchemas(0, i))
        Next i
    End If
    If Not bInitializing Then
        StatusForm.Hide
    End If
End Sub

Function isValueInList(list As comboBox, value As String)
    Dim i As Integer
    For i = 0 To list.ListCount - 1
        If UCase(list.list(i)) = UCase(value) Then
            isValueInList = True
            Exit Function
        End If
    Next
    isValueInList = False
End Function
Function indexOfValueInList(list As comboBox, value As String)
    Dim i As Integer
    For i = 0 To list.ListCount - 1
        If UCase(list.list(i)) = UCase(value) Then
            indexOfValueInList = i
            Exit Function
        End If
    Next
    indexOfValueInList = -1
End Function

Function setCombboxValue(ByRef cb As comboBox, value As String)
    If cb.ListCount = 0 Then Exit Function

    If FormCommon.isValueInList(cb, value) Then
        cb.value = value
    Else
        cb.ListIndex = 0
    End If
End Function

Public Function getColumnArray(database As String, schema As String, table As String)
    Dim sql As String
    Dim arrColumns As Variant

    key = database + "-" + schema + "-" + table

    'Get the table arrays for the key - DB + schema. If it doesn't exsist get it
    If dictColumns.Exists(key) Then
        getColumnArray = dictColumns(key)
    Else
        sql = "select column_name, data_type from """ & database & """.information_schema.columns where table_schema = '" & _
        schema & "' and table_name = '" & table & "' order by ordinal_position"
        getColumnArray = Utils.execSQLToArray(sql)
        dictColumns.Add Item:=getColumnArray, key:=key
    End If
End Function

