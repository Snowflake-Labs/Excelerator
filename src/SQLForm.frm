VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SQLForm 
   Caption         =   "Execute SQL"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8970
   OleObjectBlob   =   "SQLForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SQLForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sqlRange As range
Dim sqlRangeName As range
Dim sqlRangeSQL As range
Dim sqlRangeLastExecutedIndex As range
Dim sqlRangeLastExecutedSQL As range

Dim sqlRangeName_Name As String
Dim sqlRangeSQL_Name As String
'Dim lastSelectedSaveSQLIndex As Integer

'Dim lockRangeTableDate As range

Dim wsWorkbookParams As Worksheet
Dim iSelectedSQLRow As Integer
Dim emptySQLMessage As String
Dim emptySavedSQLMessage As String
'Used for spaces for sql formatting
Dim spcs As String
Dim bInitializing As Boolean

Private Sub btSelectStar_Click()
    tbSQL = "SELECT" + vbCrLf + spcs + "*" + vbCrLf + getFromClause
End Sub

Private Sub btGetColumns_Click()
    Dim selectedColumns As String
    If cbTables.value = "" Then
        MsgBox ("Please select a Table.")
    Else
        Call SelectColumnsForm.initialize(cbDatabases.value, cbSchemas.value, cbTables.value)
        SelectColumnsForm.Show
        selectedColumns = SelectColumnsForm.getSelectedColunms
        If selectedColumns <> "" Then
            tbSQL = "SELECT " + vbCrLf + spcs + selectedColumns & vbCrLf & getFromClause
        End If
    End If
End Sub
Function getFromClause()
    getFromClause = "FROM" + vbCrLf + spcs + _
    """" + cbDatabases.value + """.""" + cbSchemas.value + """.""" + cbTables.value + """"
End Function

Private Sub cbDatabases_Click()
    If Not bInitializing Then
        Call StatusForm.execMethod("FormCommon", "getSchemasCombobox", cbSchemas, cbDatabases.value)
        cbSchemas.ListIndex = 0
    End If
End Sub

Private Sub cbSchemas_Click()
    If Not bInitializing Then
        Call StatusForm.execMethod("FormCommon", "getTablesCombobox", cbTables, cbDatabases.value, cbSchemas.value)
        If cbTables.ListCount > 0 Then
            cbTables.ListIndex = 0
        End If
    End If
End Sub

Private Sub lblGotoSnowflake_Click()
    snowflakeURL = "https://" & Utils.CustomRange(sgRangeServer)
    ActiveWorkbook.FollowHyperlink Address:=snowflakeURL, NewWindow:=True
End Sub

Private Sub tbSQL_Enter()
    If tbSQL = emptySQLMessage Then tbSQL = ""
End Sub

Private Sub UserForm_Activate()
    ' UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()

    Dim sqlRangeLastExecutedIndex_Name As String
    Dim sqlRangeLastExectedSQL_Name As String
    StatusForm.Update_Status ("Preparing SQL Form...")

    Set wsWorkbookParams = Utils.getWorksheet(gsSnowflakeWorkbookParamWorksheetName)

    emptySQLMessage = "Please enter SQL..."
    emptySavedSQLMessage = "No Saved SQL..."
    spcs = "  "
    'Initialization occurs every time form is opened because we set it to Nothing when it closes
    If CustomRange(sgRangeResultsWorksheet) <> "" Then
        Utils.getWorksheet(CustomRange(sgRangeResultsWorksheet)).Activate
    End If
    ' Initialize DB Comboboxes
    bInitializing = True
    Call FormCommon.initializeDBObjectsComboBoxes(cbDatabases, cbSchemas, cbTables)
    bInitializing = False
    'Initialize Table locking ranges
    'lockRangeTableName = FormCommon.initializeRange("LockTableName")
    'Set lockRangeTableDate = FormCommon.initializeRange("LockTableDate")

    sheetCode = ActiveSheet.CodeName
    '********* Saved SQL Ranges **************
    sqlRangeName_Name = sgSavedSQL_Name_RangePrefix & sheetCode
    sqlRangeSQL_Name = sgSavedSQL_SQL_RangePrefix & sheetCode
    sqlRangeLastExecutedIndex_Name = sgSavedSQL_SelectedIndex_RangePrefix & sheetCode
    sqlRangeLastExectedSQL_Name = sgSavedSQL_LastExecutedSQL_RangePrefix & sheetCode

    On Error Resume Next
    Set sqlRangeName = Utils.CustomRange(sqlRangeName_Name)
    Set sqlRangeSQL = Utils.CustomRange(sqlRangeSQL_Name)
    cbSQLList.RowSource = sqlRangeName_Name
    On Error GoTo 0
    err.Clear
    'Initialize ranges for Last saved search and last query
    Set sqlRangeLastExecutedIndex = Utils.getOrCreateRange(wsWorkbookParams, sqlRangeLastExecutedIndex_Name, igAllSingleValueRages_ColNumber)
    Set sqlRangeLastExecutedSQL = Utils.getOrCreateRange(wsWorkbookParams, sqlRangeLastExectedSQL_Name, igAllSingleValueRages_ColNumber)
    'set combox and sql text with last used values
    'lastSelectedSaveSQLIndex = sqlRangeLastExecutedIndex
    If cbSQLList.ListCount > 0 Then
        If sqlRangeLastExecutedIndex.value <> "" Then
            If sqlRangeLastExecutedIndex < cbSQLList.ListCount Then
                cbSQLList.ListIndex = sqlRangeLastExecutedIndex
            End If
        End If
    Else
        cbSQLList.value = emptySavedSQLMessage
    End If
    tbSQL = sqlRangeLastExecutedSQL
    If tbSQL = "" Then
        tbSQL = emptySQLMessage
        '   btExit.SetFocus
    End If

    tbSQL.MultiLine = True
    SQLForm.Update_Status ("")
    StatusForm.Hide
End Sub

Function assignItemToVariableFromCollection(ByRef var As Variant, col As Collection, key As String)
    'Attempts to assign an item from collection based on a key and returns true if it succeeds and false if it doesn't
    On Error Resume Next
    var = col(key)
    assignItemToVariableFromCollection = (err.Number = 0)
    err.Clear
End Function

Sub tbSQL_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If cbSQLList.ListIndex = -1 Then Exit Sub
    ' sqlRangeSQL(cbSQLList.listIndex + 1, 1) = tbSQL
End Sub

Function getNextEmptyColumn(ws As Worksheet)
    Dim i As Integer
    Dim bFoundEmpty As Boolean
    Dim checkCell As range

    With ws
        ' Loops until it finds the next epmty cell in row 1
        For i = 1 To 10000
            Set checkCell = .Cells(1, i)
            If checkCell = "" Then
                On Error Resume Next
                nm = ""
                nm = Utils.getRangeNameIgnoreError(checkCell)
                If nm = "" Then
                    getNextEmptyColumn = i
                    Exit Function
                End If
                On Error GoTo ErrorHandlerGeneral
            End If
        Next i
    End With
    MsgBox ("There is a problem getting the next empty column on sheet " & ws.name & ".")
    err.Raise 2000

ErrorHandlerGeneral:
    MsgBox (err.Description)
End Function

Sub addRowToRange(rangeName As String)
    Dim sqlRange As range
    Dim nextRow As Integer
    Dim nextColumn As Integer

    nextRow = 1
    colNumber = 0
    On Error GoTo CreateRange
    Set sqlRange = Utils.CustomRange(rangeName)
    nextRow = sqlRange.Rows.Count + 1
    colNumber = sqlRange.Column
CreateRange:
    On Error GoTo 0
    If colNumber = 0 Then
        colNumber = getNextEmptyColumn(wsWorkbookParams)
    End If
    With wsWorkbookParams
        ActiveWorkbook.Names.Add name:=rangeName, _
        RefersTo:=.range(.Cells(1, colNumber), .Cells(nextRow, colNumber))
    End With
End Sub

Public Sub ShowForm()
    Me.Show
End Sub

Public Sub Update_Status(sStatus)
    tbStatusIndicator = sStatus
    DoEvents
End Sub

Private Sub btRemoveSQL_Click()
    'Selection.Delete Shift:=xlUp
    Dim row As Integer
    ' If nothing is selected then bail
    If cbSQLList.ListIndex = -1 Then Exit Sub
    'get select row
    row = cbSQLList.ListIndex + 1
    ListCount = sqlRangeName.Rows.Count
    ' cbSQLList.RemoveItem (row - 1)
    wsWorkbookParams.range(sqlRangeName(row, 1), sqlRangeName(row, 1)).Delete Shift:=xlUp
    wsWorkbookParams.range(sqlRangeSQL(row, 1), sqlRangeSQL(row, 1)).Delete Shift:=xlUp
    If ListCount = 1 Then
        cbSQLList.value = ""
    Else

        If row = ListCount Then
            cbSQLList.ListIndex = row - 2
        Else
            cbSQLList.ListIndex = row - 1
        End If

        ' cbSQLList.value = sqlRange.Cells(iSelectedSQLRow, 1)
        tbSQL = sqlRangeSQL.Cells(cbSQLList.ListIndex + 1, 1)
        cbSQLList.RowSource = sqlRangeName_Name
    End If

End Sub

Sub btSaveSQL_Click()
    Dim index As Integer
    'Check to see if SQL is empty of the default message
    If tbSQL = "" Or tbSQL = emptySQLMessage Then
        MsgBox ("Please enter a SQL statement before saving.")
        tbSQL = emptySQLMessage
        Exit Sub
    End If
    'Check if the SQL is already in list, if it is then save it, if not create a new one
    If cbSQLList.ListIndex > -1 Then
        index = cbSQLList.ListIndex
        sqlRangeSQL.Cells(index + 1, 1) = tbSQL
        sqlRangeName.Cells(index + 1, 1) = Replace(tbSQL, vbCrLf, " ")
        'This updates the combobox
        cbSQLList.RowSource = sqlRangeName_Name
        cbSQLList.ListIndex = index
        Exit Sub
    End If
    ' Must be new so add to ranges and save value
    'Start with Name
    Call addRowToRange(sqlRangeName_Name)
    Set sqlRangeName = Utils.CustomRange(sqlRangeName_Name)
    sqlRangeName.Cells(sqlRangeName.Rows.Count, 1) = Replace(tbSQL, vbCrLf, " ")
    'Now the SQL itself
    Call addRowToRange(sqlRangeSQL_Name)
    Set sqlRangeSQL = Utils.CustomRange(sqlRangeSQL_Name)
    sqlRangeSQL.Cells(sqlRangeSQL.Rows.Count, 1) = tbSQL
    'This updates the combobox
    cbSQLList.RowSource = sqlRangeName_Name
    'set the correct item in the combobox
    cbSQLList.ListIndex = sqlRangeName.Rows.Count - 1
End Sub
Private Sub btNewSQL_Click()
    cbSQLList.ListIndex = -1
    Call btSaveSQL_Click
    ' tbSQL = ""
End Sub

Private Sub btRenameSQL_Click()
    Dim name As String
    If cbSQLList.ListIndex = -1 Then
        MsgBox ("Please select a Saved Search or create one.")
        Exit Sub
    End If
    GetValueForm.setMessage ("Enter Search Name:")
    GetValueForm.setValue (cbSQLList.value)
    GetValueForm.Show
    name = GetValueForm.Getvalue
    If name <> "" Then
        sqlRangeName.Cells(cbSQLList.ListIndex + 1, 1) = name
        ' cbSQLList.value = name
    End If
End Sub
'**************************** Drop Down List ****************************
Sub cbSQLList_click()
    On Error Resume Next
    tbSQL = sqlRangeSQL.Cells(cbSQLList.ListIndex + 1, 1)
End Sub
Private Sub cbSQLList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'This procedure is only need when the user changes the SQL and wants to get back to the original
    On Error Resume Next
    tbSQL = sqlRangeSQL.Cells(cbSQLList.ListIndex + 1, 1)
End Sub

Public Sub ExecuteButton_Click()
    Dim lastAlteredDate As String
    If worksheetBelongsToAddin Then
        Exit Sub
    End If
    If tbSQL = "" Or tbSQL = emptySQLMessage Then
        MsgBox ("Please enter a SQL statement before executing.")
        tbSQL = emptySQLMessage
        Exit Sub
    End If
    Me.Hide
    Utils.SaveAllNamedRangesToAddIn

    Set StatusForm = Nothing
    Call StatusForm.execMethod("Query", "ExecuteSQLFromNamedCell", tbSQL)
    'Check to see if anything was returned. If nothing then it could be an error so leave the window open
    If ActiveSheet.range(gsQueryResultsCell) = "" Then
        Me.Show
    End If
    'This is a hack because after an excel table is created and the selected cell is in the table,
    'the ribbon changes tabs to the table, so this brings it back to the Home tab
    Utils.RibbonactivateHomeTab
    If cbSQLList.ListIndex > -1 Then
        sqlRangeLastExecutedIndex = cbSQLList.ListIndex     'lastSelectedSaveSQLIndex
    End If
    If Not sqlRangeLastExecutedSQL Is Nothing Then
        sqlRangeLastExecutedSQL = tbSQL.value
    End If

    Call FormCommon.saveDBObjectsValues(cbDatabases, cbSchemas, cbTables)
End Sub

Private Sub btExit_Click()
    Me.Hide
    sqlRangeLastExecutedIndex = cbSQLList.ListIndex
    sqlRangeLastExecutedSQL = tbSQL.value
End Sub

Private Sub iHelpLink_Click()
    OpenHelp ("ConfigForm")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call btExit_Click
End Sub
