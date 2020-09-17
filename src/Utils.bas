Attribute VB_Name = "Utils"
Dim msConnectionStatus As String
Dim msRibbon As IRibbonUI
Dim mdConnections As New Scripting.Dictionary
Dim bSecondPass As Boolean ' used to allow execution of sql inside the open connection sub without infinite recursion


Sub Connect()
    On Error GoTo ErrorHandlerConnection
    Set conn = Utils.getOpenDBConncetion(True)
ErrorHandlerConnection:
End Sub

Function ConnectionStringDSNLess()
    Dim connectionType As String
    Set LoginForm = Nothing
    LoginForm.Show
    If Not LoginForm.bLoginOK Then
        err.Raise giCancelEvent
    End If
    #If Mac Then
    connectionType = "ODBC;"
    #Else
    connectionType = "Provider=MSDASQL.1;"
    #End If
    Dim dateInputFormat As String

    ConnectionStringDSNLess = connectionType & "driver={SnowflakeDSIIDriver};server=" & LoginForm.tbServer & ";database=" & _
       LoginForm.tbDatabase & ";schema=" & LoginForm.tbSchema & ";warehouse=" & LoginForm.tbWarehouse & ";role=" & _
       LoginForm.tbRole & ";Uid=" & LoginForm.tbUserID & ";CLIENT_SESSION_KEEP_ALIVE=true;"

    If LoginForm.rbSSO Then
        ConnectionStringDSNLess = ConnectionStringDSNLess & "Authenticator=externalbrowser;"
    Else
        ConnectionStringDSNLess = ConnectionStringDSNLess & "Pwd=" & LoginForm.tbPassword & ";"
    End If
    msConnectionStatus = "Server: " & LoginForm.tbServer & "    Database: " & LoginForm.tbDatabase & "    Schema: " & LoginForm.tbSchema & _
        "    Warehouse: " & LoginForm.tbWarehouse & "    User: " & LoginForm.tbUserID & "    Role: " & LoginForm.tbRole

    ' Consider this: "CLIENT_SESSION_KEEP_ALIVE", "true"  If not the token expires in 4 hours
End Function

Function ConnectionString()
    'Not used anymore
    ConnectionString = "ODBC;DSN=Snowflake" & ";database=" & CustomRange(sgRangeDefaultDatabase).value & ";schema=" & CustomRange(sgRangeDefaultSchema) & _
                                ";warehouse=" & CustomRange(sgRangeWarehouse) & ";role=" & CustomRange(sgRangeRole) & ";" & _
                                "Uid=" & CustomRange(sgRangeUserID) & ";  "

    If CustomRange(sgRangeAuthType) = "SSO" Then
        ConnectionString = ConnectionString & "Authenticator=externalbrowser;"
    Else
        ConnectionString = ConnectionString & "Pwd=" & LoginForm.tbPassword & ";"
    End If
End Function

Sub RemoveQueryTables(ws As Worksheet)
    Dim qt As QueryTable
    For Each qt In ws.QueryTables
        qt.Delete
    Next qt
End Sub
Sub openConnection(ByRef mDBConnection As ADODB.Connection, connString As String)
    On Error GoTo ErrorHandlerConnection
    StatusForm.Update_Status ("Authenticating...")
    mDBConnection.Open connString
    StatusForm.Update_Status ("Complete...")
    StatusForm.Hide
    Exit Sub
ErrorHandlerConnection:
    If err.Number <> giCancelEvent Then
        MsgBox "ERROR: Problem connecting: " & err.Description
        StatusForm.Hide
    End If
End Sub
Function openLogin()
    Dim connString As String
    Dim mDBConnection As ADODB.Connection
    Dim bConnected As Boolean
    Dim bEverythingOKWithConncetion As Boolean
    Dim arrDBObjs() As String
    Dim sql As String

    On Error GoTo ErrorHandlerConnection
    'Opens login form to get connection string. If the user cancels, an error is raised
    Do While Not bEverythingOKWithConncetion
        bConnected = False
        Do While Not bConnected
            connString = Utils.ConnectionStringDSNLess()
            If connString <> "" Then
                Set mDBConnection = New ADODB.Connection
                mDBConnection.CommandTimeout = 120
                mDBConnection.ConnectionTimeout = 120
                Set StatusForm = Nothing
                Call StatusForm.execMethod("Utils", "openConnection", mDBConnection, connString)
                If mDBConnection.State > 0 Then
                    bConnected = True
                End If
            End If
        Loop
        If mdConnections.Exists(ActiveWorkbook.name) Then
            mdConnections.Remove (ActiveWorkbook.name)
        End If
        mdConnections.Add key:=ActiveWorkbook.name, Item:=mDBConnection
        If Not bSecondPass Then ' this is to make sure there is no infinite recursion
            bSecondPass = True
            Call Utils.SetDateInputFormat
            bSecondPass = False
        End If
        Dim checkDatabase As String
        Dim checkSchema As String
        bEverythingOKWithConncetion = True
        'Check if warehous, DB ans schema are valid
        sql = "select IFNULL(current_warehouse(),'') ||','|| IFNULL(current_database(),'') ||','|| IFNULL(current_schema(),'')"
        dbObjsString = execSQLSingleValueOnly(sql)
        arrDBObjs = Split(dbObjsString, ",")
        If arrDBObjs(0) = "" Then
            MsgBox ("The use does not have a valid warehouse. " & _
           "Please specify one and login again.")
            bEverythingOKWithConncetion = False
        Else
            If arrDBObjs(1) = "" Then
                MsgBox ("The database entered does not exist or the user does not have access to it. " & _
                "If the database name is in mixed case then use double quotes.")
                bEverythingOKWithConncetion = False
            Else
                If arrDBObjs(2) = "" Then
                    MsgBox ("The schema entered does not exist or the user does not have access to it " & _
                    "If the schema name is in mixed case then use double quotes.")
                    bEverythingOKWithConncetion = False
                End If
            End If
        End If
    Loop
    Set openLogin = mDBConnection
    Exit Function
ErrorHandlerConnection:
    err.Raise giCancelEvent
End Function

Public Function getOpenDBConncetion(reauthenticate As Boolean)
    Dim mDBConnection As ADODB.Connection
    Dim bGetConnection As Boolean
    ' This collection of connections is to handle multiple excel workbooks being open at one time.
    If Not mdConnections.Exists(ActiveWorkbook.name) Or reauthenticate Then
        On Error GoTo ErrorHandlerConnection
        Set mDBConnection = openLogin
    Else
        Set mDBConnection = mdConnections(ActiveWorkbook.name)
    End If

    If mDBConnection.State = 0 Then
        mDBConnection.Open
    End If
    Set getOpenDBConncetion = mDBConnection

    Exit Function
ErrorHandlerConnection:
    err.Raise giCancelEvent
End Function

Function login()
    On Error GoTo ErrorHandlerConnection
    Set conn = getOpenDBConncetion(False)
    login = True
    Exit Function
ErrorHandlerConnection:
    login = False
End Function

Sub ExecSQL(ws As Worksheet, dest As String, sqlString As String)
    Dim connString As String
    Debug.Print "sqlString = " & sqlString
    #If Mac Then

    ' connString = Utils.ConnectionStringDSNLess()
    connString = Utils.ConnectionString()
    On Error GoTo 0
    With ws.QueryTables.Add(Connection:=connString, Destination:=ws.range(dest), sql:=sqlString)
        .BackgroundQuery = False
        .Refresh
        .name = "qt1"
    End With
    #Else
    Dim rs As Recordset
    Dim cTimeCols As New Collection ' this is used to track all the time col so it can be formatted properly
    Dim cTimeStampCols As New Collection ' tracks all the timestamp cols

    On Error GoTo ErrorHandlerConnection
    Set conn = Utils.getOpenDBConncetion(False)
    If conn Is Nothing Then
        Exit Sub
    End If
    'Execute query
    On Error GoTo ErrorHandlerExecSQL
    Set rs = conn.Execute(sqlString)
    'get column headers
    On Error GoTo ErrorHandlerNoRows
    If rs.Fields.Count = 0 Then
        Exit Sub
    End If
    On Error GoTo 0  ' Maybe improve error handling here

    For intColIndex = 0 To rs.Fields.Count - 1
        ws.range(dest).Offset(0, intColIndex) = rs.Fields(intColIndex).name
        ' Build collection of Time and TimeStamp columns so they can be formatted properly later
        Select Case rs.Fields(intColIndex).Type
            Case adDBTime
                cTimeCols.Add (intColIndex)
            Case adDBTimeStamp
                cTimeStampCols.Add (intColIndex)
        End Select
    Next

    'Get results of query
    ws.range(dest).Offset(1, 0).CopyFromRecordset rs
    'Close Recordset, no need to close Connection
    rs.Close

    'get the last row
    LastRowIndex = ws.UsedRange.row - 1 + ws.UsedRange.Rows.Count
    ' for each column that is a time, set the format
    For Each col In cTimeCols
        ws.range(Cells(1, col + 1), Cells(LastRowIndex, col + 1)).NumberFormat = "h:mm:ss"
    Next
    ' for each column that is a timestamp, set the format
    For Each col In cTimeStampCols
        ws.range(Cells(1, col + 1), Cells(LastRowIndex, col + 1)).NumberFormat = "m/d/yyyy h:mm:ss"
    Next
    If "ERROR:" = Left(ws.range(dest).Offset(1, 0), 6) Then
        err.Description = ws.range(dest).Offset(1, 0)
        GoTo ErrorHandlerExecSQL
    End If
    Exit Sub
ErrorHandlerExecSQL:
    Call Utils.handleError(" Exceuting SQL: " & sqlString & vbNewLine, err)
    On Error GoTo 0
    'err.Raise 2000, err.Description
    err.Raise giSQLErrorEvent
ErrorHandlerConnection:
    If err.Number = giCancelEvent Then
        err.Raise giCancelEvent
    End If
    Exit Sub
ErrorHandlerNoRows:
    Exit Sub 'Ignore error
    #End If
End Sub

Function execSQLReturnSingleValue(ws As Worksheet, dest As String, sqlString As String)
    execSQLReturnSingleValue = execSQLReturnSingleValueWithErrorMsgOption(ws, dest, sqlString, True)
End Function
Function execSQLReturnSingleValueNoErrorMsg(ws As Worksheet, dest As String, sqlString As String)
    execSQLReturnSingleValueNoErrorMsg = execSQLReturnSingleValueWithErrorMsgOption(ws, dest, sqlString, False)
End Function

Function execSQLReturnSingleValueWithErrorMsgOption(ws As Worksheet, dest As String, sqlString As String, displayErrorMsg As Boolean)
    Dim rs As Recordset
    'Execute query
    On Error GoTo ErrorHandlerExecSQL
    Set conn = getOpenDBConncetion(False)
    Set rs = conn.Execute(sqlString)

    execSQLReturnSingleValueWithErrorMsgOption = rs.Fields(0)
    ws.range(dest).Offset(0, 0) = rs.Fields(0).name
    ws.range(dest).Offset(1, 0) = execSQLReturnSingleValueWithErrorMsgOption

    rs.Close
    Exit Function
ErrorHandlerExecSQL:
    If displayErrorMsg Then
        MsgBox ("Error exceuting SQL: " & err.Description)
    End If
    err.Raise 2000
End Function
Function execSQLReturnConcatResults(sqlString As String, delimeter As String)
    Dim rs As Recordset
    execSQLReturnConcatResults = ""
    'Execute query
    On Error GoTo ErrorHandlerExecSQL
    Set conn = getOpenDBConncetion(False)
    Set rs = conn.Execute(sqlString)

    Do While Not rs.EOF
        execSQLReturnConcatResults = execSQLReturnConcatResults & delimeter & rs.Fields(0)
        rs.MoveNext
    Loop

    If execSQLReturnConcatResults <> "" Then
        'remove leading delimiter
        execSQLReturnConcatResults = Right(execSQLReturnConcatResults, Len(execSQLReturnConcatResults) - 1)
    End If
    rs.Close
    Exit Function
ErrorHandlerExecSQL:
    err.Raise 2000
End Function
Function execSQLToArray(sqlString As String) As Variant()
    'Execute query
    On Error GoTo ErrorHandlerExecSQL
    Set conn = getOpenDBConncetion(False)
    Set rs = conn.Execute(sqlString)

    If Not rs.EOF Then
        execSQLToArray = rs.GetRows
    End If
    rs.Close
    Exit Function
ErrorHandlerExecSQL:
    err.Raise 2000
End Function
Sub execSQLFireAndForget(sqlString As String)
    Dim rs As Recordset
    'Execute query
    On Error GoTo ErrorHandlerExecSQL
    Set conn = getOpenDBConncetion(False)
    Set rs = conn.Execute(sqlString)
    Exit Sub
ErrorHandlerExecSQL:
    If err.Number = giCancelEvent Then
        err.Raise giCancelEvent
    Else
        'MsgBox ("Error exceuting SQL: " & err.Description)
        err.Raise giSQLErrorEvent
    End If
End Sub
Function execSQLSingleValueOnly(sqlString As String)

    'Execute query
    On Error GoTo ErrorHandlerExecSQL
    Set conn = getOpenDBConncetion(False)
    Set rs = conn.Execute(sqlString)

    If Not rs.EOF Then
        execSQLSingleValueOnly = rs.Fields(0)
        If IsNull(execSQLSingleValueOnly) Then
            execSQLSingleValueOnly = ""
        End If
    Else
        execSQLSingleValueOnly = ""
    End If

    rs.Close
    Exit Function
ErrorHandlerExecSQL:
    err.Raise 2
End Function
Function execSQLFromStringArray(sqlString As String)
    Dim sqlArr() As String

    sqlArr = Split(sqlString, "~")
    For i = LBound(sqlArr) To UBound(sqlArr)
        If sqlArr(i) <> "" Then
            Utils.execSQLFireAndForget (sqlArr(i))
        End If
    Next i
End Function


Sub createTableForAllDataOnWorksheet(ws As Worksheet)
    #If Mac Then
    Dim qt As QueryTable
    Set qt = ws.QueryTables(1)
    Set TblRng = qt.ResultRange
    Set tableListObj = ws.ListObjects.Add(xlSrcRange, TblRng, , xlYes)  '.Name = "Table1"
    #Else
    Dim UsedRng As range
    Set UsedRng = ws.UsedRange
    If UsedRng.Cells(1, 1) <> "" Then
        'LastRowIndex = UsedRng.Row - 1 + UsedRng.Rows.Count
        ws.ListObjects.Add(xlSrcRange, UsedRng, , xlYes).name = ws.CodeName
    End If
    #End If
End Sub

Function nextStatusCellToLoad()
    Dim statusWorksheet As Worksheet
    Set statusWorksheet = Sheets(CustomRange(sgRangeLogWorksheet).value)
    i = statusWorksheet.range(sgLastCellOnLogWS).End(xlUp).row
    nextStatusCellToLoad = "A" & i + 1
End Function

Function lastPopulatedCell()
    Dim statusWorksheet As Worksheet
    Set statusWorksheet = Sheets(CustomRange(sgRangeLogWorksheet))
    i = statusWorksheet.CustomRange(sgLastCellOnLogWS).End(xlUp).row
    lastPopulatedCell = "A" & i
End Function

Sub handleError(message As String, err As ErrObject)
    Debug.Print err.Description
    MsgBox "ERROR: " & message & vbNewLine & err.Description
End Sub

Function getWorksheet(wsName As String)
    ' returns worksheet based on name and creates it if it dosen't exist
    On Error Resume Next
    Set getWorksheet = Worksheets(wsName)
    'If it doesn't exist then create it.
    If err.Number = 9 Then
        Dim sActiveWorksheet As String
        sActiveWorksheet = ActiveSheet.name
        Set getWorksheet = Worksheets.Add(After:=Sheets(Worksheets.Count))
        'Activate the worksheet that was active originally
        If wsName = Utils.CustomRange(sgRangeLogWorksheet) Or wsName = gsSnowflakeWorkbookParamWorksheetName Then
            Worksheets(getWorksheet.name).Visible = False
        End If
        Worksheets(sActiveWorksheet).Activate
        getWorksheet.name = wsName
    End If
End Function

Function checkStoredProcCompatibility(ws As Worksheet)
    Dim spVers As String
    Dim spMaxVers As Integer
    Dim spMinVers As Integer
    Dim spMinorVers As Integer
    Dim workbookVerMsg As String
    Dim msg As String
    On Error GoTo ErrorHandlerSPDoesNoteExist
    checkStoredProcCompatibility = False
    spMinorVers = 0
    spVers = Utils.execSQLReturnSingleValueNoErrorMsg(ws, Utils.nextStatusCellToLoad, "call get_stored_proc_version_number()")
    spMaxVers = Split(spVers, ".")(0)
    spMinVers = Split(spVers, ".")(1)

    workbookVerMsg = "Workbook compatibility version: " & gWorkbookSPCompatibilityVers
    ' Check if Stored Procs are outdated
    If gWorkbookSPCompatibilityVers > spMaxVers Then
        msg = "The stored procedures are outdated. If you would like to use the 'Auto-generate Data Type' feature, " & _
        "please have an administrator upgrade the stored procedures to the latest version."
        MsgBox (msg)
        ws.range(nextStatusCellToLoad) = msg
        Exit Function
    End If
    ' Check if workbook is outdated
    If gWorkbookSPCompatibilityVers < spMinVers Then
        msg = "This workbook is outdated. If you would like to use the 'Auto-generate Data Type' feature, " & _
        "please have an administrator upgrade the stored procedures to the latest version."
        MsgBox (msg)
        ws.range(nextStatusCellToLoad) = msg
        Exit Function
    End If
    checkStoredProcCompatibility = True
    Exit Function
ErrorHandlerSPDoesNoteExist:
    If err.Number = 2000 Then
        MsgBox ("In order to auto-generate data types, the Snowflake-Excel stored procedures must be created in the database you have logged into. " & _
        vbNewLine & "If you want to explicity define the data types please see click the 'Define Data Types' button in the ribbon.")
    Else
        MsgBox ("Error trying to retrieve Stored Procedure version number. " & err.Description)
    End If
    err.Raise giSuppressErrorMessage
End Function

Sub RibbonactivateHomeTab()
    On Error Resume Next
    gRibbon.ActivateTabMso "TabHome"
End Sub

Function CustomRange(sRange As String) As range
    Set CustomRange = range(sRange)
End Function

Function CustomRangeName(sRange As String)
    CustomRangeName = sRange
End Function

Sub SaveAllNamedRangesToAddIn()
    Dim nm As name
    'We want to save the named ranges to the addin so next time they open a new workbook, the connectino values will be there
    On Error Resume Next
    For Each nm In Names
        If CustomRange(nm.name).Count = 1 And InStr(nm.value, gsSnowflakeConfigWorksheetName) > 0 Then
            range(ThisWorkbook.name & "!" & nm.name) = CustomRange(nm.name)
        End If
    Next nm
    If Not PersistPassword Then
        range(ThisWorkbook.name & "!" & sgRangePassword) = ""
    End If
    On Error GoTo 0
    ThisWorkbook.save
End Sub

Function PersistPassword()
    PersistPassword = Not ThisWorkbook.IsAddin
End Function
'this is used to remove all the old connections
Sub RemoveConnections()
    Dim i As Long
    For i = ActiveWorkbook.Connections.Count To 1 Step -1
        ActiveWorkbook.Connections.Item(i).Delete
    Next 'i
End Sub

Sub OpenHelp(sSection As String)
    MsgBox ("Help is on it's way! But it might take a while:(")
    helpUrl = "https://www.snowflake.com/blog/"
    'ActiveWorkbook.FollowHyperlink Address:=helpUrl, NewWindow:=True
End Sub

Sub removeBadNamedRanges()
    'Remove all bad name ranges
    For Each n In ActiveWorkbook.Names
        If InStr(n.value, "#REF!") > 0 Then
            n.Delete
        End If
    Next n
End Sub
Sub CopySnowflakeConfgWS()
    Dim iWSVersionNumber As Integer
    If ActiveWorkbook Is Nothing Or Not ThisWorkbook.IsAddin Then
        Exit Sub
    End If
    'Check worksheet version number
    On Error GoTo HandleworksheetNotExist
    iWSVersionNumber = Utils.CustomRange(sgRangeWorksheetVersionNumber)
    On Error GoTo 0
    If iWSVersionNumber >= range(ThisWorkbook.name & "!" & sgRangeWorksheetVersionNumber) Then
        Exit Sub
    End If

    Dim tempNames As New Collection
    ' Save to apply them to the new names later
    For Each n In ActiveWorkbook.Names
        If InStr(n.value, "#REF!") > 0 Then
            n.Delete
        Else
            If InStr(n.value, gsSnowflakeConfigWorksheetName) > 0 And n.name <> sgRangeWorksheetVersionNumber Then
                tempNames.Add Utils.CustomRange(n.name).Value2, n.name
                n.Delete
            End If
        End If
    Next n
    'Delete workbook but turn off any warnings first
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(gsSnowflakeConfigWorksheetName).Delete
    Application.DisplayAlerts = True

HandleworksheetNotExist:
    On Error GoTo 0
    ' When a workbook is deleted the named ranges are not but reference in bad = #REF, so delete these
    For Each n In ActiveWorkbook.Names
        If InStr(n.value, "#REF!") > 0 Or InStr(n.value, gsSnowflakeConfigWorksheetName) > 0 Then
            n.Delete
        End If
    Next n
    ThisWorkbook.Sheets(gsSnowflakeConfigWorksheetName).Copy _
            After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    ActiveWorkbook.Sheets(gsSnowflakeConfigWorksheetName).Visible = False
    For Each n In ActiveWorkbook.Names
        ' Need to do this because the ranges with more than one cell is not moving over properly
        If InStr(n.value, ThisWorkbook.name) > 0 Then
            n.value = Replace(n.value, "[" & ThisWorkbook.name & "]", "")
        End If
    Next n
    If tempNames.Count > 0 Then
        On Error Resume Next
        For Each name In ActiveWorkbook.Names
            Utils.CustomRange(name.name) = tempNames.Item(name.name)
        Next
    End If

End Sub

Sub SetDateInputFormat()
    Utils.execSQLFireAndForget ("alter session set DATE_INPUT_FORMAT = '" & Utils.CustomRange(sgRangeDateInputFormat) & _
     "', TIMESTAMP_INPUT_FORMAT = '" & Utils.CustomRange(sgRangeTimestampInputFormat) & _
     "', TIME_INPUT_FORMAT = '" & Utils.CustomRange(sgRangeTimeInputFormat) & "'")
End Sub


Function getOrCreateRange(wsWorkbookParams As Worksheet, rangeName As String, colNumber As Integer)
    On Error GoTo CreateRange
    Set getOrCreateRange = Utils.CustomRange(rangeName)
    Exit Function

CreateRange:
    err.Clear
    On Error GoTo ErrorHandlerGeneral
    Dim bFoundEmpty As Boolean
    Dim checkCell As range

    With wsWorkbookParams
        i = 0
        ' Loops until it finds the next epmty cell in a column
        While Not bFoundEmpty
            i = i + 1
            Set checkCell = .Cells(i, 1)
            If checkCell = "" Then
                On Error Resume Next
                nm = ""
                nm = getRangeNameIgnoreError(checkCell)
                If nm = "" Then bFoundEmpty = True
                On Error GoTo ErrorHandlerGeneral
            End If
        Wend
        ActiveWorkbook.Names.Add name:=rangeName, _
        RefersTo:=wsWorkbookParams.range(.Cells(i, colNumber), .Cells(i, colNumber))
    End With
    Set getOrCreateRange = Utils.CustomRange(rangeName)
    Exit Function
ErrorHandlerGeneral:
    MsgBox (err.Description)
End Function

Function getRangeNameIgnoreError(range As range)
    On Error Resume Next
    Set getRangeNameIgnoreError = range.name
    On Error GoTo 0
End Function

Function worksheetBelongsToAddin()
    'Check to make sure data will not be loaded into one of the parameter of log worksheet that the Addin depends on.
    If ActiveSheet.name = gsSnowflakeWorkbookParamWorksheetName _
        Or ActiveSheet.name = gsSnowflakeConfigWorksheetName Or ActiveSheet.name = sgRangeLogWorksheet Then
        MsgBox ("You are not allowed to load data into one of the worksheets that is created by the Addin. Please select a different one.")
        worksheetBelongsToAddin = True
    Else
        worksheetBelongsToAddin = False
    End If
End Function

Function doesWorksheetExist()
    If ActiveSheet Is Nothing Then
        MsgBox ("There is no Worksheet available. Please create one before proceeding.")
        doesWorksheetExist = False
        Exit Function
    End If
    doesWorksheetExist = True
End Function


