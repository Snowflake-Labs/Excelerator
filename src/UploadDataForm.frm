VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UploadDataForm 
   Caption         =   "Upload Data"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5850
   OleObjectBlob   =   "UploadDataForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UploadDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim sKeyExample As String
Dim bValidMergKeys As Boolean
Dim bShowMergKeyMessage As Boolean
Dim uploadWorksheet As Worksheet
Dim wsWorkbookParams As Worksheet
Dim uploadMergeKeysByLettersRange As range
Dim uploadTypeRange As range
Dim schemaComboboxInitialized As Boolean
Dim bInitializing As Boolean



Public Sub UserForm_Initialize()
    'Set uploadWorksheet = ActiveWorkbook.Sheets(CustomRange(sgRangeUploadWorksheet).value)
    '  Dim uploadMergeKeysRange_Name As String
    Dim uploadMergeKeysByLettersRange_Name As String
    Dim uploadMergeKeysRange_Name As String
    StatusForm.Update_Status ("Preparing Upload Form...")
    Set wsWorkbookParams = Utils.getWorksheet(gsSnowflakeWorkbookParamWorksheetName)
    If CustomRange(sgRangeUploadWorksheet) <> "" Then
        Utils.getWorksheet(CustomRange(sgRangeUploadWorksheet)).Activate
    End If
    Set uploadWorksheet = ActiveSheet

    If sgShowVerifyMsg = "" Then
        sgShowVerifyMsg = "True"
    End If
    sKeyExample = "ex. A,B..."
    bValidMergKeys = True
    bShowMergKeyMessage = True
    '************* Setting Comboboxes *************
    bInitializing = True
    Call FormCommon.initializeDBObjectsComboBoxes(cbDatabases, cbSchemas, cbTables)
    bInitializing = False

    ' Initialize ranges
    Set rgUploadMergeKeysRange = FormCommon.initializeRange("MergeKeysNumbers")
    Set uploadMergeKeysByLettersRange = FormCommon.initializeRange("MergeKeysLetters")
    Set uploadTypeRange = FormCommon.initializeRange("UploadType")
    ' set merge field control source
    tbMergeKeys.ControlSource = uploadMergeKeysByLettersRange.name
    Call StatusForm.Hide
End Sub
Private Sub UserForm_Activate()
    ' tbUploadWorksheet = ActiveSheet.name  ' This is to change the update sheet to the active sheet. Maybe implement later

    Select Case uploadTypeRange
        Case "Append"
            rbAppend = True
            rbAppend_Click
        Case "Truncate"
            rbTruncate = True
            rbTruncate_Click
        Case Else 'Merge
            rbMerge = True
            rbMerge_Click
    End Select
    setUploadTypeText
    setMergeKeysBorderColor
End Sub

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

Private Sub rbMerge_Click()
    tbMergeKeys.Enabled = True
    If tbMergeKeys = "" Or tbMergeKeys = sKeyExample Then
        tbMergeKeys = sKeyExample
        tbMergeKeys.ForeColor = &H80000011
    End If
    setUploadTypeText
    uploadTypeRange = "Merge"
    setMergeKeysBorderColor
End Sub

Private Sub rbAppend_Click()
    tbMergeKeys.Enabled = False
    If tbMergeKeys = sKeyExample Then
        tbMergeKeys = ""
    End If
    setUploadTypeText
    uploadTypeRange = "Append"
    setMergeKeysBorderColor
End Sub
Private Sub rbTruncate_Click()
    rbAppend_Click
    uploadTypeRange = "Truncate"
End Sub

Private Sub tbMergeKeys_Enter()
    If tbMergeKeys = sKeyExample Then
        tbMergeKeys = ""
        tbMergeKeys.ForeColor = &H80000001
    End If
    bShowMergKeyMessage = True
End Sub
Sub setMergeKeysBorderColor()
    If rbMerge And Not (cbRecreateTable Or cbCreateNewTable) And (tbMergeKeys = "" Or tbMergeKeys = sKeyExample Or Not bValidMergKeys) Then
        tbMergeKeys.BorderColor = &HFF&  'Red
    Else
        tbMergeKeys.BorderColor = 11119017 'Black
    End If
End Sub

Sub tbMergeKeys_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    tbMergeKeys = LTrim(tbMergeKeys)
    tbMergeKeys = RTrim(tbMergeKeys)
    rbMerge_Click

    If bShowMergKeyMessage Then
        updateMergeKeysByNumber
        setMergeKeysBorderColor
    End If
    bShowMergKeyMessage = True ' need to do this because if the user clicks Update then this fires after and there is 2 messages
End Sub
Function updateMergeKeysByNumber()
    'This proc converts A,B,C to 1,2,3
    Dim colArr() As String
    Dim sMergKeysByNumber As String

    If tbMergeKeys = "" Or tbMergeKeys = sKeyExample Then
        MsgBox ("The 'Table Key Columns' is mandatory.")
        GoTo ValidationFailure
    End If
    On Error GoTo ErrorHandlerBadMergKey
    sMergKeysByNumber = ""
    colArr = Split(tbMergeKeys, ",")
    For i = LBound(colArr) To UBound(colArr)
        sMergKeysByNumber = sMergKeysByNumber & "," & range(colArr(i) & 1).Column

        If WorksheetFunction.CountA(uploadWorksheet.range(colArr(i) & 1, colArr(i) & 2)) = 0 Then
            MsgBox "Column " & colArr(i) & " is empty. A key column must contain values to identify uniqueness for each row."
            GoTo ValidationFailure
        End If
    Next i
    rgUploadMergeKeysRange = Replace(sMergKeysByNumber, ",", "", 1, 1)
    '  CustomRange(uploadMergeKeysRange_Name) = Replace(sMergKeysByNumber, ",", "", 1, 1)
    updateMergeKeysByNumber = True
    bValidMergKeys = True
    Exit Function
ErrorHandlerBadMergKey:
    MsgBox ("Invalid 'Table Key Columns'. It must be in the format of 'A,B,C', where the letters are the worksheet columns that make each row unique.")
    GoTo ValidationFailure
ValidationFailure:
    updateMergeKeysByNumber = False
    rgUploadMergeKeysRange = ""
    bValidMergKeys = False
End Function

Sub setUploadTypeText()
    Select Case True
        Case cbCreateNewTable
            lblUploadTypeText = "A new table will be created with all data will be INSERTED."
        Case rbTruncate Or cbRecreateTable
            lblUploadTypeText = "All data in the table will be DELETED." & vbNewLine & "All data in the worksheet will be INSERTED."
        Case rbAppend
            lblUploadTypeText = "All data will be INSERTED." & vbNewLine & "To UPDATE rows, please click the 'Update' radio button above."
        Case Else
            lblUploadTypeText = vbNewLine & "Existing rows will be UPDATED and new rows will be INSERTED."
    End Select
End Sub

Public Sub ShowForm()
    If Utils.login Then
        Me.Show
    End If

End Sub
Function MsgBoxVerifyUpload()
    If sgShowVerifyMsg = "True" Then
        MsgBoxVerifyUpload = MsgBox("You are about to upload data to table " & cbTables.value & vbNewLine & _
        "Do you want to continue?" & vbNewLine & vbNewLine & "(This warning will not be displayed again for this session.)", vbOKCancel + vbExclamation)
        If MsgBoxVerifyUpload Then
            sgShowVerifyMsg = "False"
        End If
    Else
        MsgBoxVerifyUpload = vbOK
    End If
End Function

Public Sub btUpload_Click()
    Dim bContinue As Boolean
    If Not cbCreateNewTable And (Right(cbTables.value, 1) = "." Or cbTables.value = "") Then
        MsgBox ("Please enter a valid table name.")
    Else
        bContinue = True
        If rbMerge And Not (cbRecreateTable Or cbCreateNewTable) Then
            If tbMergeKeys = "" Or tbMergeKeys = sKeyExample Then
                MsgBox ("The 'Table Key Columns' is mandatory.")
                bContinue = False
                bShowMergKeyMessage = False
            End If
            If bContinue Then
                If Not updateMergeKeysByNumber Then
                    bContinue = False
                    bShowMergKeyMessage = False
                End If
            End If
        End If
        If bContinue Then
            uploadData ("")
        End If
    End If
End Sub

Public Sub uploadData(sUploadType As String)
    Dim uploadTable As String
    If cbCreateNewTable Then
        cbTables.value = tbNewTableName
    End If
    Call FormCommon.saveDBObjectsValues(cbDatabases, cbSchemas, cbTables)
    If cbTables.value = "" Then
        MsgBox "Please specify which table to upload to."
        If Me.Visible = False Then
            Me.Show
        End If
        Exit Sub
    End If

    If MsgBoxVerifyUpload = vbOK Then
        'check if the table has been updated since the last time it was downloaded
        If checkIfTableHasBeenAltered = True Then
            If MsgBox("The table has been updated since the last data download time." & _
            vbNewLine & "Continue uplading?", vbOKCancel, "Update Conflict") = vbCancel Then
                UploadDataForm.Hide
                Set UploadDataForm = Nothing
                Exit Sub
            End If
        End If
        uploadTable = """" & cbDatabases.value & """.""" & cbSchemas.value & """.""" & cbTables.value & """"
        If cbAutoGenDataTypes Then ' do everything in the stored proc
            Select Case True
                Case cbRecreateTable Or cbCreateNewTable
                    sUploadType = "RecreateTable"
                Case rbTruncate
                    sUploadType = "Truncate"
                Case rbAppend
                    sUploadType = "Append"
                Case rbMerge
                    sUploadType = "Merge"
            End Select
            Else ' do everything local
                Select Case True
                    Case cbCreateNewTable
                        sUploadType = "CreateLocal"
                    Case cbRecreateTable
                        sUploadType = "RecreateLocal"
                    Case rbTruncate
                        sUploadType = "TruncateLocal"
                    Case rbAppend
                        sUploadType = "AppendLocal"
                    Case rbMerge
                        sUploadType = "MergeLocal"
                End Select
        End If
        Utils.SaveAllNamedRangesToAddIn
        Set StatusForm = Nothing
        Me.Hide
        Call StatusForm.execMethod("Load", "UploadData", sUploadType, uploadTable, rgUploadMergeKeysRange)
        'set the download datetime
        Query.setDownloadDateTime
        Set UploadDataForm = Nothing
    End If
End Sub

Function checkIfTableHasBeenAltered()
    Dim lastAlteredSQL As String
    Dim downloadedDatTime As String

    On Error GoTo ErrorHandlerGeneral
    'Get the date the data was downloaded to this worksheet
    downloadedDatTime = Query.getDownloadDateTime
    ' If the date does not exist then return false so upload will continue
    If downloadedDatTime = "" Then
        checkIfTableHasBeenAltered = "False"
        Exit Function
    End If
    'Check if the last_altered data of the table is later than the download date
    lastAlteredSQL = "Select IFF( last_altered > '" & Format(Query.getDownloadDateTime, "YYYY-MM-DD HH:mm:SS") & "' , 'TRUE' , 'FALSE' ) From """ & _
    cbDatabases.value & """.information_schema.tables where table_schema = '" & _
    cbSchemas.value & "' and table_name = '" & cbTables.value & "'"
    checkIfTableHasBeenAltered = Utils.execSQLSingleValueOnly(lastAlteredSQL)
    Exit Function
ErrorHandlerGeneral:
    Call Utils.handleError("Error ckeching if table has been altered", err)
    checkIfTableHasBeenAltered = "False"
End Function

Private Sub CancelButton_Click()
    Me.Hide
    Call FormCommon.saveDBObjectsValues(cbDatabases, cbSchemas, cbTables)
    Set UploadDataForm = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    CancelButton_Click
End Sub

Private Sub cbRecreateTable_click()
    If cbRecreateTable Then
        cbCreateNewTable = False
        handleCreateOrRecreateCheck
    End If
End Sub

Private Sub cbCreateNewTable_Click()
    If cbCreateNewTable Then
        Set GetValueForm = Nothing
        GetValueForm.setMessage ("Enter Table Name:")
        GetValueForm.setValue (tbNewTableName)
        GetValueForm.Show
        If GetValueForm.okClicked Then
            tbNewTableName = UCase(GetValueForm.Getvalue)
            If tbNewTableName = "" Then
                cbCreateNewTable = False
            End If
            cbRecreateTable = False
            handleCreateOrRecreateCheck
        Else
            cbCreateNewTable = False
        End If
    Else
        handleCreateOrRecreateCheck
    End If
End Sub

Sub handleCreateOrRecreateCheck()
    If cbRecreateTable Or cbCreateNewTable Then
        tbMergeKeys.Enabled = False
        rbAppend.Enabled = False
        rbMerge.Enabled = False
        rbTruncate.Enabled = False
        lblMergeKeys.Enabled = False

    Else
        tbMergeKeys.Enabled = True
        rbAppend.Enabled = True
        rbMerge.Enabled = True
        rbTruncate.Enabled = True
        lblMergeKeys.Enabled = True
    End If
    
    If cbCreateNewTable Then
        tbNewTableName.Enabled = True
    Else
        tbNewTableName.Enabled = False
    End If
    
    setUploadTypeText
    setMergeKeysBorderColor
End Sub
Private Sub iHelpLink_Click()
    OpenHelp ("ConfigForm")
End Sub

Private Sub iMergeKeyHelp_Click()
    MsgBox ("The 'Table Key Columns' is a comma sepeated list of worksheet columns, represented by the column letter, that is used to identify each row uniquely. " & _
    "This is needed when updating a table. For example, if a table's unique identifier is the first 2 worksheet columns, the value would be 'A,B'.")
End Sub

Function getDownloadDateTime()
    Dim lockRangeTableDate As range
    Set lockRangeTableDate = FormCommon.initializeRange("LockTableDate")
    getDownloadDateTime = lockRangeTableDate
End Function

