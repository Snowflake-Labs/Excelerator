VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigForm 
   Caption         =   "Configuration"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "ConfigForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'For cancel values
Dim temp_tbResultsWorksheet As String       'Data will be written here from resutls of query
Dim temp_tbUploadWorksheet As String        'Data will be uploaded from this worksheet
Dim temp_tbLogWorksheet As String           'The execution status is written here
Dim temp_tbWindowsTempDirectory As String   'This is where the .csv file is saved
Dim temp_tbDateFormat As String             'This is where the input date format
Dim temp_tbStage As String
Dim wsSnowflakeConfig As Worksheet

Sub SetUpConfigData()
    ConfigForm.Show
End Sub

Private Sub ChangeRoleAndWarehouse()
    Set StatusForm = Nothing
    'Call StatusForm.execMethod("FormCommon", "UpdateUserRoleAndWarehouse")
    SetRoleAndWarehouseForm.ShowMe (True) ' True Forces it to open
    tbUserRole = Utils.CustomRange(sgRangeRole)
    tbUserWarehouse = Utils.CustomRange(sgRangeWarehouse)
'    StatusForm.Hide
End Sub

Private Sub lblUserRole_Click()
    ChangeRoleAndWarehouse
End Sub

Private Sub lblUserWarehouse_Click()
    ChangeRoleAndWarehouse
End Sub

Private Sub UserForm_Initialize()
    'set variables to rollback if cancelled
    Call setCancelVariables
    sWorkbookVers = gWorkbookVers
    tbUserRole = Utils.CustomRange(sgRangeRole)
    tbUserWarehouse = Utils.CustomRange(sgRangeWarehouse)
End Sub

Private Sub setCancelVariables()
    temp_tbResultsWorksheet = CustomRange(sgRangeResultsWorksheet)
    temp_tbUploadWorksheet = CustomRange(sgRangeUploadWorksheet)
    temp_tbLogWorksheet = CustomRange(sgRangeLogWorksheet)
    temp_tbWindowsTempDirectory = CustomRange(sgRangeWindowsTempDirectory)
    temp_tbDateFormat = CustomRange(sgRangeDateInputFormat)
    temp_tbTimestampFormat = CustomRange(sgRangeTimestampInputFormat)
    temp_tbTimeFormat = CustomRange(sgRangeTimeInputFormat)
    temp_tbStage = CustomRange(sgRangeStage)
End Sub

Private Sub OKButton_Click()
    If Len(Trim(tbLogWorksheet)) = 0 Or Len(Trim(tbWindowsTempDirectory)) = 0 Then
        MsgBox "Log Worksheet and Windows Temp director are mandatory."
        Exit Sub
    End If
    'if the Date format changed then change it in snowflake
    If temp_tbDateFormat <> CustomRange(sgRangeDateInputFormat) Or temp_tbTimestampFormat <> CustomRange(sgRangeTimestampInputFormat) Or temp_tbTimeFormat <> CustomRange(sgRangeTimeInputFormat) Then
        Call Utils.SetDateInputFormat
    End If
    Call setCancelVariables
    Me.Hide
    Utils.SaveAllNamedRangesToAddIn
    Set ConfigForm = Nothing
End Sub

Private Sub CancelButton_Click()
    'Reset values to original ones
    tbResultsWorksheet.Text = temp_tbResultsWorksheet
    tbUploadWorksheet.Text = temp_tbUploadWorksheet
    tbLogWorksheet.Text = temp_tbLogWorksheet
    tbWindowsTempDirectory.Text = temp_tbWindowsTempDirectory
    tbDateFormat.Text = temp_tbDateFormat
    tbStage.Text = temp_tbStage
    Me.Hide
    Set ConfigForm = Nothing
End Sub

Private Sub iHelpLink_Click()
    OpenHelp ("ConfigForm")
End Sub

Private Sub DownloadDriver_Click()
    helpUrl = "https://sfc-repo.snowflakecomputing.com/odbc/index.html"
    ActiveWorkbook.FollowHyperlink Address:=helpUrl, NewWindow:=True
End Sub

Private Sub DropObjCache_Click()
    Call FormCommon.dropDBObjectsCache
End Sub
