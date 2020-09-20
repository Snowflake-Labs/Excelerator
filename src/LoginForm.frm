VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Connection"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'For cancel values
Dim temp_tbUserID As String
Dim temp_tbServer As String
Dim temp_tbRole As String
Dim temp_tbDatabase As String
Dim temp_tbSchema As String
Dim temp_tbWarehouse As String
Dim temp_tbStage As String
Dim temp_authType As String
Dim temp_tbPassword As String
Dim sExampleServerURL As String
' Was the OK or cancel button pressed
Public bLoginOK As Boolean

Private Sub rbSSO_Click()
    gsAuthenticationType = "SSO"
    CustomRange(sgRangeAuthType) = "SSO"
    tbPassword.Enabled = False
End Sub

Private Sub rbUserPass_Click()
    gsAuthenticationType = "UserPass"
    CustomRange(sgRangeAuthType) = "User & Pass"
    tbPassword.Enabled = True
End Sub

Private Sub setCancelVariables()
    temp_tbUserID = CustomRange(sgRangeUserID)
    temp_tbServer = CustomRange(sgRangeServer)
    temp_tbRole = CustomRange(sgRangeRole)
    temp_tbDatabase = CustomRange(sgRangeDefaultDatabase)
    temp_tbSchema = CustomRange(sgRangeDefaultSchema)
    temp_tbWarehouse = CustomRange(sgRangeWarehouse)
    temp_tbStage = CustomRange(sgRangeStage)
    temp_authType = CustomRange(sgRangeAuthType)
    temp_tbPassword = CustomRange(sgRangePassword)
End Sub

Private Sub tbServer_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If InStr(1, tbServer, "https://") = 1 Then
        'remove https:// if exists
        tbServer = Replace(tbServer, "https://", "")
        CustomRange(sgRangeServer) = tbServer
    End If
End Sub
Private Sub tbServer_Enter()
    'If the server url field has the example, then remove it upon entry
    If tbServer = sExampleServerURL Then
        tbServer = ""
        CustomRange(sgRangeServer) = tbServer
    End If
End Sub

Private Sub UserForm_Initialize()
    'to rollback if cancelled
    Call setCancelVariables
    'remove all empty ranges. Doing it here because there is no other great place
    Call Utils.removeBadNamedRanges
    If temp_authType = "SSO" Then
        rbSSO = True
        tbPassword.Enabled = False
    Else
        rbUserPass = True
        tbPassword.SetFocus
    End If
    bLoginOK = False
    If Not PersistPassword Then
        LoginForm.tbPassword.ControlSource = ""
        ' Range(ThisWorkbook.Name & "!" & sgRangePassword) = ""
    End If
    sExampleServerURL = "ex. myAccountName.snowflakecomputing.com"
    If tbServer = "" Then
        tbServer = sExampleServerURL
    End If
End Sub

Private Sub OKButton_Click()
    If rbUserPass And Len(Trim(tbPassword)) = 0 Then
        MsgBox "Password is mandatory."
        Exit Sub
    End If
    If Len(Trim(tbServer)) = 0 Or Len(Trim(tbUserID)) = 0 Then
        MsgBox "User ID and Server are mandatory."
        Exit Sub
    End If
    bLoginOK = True
    If tbRole.Text <> temp_tbRole Then
        Call FormCommon.dropDBObjectsCache
    End If
    Call setCancelVariables
    Me.Hide
    Utils.SaveAllNamedRangesToAddIn

End Sub

Private Sub CancelButton_Click()
    'Reset values to original ones
    tbUserID.Text = temp_tbUserID
    tbServer.Text = temp_tbServer
    tbRole.Text = temp_tbRole
    tbDatabase.Text = temp_tbDatabase
    tbSchema.Text = temp_tbSchema
    tbWarehouse.Text = temp_tbWarehouse
    tbStage.Text = temp_tbStage
    tbPassword.Text = temp_tbPassword
    CustomRange(sgRangeAuthType) = temp_authType
    If temp_authType = "SSO" Then
        rbSSO = True
        tbPassword.Enabled = False
    Else
        rbUserPass = True
        tbPassword.SetFocus
    End If
    bLoginOK = False
    Me.Hide
End Sub
Private Sub iHelpLink_Click()
    OpenHelp ("ConfigForm")
End Sub
