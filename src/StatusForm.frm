VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatusForm 
   Caption         =   "Upload Status"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4035
   OleObjectBlob   =   "StatusForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Dim gsLoadType As String
Dim bInProcess As Boolean
Dim gsModule As String
Dim gsMethod As String
Dim gaParams As Variant

Private Sub btClose_Click()
    Me.Hide
    btClose.Visible = False
End Sub

Private Sub UserForm_Initialize()
    Update_Status ("Preparing...")
End Sub

Public Sub Update_Status(sStatus As String)
    tbStatus = sStatus
    DoEvents
End Sub

Public Sub execMethod(sModule As String, sMethod As String, ParamArray aParams() As Variant)
    bInProcess = False
    gsModule = sModule
    gsMethod = sMethod
    gaParams = aParams
    Me.Show
End Sub


Private Sub UserForm_Activate()
    If Not bInProcess Then
        btClose.Visible = False
        bInProcess = True
        'Get the number of prameters
        iCount = UBound(gaParams) - LBound(gaParams) + 1
        'Based on the number of parameters, select the proper Run statement
        Select Case iCount
            Case 0
                Application.Run gsModule & "." & gsMethod
            Case 1
                Application.Run gsModule & "." & gsMethod, (gaParams(0))
            Case 2
                Application.Run gsModule & "." & gsMethod, (gaParams(0)), (gaParams(1))
            Case 3
                Application.Run gsModule & "." & gsMethod, (gaParams(0)), gaParams(1), gaParams(2)
        End Select
        Me.Repaint ' This might be need because the window sometime appears behind Excel.
        btClose.Visible = True
    End If
End Sub


