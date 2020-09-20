VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetValueForm 
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4125
   OleObjectBlob   =   "GetValueForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetValueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim bOKClicked As Boolean

Private Sub UserForm_Activate()
    tbValue.SetFocus
    With tbValue
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub btOK_Click()
    Me.Hide
    bOKClicked = True
End Sub

Private Sub CancelButton_Click()
    tbValue = ""
    Me.Hide
    bOKClicked = False
End Sub

Public Function Getvalue()
    Getvalue = tbValue
End Function

Public Sub setMessage(message As String)
    lblMessage = message
End Sub

Public Sub setValue(value As String)
    tbValue = value
End Sub

Public Function okClicked()
    okClicked = bOKClicked
End Function
