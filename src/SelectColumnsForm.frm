VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectColumnsForm 
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   OleObjectBlob   =   "SelectColumnsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectColumnsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim table As String
Dim database As String
Dim schema As String
Dim selectedColumnsCSV As String

Private Sub btCancel_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    selectedColumnsCSV = ""
End Sub

Private Sub btSelect_Click()
    selectedColumnsCSV = ""
    For i = 0 To lbColumns.ListCount - 1
        If lbColumns.Selected(i) = True Then
            selectedColumnsCSV = selectedColumnsCSV & ", """ & lbColumns.list(i) & """"
        End If
    Next i
    selectedColumnsCSV = Replace(selectedColumnsCSV, ",", "", 1, 1)
    Me.Hide
End Sub

Public Sub initialize(databasePassed As String, schemaPassed As String, tablePassed As String)
    Dim sql As String
    Dim arrColumns As Variant

    If table = tablePassed And database = databasePassed And schema = schemaPassed Then
        Exit Sub
    End If
    'set module level vars
    table = tablePassed
    database = databasePassed
    schema = schemaPassed
    lbColumns.Clear

    arrColumns = FormCommon.getColumnArray(database, schema, table)
    On Error Resume Next ' can't figure how to trap fo an empty array/variant
    lbColumns.ColumnCount = 2
    lbColumns.ColumnWidths = "150,50"
    'Loop through array and add to List Box
    For i = 0 To UBound(arrColumns, 2)
        lbColumns.AddItem
        lbColumns.list(i, 0) = arrColumns(0, i)
        lbColumns.list(i, 1) = arrColumns(1, i)
    Next i
End Sub

Public Function getSelectedColunms()
    getSelectedColunms = selectedColumnsCSV
End Function

