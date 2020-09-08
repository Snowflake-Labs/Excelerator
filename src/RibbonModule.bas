Attribute VB_Name = "RibbonModule"
Option Explicit

Public MyTag As String

Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Sub EnableControl(control As IRibbonControl, ByRef returnedVal)
    'Returns trueif not in read only mode
    returnedVal = True
    If isAddinReadOnly Then
        If control.Tag = "ReadWrite" Then
            returnedVal = False
        End If
    End If
End Sub

Public Function isAddinReadOnly()
    On Error GoTo ErrorHandlerNotDefined
    isAddinReadOnly = range(ThisWorkbook.name & "!" & sgRangeReadOnly)
    Exit Function
ErrorHandlerNotDefined:
    isAddinReadOnly "False"
End Function

Public Sub setAddinReadOnly()
    Utils.CustomRange(sgRangeReadOnly) = "True"
End Sub
Public Sub setAddinReadWrite()
    Utils.CustomRange(sgRangeReadOnly) = "False"
End Sub

Sub ResetRibbonButtonsVisibility()
    'not used but would be needed if ReadOnly status changes and you don't want to close and reopen workbook.
    gRibbon.Invalidate
End Sub
