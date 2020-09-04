Attribute VB_Name = "Setup"


'***These procs should not be needed anymore. Mybae useful for the Mac
Sub Add_Button(caption As String, action As String)

    Dim control As CommandBarButton
    Dim cmdBar As CommandBar
    Set cmdBar = Application.CommandBars("Worksheet Menu Bar")
    On Error Resume Next
    cmdBar.Controls(caption).Delete
    On Error GoTo 0

    Set control = cmdBar.Controls.Add
    With control
        .caption = caption
        .Style = msoButtonIconAndCaptionBelow 'msoButtonCaption
        .OnAction = action
    End With
End Sub

Sub Remove_Button() ' this is for running manually to remove a button
    Dim control As CommandBarButton
    Dim cmdBar As CommandBar
    Set cmdBar = Application.CommandBars("Worksheet Menu Bar")
    'On Error Resume Next
    cmdBar.Controls("Execute SQL").Delete
    On Error GoTo 0
End Sub
