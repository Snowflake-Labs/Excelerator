Private WithEvents app As Application

Private Sub Workbook_AddinInstall()
    Call Utils.CopySnowflakeConfgWS
End Sub

Private Sub Workbook_Open()
    Debug.Print "Snowflake thisWorkbook_Open()"
    Set app = Excel.Application
    'Only add ribbon buttons when this is not run as an Addin. These buttons are for Dev only 12
    If Not ThisWorkbook.IsAddin Then
        Call AddAllButtons
    End If

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "Snowflake thisWorkbook_BeforeClose()"
    Call RemoveAllButtons
End Sub

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

Sub AddAllButtons()
    Call Add_Button("Create Addin", "'Build.createAddin'")
    Call Add_Button("Export Project Code", "'Build.exportVbProject'")
    Call Add_Button("Import Project Code", "'Build.importVbProject'")
End Sub
Sub RemoveAllButtons()
    Call Remove_Button("Create Addin")
    Call Remove_Button("Export Project Code")
    Call Remove_Button("Import Project Code")
End Sub

Sub Remove_Button(name As String) ' this is for running manually to remove a button
    Dim control As CommandBarButton
    Dim cmdBar As CommandBar
    Set cmdBar = Application.CommandBars("Worksheet Menu Bar")
    On Error Resume Next
    cmdBar.Controls(name).Delete
    On Error GoTo 0
End Sub


'****************** Not used below here ***********************
Sub ExportParametersToFile()
    Dim nm As name
    Dim paramNamesCSV As String
    Dim paramValuesCSV As String

    paramNamesCSV = ""
    paramValuesCSV = ""
    On Error GoTo errorHandlingComplete
    For Each nm In Names
        ' This is for persisting the values to a file. Only persist some of the values
        If persistParameter(nm.name) Then
            paramNamesCSV = paramNamesCSV & "," & nm.name
            paramValuesCSV = paramValuesCSV & "," & CustomRange(nm.name)
        End If
    Next nm
    ' Remove the leading comma
    paramNamesCSV = Replace(paramNamesCSV, ",", "", 1, 1)
    paramValuesCSV = Replace(paramValuesCSV, ",", "", 1, 1)
    exportFileName = getFileName(sgParameterFileName, sgParameterFileDirectory)
    Open exportFileName For Output As 1
    Print #1, paramNamesCSV
    Print #1, paramValuesCSV
    Close #1
errorHandlingComplete:
    On Error GoTo 0
End Sub

Sub ImportParametersFromFile()
    Dim DataLine As String
    Dim paramNames() As String
    Dim paramValues() As String
    Dim importFileName As String

    importFileName = getFileName(sgParameterFileName, sgParameterFileDirectory)

    Open importFileName For Input As 1
    Line Input #1, DataLine ' read in data #1 line at a time
    paramNames = Split(DataLine, ",")
    Line Input #1, DataLine
    paramValues = Split(DataLine, ",")
    For i = LBound(paramNames) To UBound(paramNames)
        CustomRange(paramNames(i)) = paramValues(i)
    Next
    Close #1
End Sub

Function getFileName(fileName As String, windowsDirectory As String)
    Dim windowsTempDirectory As String
    Dim snowflakeDirectory As String

    #If Mac Then
    fullFileName.save = CreateFolderinMacOffice2016(NameFolder:="snowflake_put") & "/" & fileName
    fullFileName.put = fullFileName.save
    #Else
    windowsTempDirectory = windowsDirectory
    snowflakeDirectory = windowsDirectory & "\Snowflake"
    dirExists = vba.FileSystem.Dir(windowsTempDirectory, vbDirectory)
    If dirExists = vba.Constants.vbNullString Then
        MkDir windowsTempDirectory
    End If
    dirExists2 = vba.FileSystem.Dir(snowflakeDirectory, vbDirectory)
    If dirExists2 = vba.Constants.vbNullString Then
        MkDir snowflakeDirectory
    End If
    getFileName = snowflakeDirectory & "\" & fileName
    #End If
    Debug.Print "Temp directory for parameters = " & getFileName
End Function

Function persistParameter(param As String)
    If (param = sgRangeAuthType) Or (param = sgRangeServer) Or (param = sgRangeUserID) Or (param = sgRangeRole) Or (param = sgRangeDefaultDatabase) Or (param = sgRangeDefaultSchema) Or (param = sgRangeStage) Or (param = sgRangeWarehouse) Then
        persistParameter = True
    Else
        persistParameter = False
    End If
End Function