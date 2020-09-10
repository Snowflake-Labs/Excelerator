Attribute VB_Name = "Build"
'''
'The Import and xport part of this code was taken from vbaDeveloper
' https://github.com/hilkoc/vbaDeveloper
'''
Option Explicit

Private Const IMPORT_DELAY As String = "00:00:03"
Private Const DELETE_DELAY As String = "00:00:05"

'We need to make these variables public such that they can be given as arguments to application.ontime()
Public componentsToImport As Dictionary 'Key = componentName, Value = componentFilePath
Public sheetsToImport As Dictionary 'Key = componentName, Value = File object
Public vbaProjectToImport As VBProject

Private Const NAMED_RANGES_FILE_NAME As String = "NamedRanges.csv"

Private Enum columns
    name = 0
    RefersTo
    Comments
End Enum

'***************** Creates Addin *********************
Sub createAddin()
    ' In order to make changes to this workbook so we can save the .xlam and then revert those changes, we have to do:
    '1 - Save this workbook with a TEMP name, This will make the current workbook have that new name
    '2 - Make the necessary changes and save as .xlam
    '3 - Open the original file
    '4 - Call a method in the original file to delete the TEMP file from he file system with a delay. That will give the temp file some time to close itself
    '5 - Close this file
    Dim wb As Workbook
    Dim fileName As String
    Dim origFileName As String, origFullFileName As String, tempFullFileName As String
    Dim sworksheetVersionNumber As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wb = Workbooks(ActiveWorkbook.name)
    ' Save in case. This file will be closed and reopened
    wb.save
    origFileName = wb.name
    origFullFileName = wb.FullName
    tempFullFileName = ThisWorkbook.Path & "\TEMP_" & origFileName
    wb.SaveAs fileName:=tempFullFileName, CreateBackup:=False

    'Delete all worksheets except the config one
    CleanupWorksheets
    'capture worksheet version number so it can be applied back after the cleanup
    sworksheetVersionNumber = Utils.CustomRange(sgRangeWorksheetVersionNumber)
    ' Remove ranges that are invalid and set others to empty
    CleanupRanges
    're-aply the worksheet version number
    Utils.CustomRange(sgRangeWorksheetVersionNumber) = sworksheetVersionNumber
    
    ' Application.DisplayAlerts = False
    Set wb = Workbooks(ActiveWorkbook.name)
    Call RibbonModule.setAddinReadWrite
    wb.SaveAs fileName:=ThisWorkbook.Path & "\" & "SnowflakeExcelAddin.xlam", FileFormat:=xlOpenXMLAddIn, CreateBackup:=False
    
    Call RibbonModule.setAddinReadOnly
    wb.SaveAs fileName:=ThisWorkbook.Path & "\" & "SnowflakeExcelAddinReadOnly.xlam", FileFormat:=xlOpenXMLAddIn, CreateBackup:=False
    'Open the original app
    Workbooks.Open origFullFileName
    'Send a request to the orig file to delete this file
    Application.Run "'" & origFullFileName & "'!Build.deleteFileWithDelay", tempFullFileName
    'close this file
    wb.Close
    
End Sub


Sub deleteFileWithDelay(fullFileName As String)
    'this will permantly delete a file, be careful
    Application.OnTime Now() + TimeValue(DELETE_DELAY), "'Build.deleteFile """ & fullFileName & """'"
End Sub
Sub deleteFile(fullFileName As String)
    Kill fullFileName
End Sub

Sub CleanupRanges()
    Dim n As name
    
    For Each n In ActiveWorkbook.Names
        If InStr(n.value, "#REF!") > 0 Then
            n.Delete
        Else
            n.RefersToRange = ""
        End If
    Next
    'since we emptied all ranges above, we need to set the defaults
    setRangeDefaultValues
End Sub

Sub setRangeDefaultValues()
    ' all ranges set to the default except the worksheet version number. That should be set in the calling sub
    Utils.CustomRange(sgRangeSnowflakeDriver) = "{SnowflakeDSIIDriver}"
    Utils.CustomRange(sgRangeAuthType) = "User & Pass"
    Utils.CustomRange(sgRangeLogWorksheet) = "Log"
    Utils.CustomRange(sgRangeWindowsTempDirectory) = "C:\temp"
    Utils.CustomRange(sgRangeDateInputFormat) = "Auto"
    Utils.CustomRange(sgRangeTimestampInputFormat) = "Auto"
    Utils.CustomRange(sgRangeTimeInputFormat) = "Auto"
    Utils.CustomRange(sgRangeReadOnly) = "False" ' This should be set when building the addin
End Sub

Sub CleanupWorksheets()
    Dim ws As Worksheet
    'Delete all worksheets except for the config one
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(gsSnowflakeConfigWorksheetName).Visible = True
    For Each ws In Worksheets
        If ws.name <> gsSnowflakeConfigWorksheetName Then
            ws.Delete
        End If
    Next
End Sub

'***************** Import and Export code *********************

' Returns the directory where code is exported to or imported from.
' When createIfNotExists:=True, the directory will be created if it does not exist yet.
' This is desired when we get the directory for exporting.
' When createIfNotExists:=False and the directory does not exist, an empty String is returned.
' This is desired when we get the directory for importing.
'
' Directory names always end with a '\', unless an empty string is returned.
' Usually called with: fullWorkbookPath = wb.FullName or fullWorkbookPath = vbProject.fileName
' if the workbook is new and has never been saved,
' vbProject.fileName will throw an error while wb.FullName will return a name without slashes.
Public Function getSourceDir(fullWorkbookPath As String, createIfNotExists As Boolean) As String
    ' First check if the fullWorkbookPath contains a \.
    If Not InStr(fullWorkbookPath, "\") > 0 Then
        'In this case it is a new workbook, we skip it
        Exit Function
    End If

    Dim FSO As New Scripting.FileSystemObject
    Dim projDir As String
    projDir = FSO.GetParentFolderName(fullWorkbookPath) & "\"
    Dim srcDir As String
    srcDir = projDir & "src\"
    Dim exportDir As String
    exportDir = srcDir ' SEGAL commented this out because I want to export to src directory directly: & FSO.GetFileName(fullWorkbookPath) & "\"

    If createIfNotExists Then
        If Not FSO.FolderExists(srcDir) Then
            FSO.CreateFolder srcDir
            Debug.Print "Created Folder " & srcDir
        End If
        If Not FSO.FolderExists(exportDir) Then
            FSO.CreateFolder exportDir
            Debug.Print "Created Folder " & exportDir
        End If
    Else
        If Not FSO.FolderExists(exportDir) Then
            Debug.Print "Folder does not exist: " & exportDir
            exportDir = ""
        End If
    End If
    getSourceDir = exportDir
End Function


' Called from Menu
Public Sub exportVbaCode(vbaProject As VBProject)
    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = vbaProject.fileName
    On Error GoTo 0
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=True)

    Debug.Print "exporting to " & export_path
    'export all components
    Dim component As VBComponent
    For Each component In vbaProject.VBComponents
        'lblStatus.Caption = "Exporting " & proj_name & "::" & component.Name
        If hasCodeToExport(component) Then
            'Debug.Print "exporting type is " & component.Type
            Select Case component.Type
                Case vbext_ct_ClassModule
                    exportComponent export_path, component
                Case vbext_ct_StdModule
                    exportComponent export_path, component, ".bas"
                Case vbext_ct_MSForm
                    exportComponent export_path, component, ".frm"
                Case vbext_ct_Document
                    exportLines export_path, component
                Case Else
                    'Raise "Unkown component type"
            End Select
        End If
    Next component
End Sub


Private Function hasCodeToExport(component As VBComponent) As Boolean
    hasCodeToExport = True
    If component.codeModule.CountOfLines <= 2 Then
        Dim firstLine As String
        firstLine = Trim(component.codeModule.lines(1, 1))
        'Debug.Print firstLine
        hasCodeToExport = Not (firstLine = "" Or firstLine = "Option Explicit")
    End If
End Function


'To export everything else but sheets
Private Sub exportComponent(exportPath As String, component As VBComponent, Optional extension As String = ".cls")
    Debug.Print "exporting " & component.name & extension
    component.Export exportPath & "\" & component.name & extension
End Sub


'To export sheets
Private Sub exportLines(exportPath As String, component As VBComponent)
    Dim extension As String: extension = ".sheet.cls"
    Dim fileName As String
    fileName = exportPath & "\" & component.name & extension
    Debug.Print "exporting " & component.name & extension
    'component.Export exportPath & "\" & component.name & extension
    Dim FSO As New Scripting.FileSystemObject
    Dim outStream As TextStream
    Set outStream = FSO.CreateTextFile(fileName, True, False)
    outStream.Write (component.codeModule.lines(1, component.codeModule.CountOfLines))
    outStream.Close
End Sub


' Usually called after the given workbook is opened. The option includeClassFiles is False by default because
' they don't import correctly from VBA. They'll have to be imported manually instead.
Public Sub importVbaCode(vbaProject As VBProject, Optional includeClassFiles As Boolean = False)
    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = vbaProject.fileName
    On Error GoTo 0
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=False)
    If export_path = "" Then
        'The source directory does not exist, code has never been exported for this vbaProject.
        Debug.Print "No import directory for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    'initialize globals for Application.OnTime
    Set componentsToImport = New Dictionary
    Set sheetsToImport = New Dictionary
    Set vbaProjectToImport = vbaProject

    Dim FSO As New Scripting.FileSystemObject
    Dim projContents As Folder
    Set projContents = FSO.GetFolder(export_path)
    Dim file As Object
    For Each file In projContents.Files()
        'check if and how to import the file
        checkHowToImport file, includeClassFiles
    Next

    Dim componentName As String
    Dim vComponentName As Variant
    'Remove all the modules and class modules
    For Each vComponentName In componentsToImport.keys
        componentName = vComponentName
        removeComponent vbaProject, componentName
    Next
    'Then import them
    Debug.Print "Invoking 'Build.importComponents'with Application.Ontime with delay " & IMPORT_DELAY
    ' to prevent duplicate modules, like MyClass1 etc.
    Application.OnTime Now() + TimeValue(IMPORT_DELAY), "'Build.importComponents'"
    Debug.Print "almost finished importing code for " & vbaProject.name
End Sub


Private Sub checkHowToImport(file As Object, includeClassFiles As Boolean)
    Dim fileName As String
    fileName = file.name
    Dim componentName As String
    componentName = Left(fileName, InStr(fileName, ".") - 1)
    If componentName = "Build" Then
        '"don't remove or import ourself
        Exit Sub
    End If

    If Len(fileName) > 4 Then
        Dim lastPart As String
        lastPart = Right(fileName, 4)
        Select Case lastPart
            Case ".cls" ' 10 == Len(".sheet.cls")
                If Len(fileName) > 10 And Right(fileName, 10) = ".sheet.cls" Then
                    'import lines into sheet: importLines vbaProjectToImport, file
                    sheetsToImport.Add componentName, file
                Else
                    ' .cls files don't import correctly because of a bug in excel, therefore we can exclude them.
                    ' In that case they'll have to be imported manually.
                    If includeClassFiles Then
                        'importComponent vbaProject, file
                        componentsToImport.Add componentName, file.Path
                    End If
                End If
            Case ".bas", ".frm"
                'importComponent vbaProject, file
                componentsToImport.Add componentName, file.Path
            Case Else
                'do nothing
                Debug.Print "Skipping file " & fileName
        End Select
    End If
End Sub


' Only removes the vba component if it exists
Private Sub removeComponent(vbaProject As VBProject, componentName As String)
    If componentExists(vbaProject, componentName) Then
        Dim c As VBComponent
        Set c = vbaProject.VBComponents(componentName)
        Debug.Print "removing " & c.name
        vbaProject.VBComponents.Remove c
    End If
End Sub


Public Sub importComponents()
    If componentsToImport Is Nothing Then
        Debug.Print "Failed to import! Dictionary 'componentsToImport' was not initialized."
        Exit Sub
    End If
    Dim componentName As String
    Dim vComponentName As Variant
    For Each vComponentName In componentsToImport.keys
        componentName = vComponentName
        importComponent vbaProjectToImport, componentsToImport(componentName)
    Next

    'Import the sheets
    For Each vComponentName In sheetsToImport.keys
        componentName = vComponentName
        importLines vbaProjectToImport, sheetsToImport(componentName)
    Next

    Debug.Print "Finished importing code for " & vbaProjectToImport.name
    'We're done, clear globals explicitly to free memory.
    Set componentsToImport = Nothing
    Set vbaProjectToImport = Nothing
End Sub


' Assumes any component with same name has already been removed.
Private Sub importComponent(vbaProject As VBProject, filePath As String)
    Debug.Print "Importing component from  " & filePath
    'This next line is a bug! It imports all classes as modules!
    vbaProject.VBComponents.Import filePath
End Sub


Private Sub importLines(vbaProject As VBProject, file As Object)
    Dim componentName As String
    componentName = Left(file.name, InStr(file.name, ".") - 1)
    Dim c As VBComponent
    If Not componentExists(vbaProject, componentName) Then
        ' Create a sheet to import this code into. We cannot set the ws.codeName property which is read-only,
        ' instead we set its vbComponent.name which leads to the same result.
        Dim addedSheetCodeName As String
        addedSheetCodeName = addSheetToWorkbook(componentName, vbaProject.fileName)
        Set c = vbaProject.VBComponents(addedSheetCodeName)
        c.name = componentName
    End If
    Set c = vbaProject.VBComponents(componentName)
    Debug.Print "Importing lines from " & componentName & " into component " & c.name

    ' At this point compilation errors may cause a crash, so we ignore those.
    On Error Resume Next
    c.codeModule.DeleteLines 1, c.codeModule.CountOfLines
    c.codeModule.AddFromFile file.Path
    On Error GoTo 0
End Sub


Public Function componentExists(ByRef proj As VBProject, name As String) As Boolean
    On Error GoTo doesnt
    Dim c As VBComponent
    Set c = proj.VBComponents(name)
    componentExists = True
    Exit Function
doesnt:
    componentExists = False
End Function


' Returns a reference to the workbook. Opens it if it is not already opened.
' Raises error if the file cannot be found.
Public Function openWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(filePath)
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath) 'can raise error
    End If
    Set openWorkbook = wb
End Function


' Returns the CodeName of the added sheet or an empty String if the workbook could not be opened.
Public Function addSheetToWorkbook(sheetName As String, workbookFilePath As String) As String
    Dim wb As Workbook
    On Error Resume Next 'can throw if given path does not exist
    Set wb = openWorkbook(workbookFilePath)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Dim ws As Worksheet
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = sheetName
        'ws.CodeName = sheetName: cannot assign to read only property
        Debug.Print "Sheet added " & sheetName
        addSheetToWorkbook = ws.CodeName
    Else
        Debug.Print "Skipping file " & sheetName & ". Could not open workbook " & workbookFilePath
        addSheetToWorkbook = ""
    End If
End Function

Public Sub exportVbProject()
    On Error GoTo exportVbProject_Error

    Dim project As VBProject
    Set project = ThisWorkbook.VBProject
    Build.exportVbaCode project
    Dim wb As Workbook
    Set wb = Build.openWorkbook(project.fileName)
    Build.exportNamedRanges wb
    MsgBox "Finished exporting code for: " & project.name

    Exit Sub
exportVbProject_Error:
    MsgBox ("Error Exporting" & vbNewLine & err.Description)
    'ErrorHandling.handleError "Menu.exportVbProject"
End Sub


Public Sub importVbProject()
    On Error GoTo importVbProject_Error
    
    Dim project As VBProject
    Set project = ThisWorkbook.VBProject
    
    Build.importVbaCode project
    Dim wb As Workbook
    Set wb = Build.openWorkbook(project.fileName)
    Build.importNamedRanges wb
    MsgBox "Finished importing code for: " & project.name

    On Error GoTo 0
    Exit Sub
importVbProject_Error:
    MsgBox ("Error Importing" & vbNewLine & err.Description)
End Sub




''''''''''''''''''  Named Ranges '''''''''''''''''''''''''




' Import named ranges from csv file
' Existing ranges with the same identifier will be replaced.
Public Sub importNamedRanges(wb As Workbook)
    Dim importDir As String
    importDir = Build.getSourceDir(wb.FullName, createIfNotExists:=False)
    If importDir = "" Then
        Debug.Print "No import directory for workbook " & wb.name & ", skipping"
        Exit Sub
    End If

    Dim fileName As String
    fileName = importDir & NAMED_RANGES_FILE_NAME
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(fileName) Then
        Dim inStream As TextStream
        Set inStream = FSO.OpenTextFile(fileName, ForReading, Create:=False)
        Dim line As String
        Do Until inStream.AtEndOfStream
            line = inStream.ReadLine
            importName wb, line
        Loop
        inStream.Close
    End If
End Sub


Private Sub importName(wb As Workbook, line As String)
    Dim parts As Variant
    parts = Split(line, ",")
    Dim rangeName As String, rangeAddress As String, comment As String
    rangeName = parts(columns.name)
    rangeAddress = parts(columns.RefersTo)
    comment = parts(columns.Comments)

    ' Existing namedRanges don't need to be removed first.
    ' wb.Names.Add will automatically replace or add the given namedRange.
    wb.Names.Add(rangeName, rangeAddress).comment = comment
End Sub


'Export named ranges to csv file
Public Sub exportNamedRanges(wb As Workbook)
    Dim exportDir As String
    exportDir = Build.getSourceDir(wb.FullName, createIfNotExists:=True)
    Dim fileName As String
    fileName = exportDir & NAMED_RANGES_FILE_NAME

    Dim lines As Collection
    Set lines = New Collection
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If hasValidRange(aName) Then
            'Only pull ranges from the 'SnowflakeConfig' worksheet
            If InStr(aName.value, gsSnowflakeConfigWorksheetName) Then
                lines.Add aName.name & "," & aName.RefersTo & "," & aName.comment
            End If
        End If
    Next
    If lines.Count > 0 Then
        'We have some names to export
        Debug.Print "writing to  " & fileName

        Dim FSO As New Scripting.FileSystemObject
        Dim outStream As TextStream
        Set outStream = FSO.CreateTextFile(fileName, overwrite:=True, Unicode:=False)
        On Error GoTo closeStream
        Dim line As Variant
        For Each line In lines
            outStream.WriteLine line
        Next line
closeStream:
        outStream.Close
    End If
End Sub


Private Function hasValidRange(aName As name) As Boolean
    On Error GoTo no
    hasValidRange = False
    Dim aRange As range
    Set aRange = aName.RefersToRange
    hasValidRange = True
no:
End Function


' Clean up all named ranges that don't refer to a valid range.
' This sub is not used by the import and export functions.
' It is provided only for convenience and can be run manually.
Public Sub removeInvalidNamedRanges(wb As Workbook)
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If Not hasValidRange(aName) Then
            aName.Delete
        End If
    Next
End Sub


