Attribute VB_Name = "TestHarness"
'This module executes automated test cases. It creates it's own data and then uploads it to Snowflake.
'It updates the data and performs different kinds of uploads

Const testHarnessModuleName As String = "TestHarness"
Dim wsTestWorksheet As Worksheet
Dim wsStatusWorksheet As Worksheet
Dim testCells() As testCell
Const uploadTestTable = "ExcelTestTable"
Const testWorksheet = "Test"
Const storedProc = "StoredProc"
Const vba = "VBA"
Public Type testCell
    name As String
    value As String
    datatype As String
    colPosition As Integer
    checkValue As Boolean
End Type

Sub executeTest()
    Dim originalResultsWS As String
    Dim originalUploadWS As String
    Dim originalUploadTable As String
    Dim originalMergeKeys As String

    If Utils.login Then
        'store orginal values so we can reset them later
        originalResultsWS = Utils.CustomRange(sgRangeResultsWorksheet)
        originalUploadWS = Utils.CustomRange(sgRangeUploadWorksheet)
        'originalUploadTable = Utils.CustomRange(sgRangeTableName)
        'originalMergeKeys = Utils.CustomRange(uploadMergeKeysRange)
        ' Set up worksheets and other vars for test
        Utils.CustomRange(sgRangeResultsWorksheet) = testWorksheet
        Utils.CustomRange(sgRangeUploadWorksheet) = testWorksheet
        Set wsTestWorksheet = Utils.getWorksheet(Utils.CustomRange(sgRangeResultsWorksheet))
        Set wsStatusWorksheet = Utils.getWorksheet(CustomRange(sgRangeLogWorksheet))
        'This is needed because the local VBA tests rely on the data types set properly so I've explicity set the types in the data
        Utils.execSQLFireAndForget ("alter session set TIMESTAMP_INPUT_FORMAT = 'MM/DD/YYYY HH24:MI:SS'")

        '********************************* Start Testing ***********************************************************
        '********************************* Stored Procs
        'Test 1 - Create new table with infering data types
        Call StatusForm.execMethod(testHarnessModuleName, "createTest", storedProc, False)
        ' Test 2 - Change data and merge
        Call StatusForm.execMethod(testHarnessModuleName, "mergeTest", storedProc)
        'Test 3 Truncate table and load
        Call StatusForm.execMethod(testHarnessModuleName, "truncateTest", storedProc)
        'Test 4 Append data to table
        Call StatusForm.execMethod(testHarnessModuleName, "appendTest", storedProc)
        'Test 5 - Create new table with explicity defined Data types
        Call StatusForm.execMethod(testHarnessModuleName, "createTest", storedProc, True)


        '********************************* VBA logic
        'Test 1 - Create new table with explicity defined Data types
        Call StatusForm.execMethod(testHarnessModuleName, "createTest", vba, True)
        ' Test 2 - Change data and merge
        Call StatusForm.execMethod(testHarnessModuleName, "mergeTest", vba)
        'Test 3 Truncate table and load
        Call StatusForm.execMethod(testHarnessModuleName, "truncateTest", vba)
        'Test 4 Append data to table
        Call StatusForm.execMethod(testHarnessModuleName, "appendTest", vba)
        '******************************** End Testing ************************************************************

        ' Set UploadDataForm = Nothing
        ' Reset all ranges back to the original values
        'Utils.CustomRange(sgRangeTableName) = originalUploadTable
        'Utils.CustomRange(uploadMergeKeysRange) = originalMergeKeys
        Utils.CustomRange(sgRangeResultsWorksheet) = originalResultsWS
        Utils.CustomRange(sgRangeUploadWorksheet) = originalUploadWS
        'Reset Date formats back.
        Call Utils.SetDateInputFormat
        'Drop table
        Call Utils.execSQLFireAndForget("Drop table " & uploadTestTable)
        Call StatusForm.Update_Status("Test cases executed successfully")
        StatusForm.Show
    End If
End Sub

Sub createTest(storedProcOrVBA As String, addExplicitDataTypes As Boolean)
    ' This test add data to the WS then creats a table
    ' If addExplicitDataTypes is true it will add a row for the Data Types, if not it won't and the data types will be inferred
    Dim i As Integer
    Dim sqlString As String
    Dim copType As String

    If storedProcOrVBA = storedProc Then
        copType = "recreateTable"
    Else
        copType = "RecreateLocal"
    End If

    ' clear the sheet
    ' Cells.Select
    ' Selection.Delete Shift:=xlUp
    wsTestWorksheet.Cells.Clear
    createTestData
    'update the worksheet with the data from the array
    If addExplicitDataTypes Then
        Call Load.AddDataTypeDropDowns
    End If
    Call updateWorksheetWithTestData(1, addExplicitDataTypes)
    'set upload table name
    Utils.CustomRange(sgRangeTableName) = uploadTestTable
    Call Load.uploadData(copType, uploadTestTable, "")
    Call StatusForm.Update_Status("Checking Datatypes...")

    ' Test datatpyes
    testDataTypes
    'Test for data values
    testDataValues

    Call StatusForm.Update_Status("Create Table Test Complete")
    Call StatusForm.Hide
End Sub

Sub mergeTest(storedProcOrVBA As String)
    Dim newSize As Integer
    Dim copType As String

    If storedProcOrVBA = storedProc Then
        copType = "merge"
    Else
        copType = "MergeLocal"
    End If
    'set merge Keys
    Utils.CustomRange(uploadMergeKeysRange) = "1"
    'Update one value
    testCells(2).value = "Goodbye"
    'Add another column
    newSize = UBound(testCells) + 2
    ReDim Preserve testCells(newSize)
    testCells(newSize - 1).name = "NumNew"
    testCells(newSize - 1).value = "100"
    testCells(newSize - 1).datatype = "NUMBER(38,0)"
    testCells(newSize - 1).checkValue = True

    testCells(newSize).name = "NumExplicit"
    testCells(newSize).value = "100.01"
    testCells(newSize).datatype = "NUMBER(38,2)"
    testCells(newSize).checkValue = True

    Call updateWorksheetWithTestData(2, False)

    'Add data type row
    Call Load.AddDataTypeDropDowns
    'Add data type for explicit declaration
    wsTestWorksheet.Cells(1, newSize - 1) = testCells(newSize - 1).datatype
    wsTestWorksheet.Cells(1, newSize) = testCells(newSize).datatype

    Call Load.uploadData(copType)

    ' test if update happened properly
    testDataValues
    Call StatusForm.Hide
End Sub
Sub truncateTest(storedProcOrVBA As String)
    Dim copType As String

    If storedProcOrVBA = storedProc Then
        copType = "truncate"
    Else
        copType = "TruncateLocal"
    End If
    Call Load.uploadData(copType)
    testDataValues
End Sub
Sub appendTest(storedProcOrVBA As String)
    Dim copType As String

    If storedProcOrVBA = storedProc Then
        copType = "append"
    Else
        copType = "AppendLocal"
    End If

    Call Load.uploadData(copType)
    testDataValues

End Sub

Sub createTestData()
    ReDim Preserve testCells(10)

    testCells(1).name = "Int1"
    testCells(1).value = "1"
    testCells(1).datatype = "NUMBER(38,0)"
    testCells(1).checkValue = False

    testCells(2).name = "Varchar1"
    testCells(2).value = "HELLO"
    testCells(2).datatype = "TEXT"
    testCells(2).checkValue = True

    testCells(3).name = "Bool1"
    testCells(3).value = "True"
    testCells(3).datatype = "BOOLEAN"
    testCells(2).checkValue = True

    testCells(4).name = "Date1"
    testCells(4).value = "1/1/2020"
    testCells(4).datatype = "DATE"
    testCells(4).checkValue = False

    testCells(5).name = "Date2"
    testCells(5).value = "11/11/2020"
    testCells(5).datatype = "DATE"
    testCells(5).checkValue = True

    testCells(6).name = "Date3"
    testCells(6).value = "2020-01-01"
    testCells(6).datatype = "DATE"
    testCells(6).checkValue = False

    testCells(7).name = "Time1"
    testCells(7).value = "13:00:00"
    testCells(7).datatype = "TIME"
    testCells(7).checkValue = False

    testCells(8).name = "Time2"
    testCells(8).value = "11:00:00 AM"
    testCells(8).datatype = "TIME"
    testCells(8).checkValue = False

    testCells(9).name = "Timestamp1"
    testCells(9).value = "1/1/2020 11:00:00 AM"
    testCells(9).datatype = "TIMESTAMP_NTZ"
    testCells(9).checkValue = True

    testCells(10).name = "NumPrec"
    testCells(10).value = "11.001"
    testCells(10).datatype = "NUMBER(38,3)"
    testCells(10).checkValue = True
End Sub

Sub testDataValues()
    Dim row As Integer
    ' Clear worksheet and then get data that was just uploaded
    wsTestWorksheet.Cells.Clear
    Call Query.ExecuteSelectAllFromUploadTable
    numberOfRows = wsTestWorksheet.UsedRange.Rows.Count
    For row = 2 To numberOfRows
        For i = 1 To UBound(testCells)
            If testCells(i).checkValue And CStr(wsTestWorksheet.Cells(row, i)) <> testCells(i).value Then
                MsgBox ("Create Table test failed. Incorrect values.")
                Stop
            End If
        Next
    Next
End Sub

Sub testDataTypes()
    Dim sqlString As String
    ' Get data types
    sqlString = "select column_name, case data_type when 'NUMBER' then " & _
            "data_type||'('||numeric_precision||','||numeric_scale||')' else data_type end as Datatype " & _
            "from information_schema.columns where table_name='" & UCase(uploadTestTable) & "' order by ordinal_position;"
    wsTestWorksheet.Cells.Clear
    Call Utils.ExecSQL(wsTestWorksheet, "A1", sqlString)

    ' Loop through the dataypes starting at the second line because the first one has the header
    For i = 1 To UBound(testCells)
        If wsTestWorksheet.Cells(i + 1, 2) <> testCells(i).datatype Then
            MsgBox ("Create Table test failed. Incorrect datatype.")
            Stop
        End If
    Next
End Sub

Sub updateWorksheetWithTestData(numberOfRows As Integer, addDataTypeRow As Boolean)
    Dim row As Integer
    Dim HeaderRowNumber As Integer

    ' clear all data on worksheet first
    wsTestWorksheet.Cells.Clear
    'If we need to add the Data type row, then the header row should be the 2nd row, else it's the 1st
    HeaderRowNumber = 1
    If addDataTypeRow Then
        HeaderRowNumber = 2
    End If
    For row = 1 To numberOfRows
        For i = 1 To UBound(testCells)
            If row = 1 Then
                If addDataTypeRow Then
                    wsTestWorksheet.Cells(1, i) = testCells(i).datatype
                End If 'if we are adding the data types then add the header to row 2
                wsTestWorksheet.Cells(HeaderRowNumber, i) = testCells(i).name
            End If
            ' if its the first column than treat it like the key and update it with the rownum
            If i = 1 Then
                wsTestWorksheet.Cells(row + HeaderRowNumber, i) = row
            Else
                wsTestWorksheet.Cells(row + HeaderRowNumber, i) = testCells(i).value
                If testCells(i).datatype = "TIMESTAMP_NTZ" Then
                    wsTestWorksheet.Cells(row + HeaderRowNumber, i).NumberFormat = "m/d/yyyy h:mm:ss"
                End If
            End If
        Next
    Next
End Sub
