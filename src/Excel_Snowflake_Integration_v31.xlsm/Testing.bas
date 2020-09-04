Attribute VB_Name = "Testing"
Sub putFile()
    Dim statusWorksheet As Worksheet
    Dim connString As String
    Dim sqlString As String
    Set statusWorksheet = Sheets(CustomRange(sgRangeLogWorksheet))
    connString = Utils.ConnectionString()
    Call Utils.RemoveQueryTables(statusWorksheet)
    statusWorksheet.Cells.Clear
    sqlString = "put 'file:////Users/ssegal/Documents/test1111.csv'  @my_internal_stage;"
    Call Utils.ExecSQL(statusWorksheet, connString, nextStatusCellToLoad, sqlString)
End Sub
