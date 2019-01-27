Attribute VB_Name = "DataAccess"
Option Explicit
'===================================
'DESCRIPTION: Data Access Module
'===================================

Function ExecuteSQLSelect(db As IDatabase, path As String, SQLstmt As String) As DatabaseRecord
' Returns an table like array
    Dim record: Set record = New DatabaseRecord
    db.openDb path
    db.selectQry SQLstmt
    record.data = db.data
    record.header = db.header
    ExecuteSQLSelect = db.data
End Function

Sub ExecuteSQL(db As IDatabase, path As String, SQLstmt As String)
' Performs update or insert querys returns error on select.
    If Left(SQLstmt,6) = "SELECT" Then
        Debug.Print("Use ExecuteSQLite3Select() for SELECT query")
        Exit Sub
    Else
        db.openDb(path)
        db.execute(SQLstmt)
    End If
End Sub

Sub DatabaseToWorksheet(db As IDatabase, path As String, tblName As String)
' Copies a database table into a worksheet
    Dim shtName As String
    Dim ws As Worksheet
    Dim record: Set record = ExecuteSQLSelect(db, path, "SELECT * FROM " & tblName)
    ' Disables unpleasent ui effects
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ' Creates a sheet to store the database.
    shtName = Utility.CreateNewSheet(tblName & " " & Format(Now, "mm,dd,yyyy"))
    Set ws = Sheets(shtName)
    ' Creates the headers
    ws.Range(Cells(1, 1), Cells(1, record.columns)).Value = record.header
    ' Copies in the data
    ws.Range(Cells(2, 1), Cells(record.rows + 1, record.columns)).Value = record.data
    ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    ' Re-enable ui updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

