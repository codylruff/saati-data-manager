Attribute VB_Name = "DataHistory"
Option Explicit
'=================================
'NOTES: This module contains all
'data base subs/funcs.
'=================================

Function ExecuteSQLite3Select(SQLstmt As String) As DatabaseRecord
    Dim RetVal As Long
    Dim recordsAffected As Long
    Dim InitReturn As Long
    Dim ErrorVal As Long
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    Dim colCount As Long
    Dim colName As String
    Dim colType As Long
    Dim colTypeName As String
    Dim colValue As Variant
    Dim i As Long
    Dim stepMsg As String
    Dim record As DatabaseRecord

    Set record = New DatabaseRecord
    ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    InitReturn = SQLite3Initialize
    
    If InitReturn <> SQLITE_INIT_OK Then
        MsgBox "Error Initializing SQLite. Error: " & Err.LastDllError & "Contact Admin"
        Exit Function
    End If

    ' Open the database - getting a DbHandle back
    RetVal = SQLite3Open(PathToSQLite3Database, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal

    '-------------------------
    ' Select statement
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, SQLstmt, myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        Debug.Print "SQLite3Step Row Ready"

        ' Use this section to build an array of values
        ' then return it back to the caller.
        colCount = SQLite3ColumnCount(myStmtHandle)
        Debug.Print "Column count: " & colCount

        For i = 0 To colCount - 1
            colName = SQLite3ColumnName(myStmtHandle, i)
            colType = SQLite3ColumnType(myStmtHandle, i)
            colValue = ColumnValue(myStmtHandle, i, colType)
            record.AddField colName
            record.AddValue colName, colValue
            Debug.Print "Column " & i & ":", colName, colTypeName, colValue
        Next
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)

    Set ExecuteSQLite3Select = record

End Function

Function ExecuteSQLite3(SQLstmt As String) As Long
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    Dim RetVal As Long
    Dim recordsAffected As Long
    Dim InitReturn As Long
    Dim ErrorVal As Long
    ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    InitReturn = SQLite3Initialize
    
    If InitReturn <> SQLITE_INIT_OK Then
        MsgBox "Error Initializing SQLite. Error: " & Err.LastDllError & "Contact Admin"
        Exit Function
    End If

    Dim stepMsg As String
    
    ' Open the database - getting a DbHandle back
    RetVal = SQLite3Open(PathToSQLite3Database, myDbHandle)
    
    '------------------------
    ' Execute SQLstmt
    ' ================
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, SQLstmt, myStmtHandle)
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
        ExecuteSQLite3 = RetVal
    Else
        Debug.Print "SQLite3Step returned " & RetVal
        ExecuteSQLite3 = RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    RetVal = SQLite3Close(myDbHandle)
End Function

Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

