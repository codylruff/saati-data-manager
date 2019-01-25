Attribute VB_Name = "DataAccess"
Option Explicit
'===================================
'DESCRIPTION: Data Access Functions
'===================================
' TODO: This module needs major refactoring
'       perhaps make a database class.
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

Function SQLite3SelectAll(tblName As String) As Collection
' Retrieve the entire table as a collection of objects
    Dim dbHandle    As Long
    Dim stmtHandle  As Long
    Dim RetVal      As Long
    Dim InitReturn  As Long
    Dim sqlQuery    As String
    Dim records     As Collection
    
    ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    InitReturn = SQLite3Initialize
    
    If InitReturn <> SQLITE_INIT_OK Then
        MsgBox "Error Initializing SQLite. Error: " & Err.LastDllError & "Contact Admin"
        Exit Function
    End If

    ' Open the database - getting a DbHandle back
    RetVal = SQLite3Open(PathToSQLite3Database, dbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    '--------------
    ' QUERY
    '--------------
    sqlQuery = "SELECT * FROM " & tblName & ";"

    RetVal = SQLite3PrepareV2(dbHandle, sqlQuery, stmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(stmtHandle)
    If RetVal = SQLITE_ROW Then
        Debug.Print "SQLite3Step Row Ready"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If

    Set records = New Collection
    ' Move to next row
    RetVal = SQLite3Step(stmtHandle)
    Do While RetVal = SQLITE_ROW
    
        ' Use this section to build an array of values
        ' then return it back to the caller.
        records.Add CreateDatabaseRecordFromTable(stmtHandle)
        RetVal = SQLite3Step(stmtHandle)
    Loop

    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(stmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    Set SQLite3SelectAll = records
End Function

Function CreateDatabaseRecordFromTable(ByVal stmtHandle As Long) As DatabaseRecord
' Creates a database record object from a row in the database
    Dim colCount    As Long
    Dim colName     As String
    Dim colType     As Long
    Dim colTypeName As String
    Dim colValue    As Variant
    Dim i           As Long
    Dim stepMsg     As String
    Dim record      As DatabaseRecord
    
    Set record = New DatabaseRecord
    colCount = SQLite3ColumnCount(stmtHandle)
    For i = 0 To colCount - 1
        colName = SQLite3ColumnName(stmtHandle, i)
        colType = SQLite3ColumnType(stmtHandle, i)
        colValue = ColumnValue(stmtHandle, i, colType)
        record.AddField colName
        record.AddValue colName, colValue
    Next
    Set CreateDatabaseRecord = record
End Function

Public Sub DatabaseToWorksheet(tblName As String)
' Copies a database table in a worksheet
    Dim exists As Boolean
    Dim shtName As String
    Dim i, k As Integer
    Dim field As Variant
    Dim records As Collection
    Dim record As DatabaseRecord
    Dim dict As Dictionary
    Dim rng As Range
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ' Creates a sheet to store the database.
    shtName = tblName & " " & Format(Now, "mm,dd,yyyy")
    With ThisWorkbook
        For i = 1 To Worksheets.count
            If Worksheets(i).name = shtName Then
                exists = True
            End If
        Next i
        If exists = True Then
            .Sheets(shtName).Delete
        End If
        .Sheets.Add(After:=.Sheets(.Sheets.count)).name = shtName
    End With
    ' Query database and create a collection of database records
    Set records = SQLite3SelectAll(tblName)
    ' Creates the headers
    k = 1
    Set record = records(1)
    Set dict = record.Fields
    For Each field In dict
        ActiveCell.value = field
        ActiveWorkbook.Names.Add _
            name:=field & "_", _
            RefersToR1C1:="=R1C" & k
        ActiveCell.offset(1, 0).Select
        ActiveCell.value = dict.item(field)
        ActiveCell.offset(-1, 1).Select
        k = k + 1
    Next field
    Set record = New DatabaseRecord
    records(1).Remove
    ' Copies the recordset into the sheet
    For Each record In records
        Set dict = record.Fields
        For Each field In dict
            Set rng = Worksheets(shtName).Range(field & "_")
            Utility.Update rng, dict.item(field)
        Next field
    Next record
    Range("A1").CurrentRegion.EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
'----------------------------------------------------
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

Function SqlCreateInsert(table As String, jsonText As String) As String
    Dim parser As New JSON13
    Dim doc As Variant
    Set doc = parser.parse(jsonText)
    
    Dim key As Variant
    Dim SQLstmt As String
    Dim INSERTstmt, VALUESstmt As String
    ' Set object properties

    ' Create the insert portion of the statement
    INSERTstmt = "INSERT INTO " & table & " (jsonText, "
    ' Create the values portion of the statement
    VALUESstmt = "VALUES ('" & jsonText & "', "

    For Each key In doc
        INSERTstmt = INSERTstmt & key & ", "
        VALUESstmt = VALUESstmt & "'" & doc.getString(key) & "', "
    Next key

    INSERTstmt = INSERTstmt & "Time_Stamp) "
    VALUESstmt = VALUESstmt & "'" & Now() & "')"

    ' Create SQL statement from object
    SQLstmt = INSERTstmt & vbNewLine & VALUESstmt
    Debug.Print SQLstmt
    RetVal = ExecuteSQLite3(SQLstmt)
    If Not RetVal = SQLITE_DONE Then Err.Raise Number:=1

End Function