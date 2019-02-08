Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================
Private Function GetTableName(IsDefault As Boolean) As String
    GetTableName = Iif(IsDefault, "standard_specifications", "modified_specifications")
End Function
    
Function GetSpec(ByRef materialId As String, IsDefault As Boolean) As DatabaseRecord
' Get a record(s) from the database

    Dim SQLstmt As String
    ' build the sql query
    SQLstmt = "SELECT Json_Text FROM " & GetTableName(IsDefault) & _
              " WHERE Material_Id= '" & materialID & "'" '&
              " AND Id= (" & _
              " SELECT max(Id) FROM " & GetTableName(IsDefault) & ")"
    Set GetRecords = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, SQLITE_PATH, sqlstmt)
End Function

Function PushSpec(ByRef spec As ISpec, IsDefault As Boolean)
' Push a new records
    Dim sqlstmt As String
    Dim tblName As String
    
    ' Create SQL statement from objects
    sqlstmt = "INSERT INTO " & tblName & " " & _
              "(Material_Id, Time_Stamp, Json_Text, Spec_Type) " & vbNewLine & _
              "VALUES ('" & spec.MaterialId & "', " & _
                      "'" & Now() & "', " & _
                      "'" & spec.ObjectToJson & ", " & _ 
                      "'" & spec.SpecType & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, sqlstmt
End Function

Private Function ExecuteSQLSelect(db As IDatabase, path As String, & _
                                                   SQLstmt As String) As DatabaseRecord
' Returns an table like array
    Dim record: Set record = New DatabaseRecord
    Debug.Print "-----------------------------------"
    Debug.Print SQLstmt
    db.openDb path
    db.selectQry SQLstmt
    record.data = db.data
    record.header = db.header
    Set ExecuteSQLSelect = record
End Function

Private Sub ExecuteSQL(db As IDatabase, path As String, SQLstmt As String)
' Performs update or insert querys returns error on select.
    Debug.Print "-----------------------------------"
    Debug.Print SQLstmt
    If Left(SQLstmt, 6) = "SELECT" Then
        Debug.Print ("Use ExecuteSQLite3Select() for SELECT query")
        Exit Sub
    Else
        db.openDb (path)
        db.execute (SQLstmt)
    End If
End Sub



