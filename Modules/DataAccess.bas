Attribute VB_Name = "DataAccess"
Option Explicit
'@Folder("Modules")
'===================================
'DESCRIPTION: Data Access Module
'===================================
Private Function GetTableName(IsStandard As Boolean) As String
    GetTableName = IIf(IsStandard, "standard_specifications", "modified_specifications")
End Function
    
Function GetSpec(ByRef MaterialId As String, IsStandard As Boolean) As DatabaseRecord
' Get a record(s) from the database

    Dim SQLstmt As String
    ' build the sql query
    SQLstmt = "SELECT * FROM " & GetTableName(IsStandard) & _
              " WHERE Material_Id= '" & MaterialId & "'"
    Set GetSpec = ExecuteSQLSelect( _
                     Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt)
End Function

Function PushSpec(ByRef spec As Specification, IsStandard As Boolean)
' Push a new records
    Dim SQLstmt As String
    Dim tblName As String
    
    ' Create SQL statement from objects
    SQLstmt = "INSERT INTO " & tblName & " " & _
              "(Material_Id, Time_Stamp, Properties_Json, Tolerances_Json, Spec_Type) " & vbNewLine & _
              "VALUES ('" & spec.MaterialId & "', " & _
                      "'" & CStr(Now()) & "', " & _
                      "'" & spec.PropertiesJson & ", " & _
                      "'" & spec.TolerancesJson & ", " & _
                      "'" & spec.SpecType & "')"
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt
End Function

Private Function ExecuteSQLSelect(db As IDatabase, path As String, _
                                                   SQLstmt As String) As DatabaseRecord
' Returns an table like array
    Dim record: Set record = New DatabaseRecord
    Logger.Log "-----------------------------------"
    Logger.Log SQLstmt
    db.openDb path
    db.selectQry SQLstmt
    record.data = db.data
    record.header = db.header
    Set ExecuteSQLSelect = record
End Function

Private Sub ExecuteSQL(db As IDatabase, path As String, SQLstmt As String)
' Performs update or insert querys returns error on select.
    Logger.Log "-----------------------------------"
    Logger.Log SQLstmt
    If Left(SQLstmt, 6) = "SELECT" Then
        Logger.Log ("Use ExecuteSQLSelect() for SELECT query")
        Exit Sub
    Else
        db.openDb (path)
        db.execute (SQLstmt)
    End If
End Sub


