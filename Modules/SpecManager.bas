Attribute VB_Name = "SpecManager"

'@Folder("Modules")
' CREATE-----------------------------
Function CreateDefaultSpec(spec As ISpec, materialID As String, MaterialDescription As String) As ISpec
' Maps properties onto an ISpec object by parsing a material code and description,
' then performing various calculations to determine any calcuable baseline properties.
    Dim dict: Set dict = New Dictionary
    With dict
        .Add key:="MaterialNumber", Item:=materialID
        .Add key:="MaterialDescription", Item:=MaterialDescription
        .Add key:="Style", Item:=CLng(Mid(materialID, 6, 3))
    End With
    With spec
        .JsonToObject (JsonConverter.ConvertToJson(dict))
        If .SpecType = "Warping" Then
            Dim warpSpec: Set warpSpec = New WarpingSpecification
            Dim Style: Set Style = New StyleSpecification
            Set warpSpec = spec
            warpSpec.StyleSpec = GetISpec(Style, dict.Item("Style"), "style_specifications")
        End If
        .SetDefaultProperties
    End With
    Set CreateDefaultSpec = spec
End Function

Sub AddNewSpec(spec As ISpec)
' Add a new spec to the database
    Dim key As Variant
    Dim SQLstmt As String
    Dim INSERTstmt, VALUESstmt As String
    Dim jsonText As String
    ' Set object properties if nothing
    If spec.SpecType = "warping" Or spec.SpecType = "style" Then
        jsonText = JsonConverter.ConvertToJson(spec.Properties)
        ' Create the insert portion of the statement
        INSERTstmt = "INSERT INTO " & _
                     spec.SpecType & "_specifications (" & _
                     "Material_Id, Time_Stamp, Json_Text) "
        ' Create the values portion of the statement
        VALUESstmt = "VALUES ('" & spec.Properties.Item("Style") & "', " & _
                             "'" & Now() & "', " & _
                             "'" & jsonText & "')"
    End If
    
    ' Create SQL statement from objects
    SQLstmt = INSERTstmt & vbNewLine & VALUESstmt
    Debug.Print "-----------------------------------"
    Debug.Print SQLstmt
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt
End Sub

Private Sub MassLoadSpecifications()
' One time use to mass upload specs to the database
    Dim i As Integer
    Dim code As Range
    Dim Description As Range
    Dim ws As Worksheet

    Set ws = Sheets("Materials")
    
    For i = 2 To 80
        Set code = ws.Cells(i, 1)
        Debug.Print code.value
        Set Description = ws.Cells(i, 2)
        AddNewSpec (CreateDefaultSpec(Factory.CreateWarpingSpecification, _
                        code.value, Description.value))
    Next i

End Sub

Private Sub MassLoadStyles()
' One time use to mass upload styles to the database
    Dim i As Long
    Dim query As String
    Dim db As SQLiteDatabase
    Set db = Factory.CreateSQLiteDatabase
    Dim path As String
    path = "S:\Data Manager\Protection_Quality_Control.db3"
    Dim record: Set record = New DatabaseRecord
    Dim Style As StyleSpecification
    For i = 1 To 207
        query = "SELECT * FROM " & "tblStyleSpecs " & _
                "WHERE style_id=" & i
        Set record = DataAccess.ExecuteSQLSelect(db, path, query)
        record.SetDictionary
        Set Style = Factory.CreateStyleSpecification
        Style.ISpec_JsonToObject JsonConverter.ConvertToJson(record.Fields)
        AddNewSpec Style
    Next i
End Sub

' READ-------------------------------

Function GetSpecJson(materialID As String, tblName As String) As String
' Gets json text stored in a database that represents an ISpec Object
    Dim db As SQLiteDatabase
    Set db = New SQLiteDatabase
    Dim record As DatabaseRecord
    Dim SQLstmt As String
    ' build the sql query
    SQLstmt = "SELECT Json_Text FROM " & tblName & _
              " WHERE Material_Id= '" & materialID & "'" '&
              '" AND Id= (" & _
              '" SELECT max(Id) FROM " & tblName & ")"
    Set record = ExecuteSQLSelect(db, Constants.SQLITE_PATH, SQLstmt)
    ' set the record objects fields dictionary in order access fields by name
    With record
        .SetDictionary
        GetSpecJson = .Fields.Item("Json_Text")
    End With
End Function

Sub GetISpec(ByRef spec As ISpec, ByVal materialID As String, ByVal tblName As String)
    spec.JsonToObject GetSpecJson(materialID, tblName)
End Sub

Sub PrintSpecToConsole(frm As UserForm, ByRef spec As ISpec)
' Print object to console
    Dim key As Variant
    With frm.txtConsole
        ' Clear the console
        .Text = vbNullString
        
        For Each key In spec.Properties
            .Text = .Text & Utility.GetLine(Utility.SplitCamelCase(key.value), spec.Properties(key.value))
        Next key
    End With
End Sub


' UPDATE------------------------------


' DELETE

