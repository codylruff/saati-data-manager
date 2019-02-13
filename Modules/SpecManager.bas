Attribute VB_Name = "SpecManager"
'@Folder("Modules")
' CREATE-----------------------------
Function CreateDefaultSpec(spec As ISpec, materialID As String, MaterialDescription As String) As ISpec
' Maps properties onto an ISpec object by parsing a material code and description,
' then performing various calculations to determine any calcuable baseline properties.
    ' TODO: Implement function
    Set CreateDefaultSpec = spec
End Function

Sub AddNewSpec(spec As ISpec, IsDefault As Boolean)
' Add a new spec to the database
    DataAccess.PushRecord(spec, IsDefault)
End Sub

' READ-------------------------------
Function GetSpecJson(materialID As String, IsDefault As Boolean) As String
' Gets json text stored in a database that represents an ISpec Object
    Dim record As DatabaseRecord
    Set record = DataAccess.GetSpec(materialId, IsDefault)
    ' set the record objects fields dictionary in order access fields by name
    With record
        .SetDictionary
        GetSpecJson = .Fields.Item("Json_Text")
    End With
End Function

Function GetISpec(ByRef spec As ISpec, ByVal materialId As String) As ISpec
' Copy a specification object from the database
    ' check db tables 1 and 2 to get the most recent spec
    If IsDefault Then 
        Set GetISpec = spec.JsonToObject(GetSpecJson(materialID, _
                        "standard_specifications"))
    Else
        Set GetISpec = spec.JsonToObject(GetSpecJson(materialId, _
                        "modified_specifications"))
    End if
End Sub

Sub PrintSpecToConsole(frm As UserForm, ByRef spec As ISpec)
' Print object to console
    Dim key As Variant
    With frm.txtConsole
        ' Clear the console
        .Text = vbNullString
        
        For Each key In spec.Properties
            .Text = .Text & Utils.GetLine(Utils.SplitCamelCase(key.value), _
                        spec.Properties(key.value))
        Next key
    End With
End Sub

' TODO: This function is broken!!!
Sub DatabaseToWorksheet(db As IDatabase, path As String, tblName As String)
' Copies a database table into a worksheet
    Dim ws As Worksheet, record As DatabaseRecord
    Set record = DataAccess.Get(db, path, "SELECT * FROM " & tblName)
    ' Creates a sheet to store the database.
    Utils.ToggleExcelGui False
    Set ws = Sheets(Utils.CreateNewSheet(tblName & " " & Format(Now, "mm,dd,yyyy")))
    ' Creates the headers
    ws.Range(Cells(1, record.columns), Cells(1, record.columns)).value = record.header
    ' Copies in the data
    ws.Range(Cells(record.rows, record.colmns), Cells(record.rows + 1, record.columns)).value = record.data
    ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Utils.ToggleExcelGui True
End Sub

' UPDATE------------------------------


' DELETE

