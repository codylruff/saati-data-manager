Attribute VB_Name = "Main"
'// This object allows information to persist throughout the Application lifecycle
'Public manager As App

'@Folder("Modules")
' CREATE-----------------------------

Sub AddNewSpec(spec As Specification)
' Add a new spec to the database
    DataAccess.PushRecord(spec, spec.IsStandard)
End Sub

' READ-------------------------------
Function GetSpecJson(MaterialId As String, IsDefault As Boolean) As Collection
' Gets json text stored in a database that represents an Specification Object
    Dim record As DatabaseRecord
    Set record = DataAccess.GetSpec(MaterialId, IsDefault)
    ' set the record objects fields dictionary in order access fields by name
        record.SetDictionary
        Set GetSpecJson = record.records
End Function

Function GetSpecification(ByRef spec As Specification, ByVal MaterialId As String) As Specification
' Copy a specification object from the database
    ' check db tables 1 and 2 to get the most recent spec
    If IsDefault Then
        Set GetSpecification = spec.JsonToObject(GetSpecJson(MaterialId, _
                        "standard_specifications"))
    Else
        Set GetSpecification = spec.JsonToObject(GetSpecJson(MaterialId, _
                        "modified_specifications"))
    End If
End Sub

Sub PrintSpecToConsole(frm As UserForm, ByRef spec As Specification)
' Print object to console
    Dim key As Variant
    With frm.txtConsole
        ' Clear the console
        .text = vbNullString
        
        For Each key In spec.Properties
            .text = .text & Utility.GetLine(Utility.SplitCamelCase(key.value), _
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
    Utility.ToggleExcelGui False
    Set ws = Sheets(Utility.CreateNewSheet(tblName & " " & Format(Now, "mm,dd,yyyy")))
    ' Creates the headers
    ws.Range(Cells(1, record.columns), Cells(1, record.columns)).value = record.header
    ' Copies in the data
    ws.Range(Cells(record.rows, record.colmns), Cells(record.rows + 1, record.columns)).value = record.data
    ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Utility.ToggleExcelGui True
End Sub

' UPDATE------------------------------


' DELETE
