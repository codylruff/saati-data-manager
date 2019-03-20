Attribute VB_Name = "Factory"

Function CreateDictionary() As Dictionary
    Set CreateDictionary = New Dictionary
End Function

Function CreateSpecification() As Specification
    Set CreateSpecification = New Specification
End Function

Function CreateTemplate() As SpecTemplate
    Set CreateTemplate = New SpecTemplate
End Function

Function CreateSpecTemplate(spec_type As String) As SpecTemplate
    Dim template As SpecTemplate
    Set template = New SpecTemplate
    template.SpecType = spec_type
    Set CreateSpecTemplate = template
End Function

Function CreateTemplateFromJson(template As SpecTemplate, json_text As String) As SpecTemplate
    template.JsonToObject json_text
End Function

Function CreateSpecFromJson(spec As Specification, properties_json As String, tolerances_json As String) As Specification
    spec.JsonToObject properties_json, tolerances_json
    Set CreateSpecFromJson = spec
End Function

Function CreateConsoleBox(frm As UserForm) As ConsoleBox
    Dim obj As ConsoleBox
    Set obj = New ConsoleBox
    Set obj.FormId = frm
    Set CreateConsoleBox = obj
End Function

Function CreateWarp(frm As UserForm) As Warp
' Create warp object based on current_specification
    Dim w As Warp
    If manager.current_spec.SpecType = "warp" Then
        Set w = New Warp
        With w
            .Specification = manager.current_spec
            .NumberOfBobbins = frm.txtNumberOfBobbins
            .PackageWeightlbs = frm.txtPakageWeightlbs
            .WarpLengthYds = frm.txtWarpLength
        End With
        Set CreateWarp = w
    Else
        MsgBox "Material has no valid warping specification."
        Exit Function
    End If

End Function

Function CreateDatabaseRecord() As DatabaseRecord
' Creates a database record object
    Dim record: Set record = New DatabaseRecord
    Set CreateDatabaseRecord = record
End Function

Function CreateSQLiteDatabase() As SQLiteDatabase
' Creates a SQLite Database object
    Dim sqlite: Set sqlite = New SQLiteDatabase
    Set CreateSQLiteDatabase = sqlite
End Function
