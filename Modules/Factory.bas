Attribute VB_Name = "Factory"

Function CreateSpecification() As Specification
    Set CreateSpecification = New Specification
End Function

Function CreateSpecFromJson(spec As Specification, json_text As String) As Specification
    spec.JsonToObject json_text
    Debug.Print spec.Revision
    Debug.Print json_text
    Set CreateSpecFromJson = spec
End Function

Function CreateConsoleBox(frm As UserForm) As ConsoleBox
    Dim obj As ConsoleBox
    Set obj = New ConsoleBox
    Set obj.FormId = frm
    Set CreateConsoleBox = obj
End Function
