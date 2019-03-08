Attribute VB_Name = "Factory"

Function CreateSpecification() As Specification
    Set CreateSpecification = New Specification
End Function

Function CreateSpecTemplate(spec_type As String) As SpecTemplate
    Dim template As SpecTemplate
    Set template = New SpecTemplate
    template.SpecType = spec_type
    Set CreateSpecTemplate = template
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

Function CreateWarp(frm As UserForm) As Warp
' Create warp object based on current_specification
    Dim w As Warp
    If App.current_spec.SpecType = "warp" Then
        Set w = New Warp
        With w
            .Specification = App.current_spec
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
