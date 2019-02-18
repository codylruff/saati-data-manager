Attribute VB_Name = "SpecManager"

Function GetSpec(material_id As String) As Specification
    Dim json_text As String
    json_text = ComService.GetSpecJson(material_id)
    If json_text = VbNullString Then
        Set GetSpec = Nothing
    Else
        Set GetSpec = CreateSpecFromJson(Factory.CreateSpecification, json_text)
    End If
End Function

Private Function CreateSpecFromJson(spec As Specification, json_text As String)
    Debug.Print json_text
    spec.JsonToObject json_text
    Set CreateSpecFromJson = spec
End Function
