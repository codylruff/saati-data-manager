Attribute VB_Name = "SpecManager"

Function GetStandard(material_id As String) As Specification
    Set GetStandard = Factory.CreateSpecFromJson(Factory.CreateSpecification, ComService.GetStandardJson(material_id))
End Function

Function GetSpec(material_id As String) As Object
    Dim json_dict, specs_dict As Dictionary
    Dim spec As Specification
    Dim rev As String
    Dim key As Variant
    
    Set json_dict = JsonConverter.ParseJson(ComService.GetSpecJson(material_id))
    Set specs_dict = New Dictionary
    
    If json_dict Is Nothing Then
        Set GetSpec = Nothing
        Exit Function
    Else
        specs_dict.Add standard.Revision, standard
        set spec = standard
        rev = standard.Revision
        For Each key In json_dict
            Set spec = Factory.CreateSpecFromJson(Factory.CreateSpecification, json_dict.Item(key))
            specs_dict.Add spec.Revision, spec
            rev = spec.Revision
        Next key
        specs_dict.Item(rev).IsLatest = True
        Set GetSpec = specs_dict
    End If

End Function

Sub SaveSpecification(spec As Specification)
    Dim return_value As Long
    return_value = ComService.PushSpecJson(spec)
    If return_value <> COM_PUSH_COMPLETE Then
        Debug.Print "COM Server returned: ", return_value
    Else
        MsgBox "New Specification Succesfully Saved."
    End If
End Sub

Sub SaveStandardSpecification(spec As Specification)
    Dim return_value As Long
    return_value = ComService.PushSpecJson(spec, True)
End Sub

Function MaterialInputValidation(material_id As String)
    Dim correct_id As String
    If Len(material_id) >= 5 Then
        correct_id = Mid(material_id, 5, 3) & Mid(material_id, 2, 2)
    Else
        Dim question As Integer
        question = MsgBox("Click Yes for Style 101 Kevlar or No for Hyosung.", vbYesNo + vbQuestion, "Style 101 has two version")
        If question = vbYes Then
            correct_id = "101" & "KE"
        Else
            correct_id = "101" & "HY"
        End If
    End If
    MaterialInputValidation = correct_id
End Function
