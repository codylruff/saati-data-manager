Attribute VB_Name = "SpecManager"
'// This object allows information to persist throughout the Application lifecycle
Public manager As App

Public Sub StartSpecManager()
    Set manager = New App
End Sub

Public Sub StopSpecManager()
    Set manager = Nothing
End Sub

Function TemplateInput() As String
    Dim template_name As String
    template_name = InputBox("Enter a template name :", "Custom Template Name")
    If template_name = vbNullString Then
        MsgBox "Must Enter a template name."
    End If
    TemplateInput = template_name
End Function

Sub MaterialInput(material_id As String)
' Takes user input for material search
    Dim ret_val As Long
    If material_id = vbNullString Then Exit Sub
    If SpecManager.ExecuteSearch(material_id) = SM_SEARCH_FAILURE Then
        MsgBox "Specification not found!", , "Null Spec Exception"
        Exit Sub
    End If
End Sub

Function ExecuteSearch(material_id As String) As Long
' Manages the search procedure
    Set manager.standard = SpecManager.GetStandard(material_id)
    If manager.standard Is Nothing Then
        ExecuteSearch = SM_SEARCH_FAILURE
    Else
        Set manager.specs = SpecManager.GetSpec(material_id)
        Set manager.current_spec = GetLatestSpec(manager.specs)
        ' Return 0/1 on success/failure
        ExecuteSearch = SM_SEARCH_SUCCESS
    End If
End Function

Function GetTemplate(template_name As String) As SpecTemplate
    Dim template As SpecTemplate
    Dim json As String
    Set template = Factory.CreateTemplate(template_name)
    json = ComService.GetSpecTemplate(template.SpecType)
    If json <> vbNullString Then
        Set GetTemplate = Factory.CreateTemplateFromJson( _
            template:=template, _
            json_text:=json)
    Else
        Set GetTemplate = Nothing
    End If

End Function

Function GetStandard(material_id As String) As Specification
    Dim spec As Specification
    Dim json_coll As Collection
    Dim json_dict As Dictionary
    Set spec = Factory.CreateSpecification()
    ' TODO: Apply template to spec object
    spec.IsStandard = True
    Set json_coll = Main.GetSpecJson(MaterialInputValidation(material_id), True)
    If json_coll Is Nothing Then
        Set GetStandard = Nothing
        Exit Function
    Else
        For Each json_dict In json_coll
            Set spec = Factory.CreateSpecification
                With spec
                    spec.MaterialId = json_dict.Item("Material_Id")
                    spec.SpecType = json_dict.Item("Spec_Type")
                    spec.Revision = json_dict.Item("Revision")
                End With
            Set spec = Factory.CreateSpecFromJson(spec, json_dict.Item("Properties_Json"), json_dict.Item("Tolerances_Json"))
        Next json_dict
        Set GetStandard = spec
    End If
End Function

Function GetSpec(material_id As String) As Object
    Dim json_coll As Collection
    Dim json_dict, specs_dict As Dictionary
    Dim spec As Specification
    Dim rev As String
    Dim key As Variant
    Set json_coll = Main.GetSpecJson(MaterialInputValidation(material_id), False)
    Set specs_dict = New Dictionary
    
    If json_coll Is Nothing Then
        Set GetSpec = Nothing
        Exit Function
    Else
        specs_dict.Add manager.standard.Revision, manager.standard
        Logger.Log manager.standard.Revision
        Set spec = manager.standard
        rev = manager.standard.Revision
        For Each json_dict In json_coll
                Set spec = Factory.CreateSpecification
                With spec
                    spec.MaterialId = json_dict.Item("Material_Id")
                    spec.SpecType = json_dict.Item("Spec_Type")
                    spec.Revision = json_dict.Item("Revision")
                End With
                Set spec = Factory.CreateSpecFromJson(spec, json_dict.Item("Properties_Json"), json_dict.Item("Tolerances_Json"))
                Logger.Log spec.MaterialId & " : " & spec.Revision
                specs_dict.Add spec.Revision, spec
                rev = spec.Revision
        Next json_dict
        specs_dict.Item(rev).IsLatest = True
        Set GetSpec = specs_dict
    End If

End Function

Sub PrintSpecification(frm As MSForms.UserForm)
    Set manager.console = Factory.CreateConsoleBox(frm)
    manager.console.PrintObject manager.current_spec
End Sub

Function SaveSpecification(spec As Specification) As Long
    SaveSpecification = IIf(ComService.PushSpecJson(spec, False) = COM_PUSH_COMPLETE, _
                            COM_PUSH_COMPLETE, COM_PUSH_FAILURE)
End Function

Function SaveStandardSpecification(spec As Specification) As Long
    SaveStandardSpecification = ComService.PushSpecJson(spec, True)
End Function

Function SaveSpecTemplate(template As SpecTemplate) As Long
    SaveSpecTemplate = ComService.PushSpecTemplate(template)
End Function

Private Function MaterialInputValidation(material_id As String)
' Ensures that the material id input by the user is parseable.
' TODO: This function is awful need to refactor unsure how due to the
'       ridiculous lack of uniqueness in the database.
'       "The style 101 problem"
    If (material_id <> "101") And (Mid(material_id, 5, 3) <> "101") Then
        MaterialInputValidation = material_id
        Exit Function
    End If
    If Len(material_id) >= 5 Then
        MaterialInputValidation = Mid(material_id, 5, 3) & Mid(material_id, 2, 2)
    Else
        Dim question As Integer
        question = MsgBox("Click Yes for Style 101 Kevlar or No for Hyosung.", vbYesNo + vbQuestion, "Style 101 has two version")
        If question = vbYes Then
            MaterialInputValidation = "101" & "KE"
        Else
            MaterialInputValidation = "101" & "HY"
        End If
    End If
End Function

Function GetLatestSpec(specs As Object) As Specification
    Dim key As Variant
    For Each key In manager.specs
        If manager.specs.Item(key).IsLatest = True Then
            Set GetLatestSpec = manager.specs.Item(key)
        End If
    Next key
End Function
