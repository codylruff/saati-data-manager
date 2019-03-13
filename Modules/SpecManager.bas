Attribute VB_Name = "SpecManager"
'// This object allows information to persist throughout the Application lifecycle
Public App As App

Function TemplateInput() As String
    Dim template_name As String
    template_name = InputBox("Enter a template name :", "Custom Template Name")
    If template_name = vbNullString Then
        MsgBox "Must Enter a template name."
    End If
    TemplateInput = template_name
End Function

Function MaterialInput() As String
' Takes user input for material search
    Dim ret_val As Long
    Dim inpt As String
    ret_val = SM_SEARCH_FAILURE
        Do While ret_val = SM_SEARCH_FAILURE
            inpt = InputBox("Enter a Material :", "Material Search", "Search Materials")
            If inpt = vbNullString Then
                 MaterialInput = inpt
                 Exit Function
            Else
                ret_val = SpecManager.ExecuteSearch(inpt)
                If ret_val = SM_SEARCH_FAILURE Then
                   MsgBox "Specification not found!", , "Null Spec Exception"
                End If
            End If
        Loop
        MaterialInput = inpt
End Function

Function ExecuteSearch(material_id As String) As Long
' Manages the search procedure
    Set App.standard = SpecManager.GetStandard(material_id)
    If App.standard Is Nothing Then
        ExecuteSearch = SM_SEARCH_FAILURE
    Else
        Set App.specs = SpecManager.GetSpec(material_id)
        Set App.current_spec = GetLatestSpec(App.specs)
        ' Return 0/1 on success/failure
        ExecuteSearch = SM_SEARCH_SUCCESS
    End If
End Function

Function GetStandard(material_id As String) As Specification
    Dim spec_ As Specification
    Dim json As String
    Set spec_ = Factory.CreateSpecification()
    spec_.IsStandard = True
    json = ComService.GetStandardJson(MaterialInputValidation(material_id))
    If json <> vbNullString Then
        Set GetStandard = Factory.CreateSpecFromJson( _
            spec:=spec_, _
            json_text:=json)
    Else
        Set GetStandard = Nothing
    End If
End Function

Function GetSpec(material_id As String) As Object
    Dim json_dict, specs_dict As Dictionary
    Dim spec As Specification
    Dim rev As String
    Dim key As Variant
    
    Set json_dict = JsonConverter.ParseJson(ComService.GetSpecJson( _
                    MaterialInputValidation(material_id)))
    Set specs_dict = New Dictionary
    
    If json_dict Is Nothing Then
        Set GetSpec = Nothing
        Exit Function
    Else
        specs_dict.Add App.standard.Revision, App.standard
        Set spec = App.standard
        rev = App.standard.Revision
        For Each key In json_dict
            Set spec = Factory.CreateSpecFromJson(Factory.CreateSpecification, json_dict.Item(key))
            specs_dict.Add spec.Revision, spec
            rev = spec.Revision
        Next key
        specs_dict.Item(rev).IsLatest = True
        Set GetSpec = specs_dict
    End If

End Function

Sub PrintSpecification(frm As MSForms.UserForm)
    Set App.console = Factory.CreateConsoleBox(frm)
    App.console.PrintObject App.current_spec
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
    For Each key In App.specs
        If App.specs.Item(key).IsLatest = True Then
            Set GetLatestSpec = App.specs.Item(key)
        End If
    Next key
End Function
