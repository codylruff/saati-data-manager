VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCreateGeneric 
   Caption         =   "Create New Spec Type"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9285
   OleObjectBlob   =   "formCreateGeneric.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCreateGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private template_name As String

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub UserForm_Initialize()
    Set App = New App
    template_name = SpecManager.TemplateInput
    If template_name <> vbNullString Then
      Set App.current_template = Factory.CreateSpecTemplate(template_name)
      Set App.console = Factory.CreateConsoleBox(Me)
      lblInstructions.Caption = " Instructions :" & vbNewLine & _
            " Create the template parameters one at a time," & _
            " selecting a parameter type (text, number, True/False)," & _
            " entering the parameter name, and clicking the 'Set Property'" & _
            " button. The template can be saved by clicking the save button."
   Else
      GuiCommands.UnloadAllForms
      Err.Raise (1)
   End If
End Sub

Private Sub cmdAddProperty_Click()
   App.console.PrintLine Me.txtPropertyName
   App.current_template.AddProperty Utils.ConvertToCamelCase(Me.txtPropertyName)
End Sub

Private Sub cmdSubmitTemplate_Click()
   App.current_template.Revision = 0
   If SpecManager.SaveSpecTemplate(App.current_template) <> COM_PUSH_COMPLETE Then
      Debug.Print "COM Server returned: ", COM_PUSH_FAILURE
        MsgBox "New Template Was Not Saved. Contact Admin."
    Else
        Debug.Print "COM Server returned: ", COM_PUSH_COMPLETE
        MsgBox "New Template Succesfully Created."
    End If
End Sub

Private Sub UserForm_Terminate()
    Set App = Nothing
End Sub
