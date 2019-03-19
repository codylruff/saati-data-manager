VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCreateSpec 
   Caption         =   "Specification Control"
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735
   OleObjectBlob   =   "formCreateSpec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCreateSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()
    manager.Logger.Log "--------- " & Me.Name & " ----------"
    PopulateCboSelectSpecType
End Sub

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdExportPdf_Click()
    GuiCommands.ConsoleBoxToPdf
End Sub

Private Sub cmdSaveChanges_Click()
' Calls method to save a new specification incremented the revision by +0.1)
    If SpecManager.SaveSpecification(manager.current_spec) <> COM_PUSH_COMPLETE Then
        manager.Logger.Log "COM Server returned: ", COM_PUSH_FAILURE
        MsgBox "New Specification Was Not Saved. Contact Admin."
    Else
        manager.Logger.Log "COM Server returned: ", COM_PUSH_COMPLETE
        MsgBox "New Specification Succesfully Saved."
    End If
End Sub

Private Sub cmdSubmit_Click()
' This executes a set property command
' TODO: Change the name of this to cmdSetProperty
    With manager.current_spec
        .Properties.Item(Utils.ConvertToCamelCase( _
                cboSelectProperty.value)) = txtPropertyValue
        .Revision = .Properties.Item("Revision")
    End With
    SpecManager.PrintSpecification Me
End Sub

Private Sub cmdSearch_Click()
    Set manager.current_spec = manager.specs.Item(cboSelectRevision.value)
    SpecManager.PrintSpecification Me
End Sub

Private Sub PopulateCboSelectSpecType()
' TODO: This should pull from a database or textfile or something.
    With cboSelectRevision
        .AddItem "warping"
        .AddItem "style"
        .AddItem "fabric"
    End With
End Sub

Private Sub PopulateCboSelectProperty()
    Dim prop As Variant
    With cboSelectProperty
        For Each prop In manager.current_spec.Properties
          .AddItem Utils.SplitCamelCase(CStr(prop))
        Next prop
    End With
    txtPropertyValue.value = vbNullString
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearForm Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' This
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    Set manager = Nothing
End Sub
