VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formSpecConfig 
   Caption         =   "Specification Control"
   ClientHeight    =   11865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810
   OleObjectBlob   =   "formSpecConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formSpecConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    Dim ret_val As String
    Logger.Log "--------- Start " & Me.Name & " ----------"
End Sub

Private Sub cmdMaterialSearch_Click()
    SpecManager.MaterialInput txtMaterialId
    SpecManager.PrintSpecification Me
    PopulateCboSelectProperty
    PopulateCboSelectRevision
    cboSelectRevision.value = manager.current_spec.Revision
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
        Logger.Log "COM Server returned: ", COM_PUSH_FAILURE
        MsgBox "New Specification Was Not Saved. Contact Admin."
    Else
        Logger.Log "COM Server returned: ", COM_PUSH_COMPLETE
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

Private Sub PopulateCboSelectRevision()
    Dim rev As Variant
    With cboSelectRevision
        For Each rev In manager.specs
            .AddItem rev
        Next rev
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
    Logger.Log "--------- End " & Me.Name & " ----------"
End Sub
