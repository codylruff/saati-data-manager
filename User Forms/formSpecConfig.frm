VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formSpecConfig 
   Caption         =   "Specification Control"
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735
   OleObjectBlob   =   "formSpecConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formSpecConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private material_id As String

Private Sub UserForm_Initialize()
    material_id = InputBox("Enter a Material :", "Material Search")
    If SpecManager.ExecuteSearch(material_id) = SM_SEARCH_FAILURE Then
        MsgBox "Specification not found!", , "Null Spec Exception"
        Exit Sub
    End If
    SpecManager.PrintSpecification Me
    PopulateCboSelectProperty
    PopulateCboSelectRevision
    cboSelectRevision.value = App.current_spec.Revision
End Sub

Private Sub cmdBack_Click()
    Unload Me
    GuiCommands.GoToMain
End Sub

Private Sub cmdExportPdf_Click()
    GuiCommands.ConsoleBoxToPdf
End Sub

Private Sub cmdSaveChanges_Click()
    Dim RetVal As Long
    RetVal = SpecManager.SaveSpecification(App.current_spec)
    If RetVal <> COM_PUSH_COMPLETE Then
        Debug.Print "COM Server returned: ", RetVal
        MsgBox "New Specification Was Not Saved. Contact Admin."
    Else
        Debug.Print "COM Server returned: ", RetVal
        MsgBox "New Specification Succesfully Saved."
    End If
End Sub

Private Sub cmdSubmit_Click()
' This executes a set property command
' TODO: Change the name of this to cmdSetProperty
    With App.current_spec
        .Properties.Item(Utils.ConvertToCamelCase(cboSelectProperty.value)) = txtPropertyValue
        .Revision = .Properties.Item("Revision")
    End With
    SpecManager.PrintSpecification Me
End Sub

Private Sub cmdSearch_Click()
    Set App.current_spec = App.specs.Item(cboSelectRevision.value)
    SpecManager.PrintSpecification Me
End Sub

Private Sub PopulateCboSelectRevision()
    Dim rev As Variant
    With cboSelectRevision
        For Each rev In App.specs
            .AddItem rev
        Next rev
    End With
End Sub

Private Sub PopulateCboSelectProperty()
    Dim key As Variant
    With cboSelectProperty
        For Each key In App.current_spec.Properties
          .AddItem Utils.SplitCamelCase(CStr(key))
        Next key
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


