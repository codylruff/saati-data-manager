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

Public console As ConsoleBox

Private Sub cmdBack_Click()
    Unload Me
    GoToMain
End Sub

Private Sub cmdSaveChanges_Click()
    SpecManager.SaveSpecification current_spec
End Sub

Private Sub cmdSubmit_Click()
    With spec
        .Properties.Item(Utils.ConvertToCamelCase(cboSelectProperty.value)) = txtPropertyValue
        .Revision = .Properties.Item("Revision")
    End With
    console.PrintObject current_spec
End Sub

Private Sub cmdSearch_Click()
    Dim key As Variant
    Dim material_id As String
    If (txtSAPcode.value <> "101") And (Mid(txtSAPcode.value, 5, 3) <> "101") Then
        material_id = txtSAPcode.value
    Else
        material_id = SpecManager.MaterialInputValidation(txtSAPcode.value)
    End If
    Set standard = SpecManager.GetStandard(material_id)
    standard.IsStandard = True
    Set specs = SpecManager.GetSpec(material_id)
    For Each key In specs
        If specs.Item(key).IsLatest = True Then
            Set current_spec = specs.Item(key)
        End If
    Next key
    If current_spec Is Nothing Then
        MsgBox "Specification not found!", , "Null Spec Exception"
        Exit Sub
    Else
        Set console = Factory.CreateConsoleBox(Me)
        console.PrintObject current_spec
        PopulateCboSelectProperty
        txtPropertyValue.value = vbNullString
    End If
End Sub

Private Sub PopulateCboSelectProperty()
    Dim key As Variant
    With cboSelectProperty
        For Each key In current_spec.Properties
          .AddItem Utils.SplitCamelCase(CStr(key))
        Next key
    End With
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearForm Me
End Sub

Private Sub UserForm_Initialize()
' Constructor
    Set server = CreateObject("DM_LIB.DmComServer")
    Set current_spec = New Specification
    Set standard = New Specification
    Set specs = CreateObject("Scripting.Dictionary")
    Set console = New ConsoleBox
End Sub

Private Sub UserForm_Terminate()
' Deconstructor
    Set server = Nothing
    Set current_spec = Nothing
    Set standard = Nothing
    Set specs = Nothing
    Set console = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' This
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub


