VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formSpecConfig 
   Caption         =   "Specification Control"
   ClientHeight    =   8595
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
Public spec As Specification
Public console As ConsoleBox

Private Sub cmdBack_Click()
    Unload Me
    GoToMain
End Sub

Private Sub cmdSaveChanges_Click()
    'spec.SaveSpecification
End Sub

Private Sub cmdSubmit_Click()
    spec.Properties.Item(Utils.ConvertToCamelCase(cboSelectProperty.value)) = txtPropertyValue
    console.PrintObject spec
End Sub

Private Sub cmdSearch_Click()
    Dim key As Variant
    Set spec = SpecManager.GetSpec(txtSAPcode.value)
    If spec Is Nothing Then
        MsgBox "Specification not found!", , "Null Spec Exception"
        Exit Sub
    Else
        Set console = Factory.CreateConsoleBox(Me)
        console.PrintObject spec
        PopulateCboSelectProperty
    End If
End Sub

Private Sub PopulateCboSelectProperty()
    Dim key As Variant
    With cboSelectProperty
        For Each key In spec.Properties
          .AddItem Utils.SplitCamelCase(CStr(key))
        Next key
    End With
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearForm Me
End Sub

' Constructor
Private Sub UserForm_Initialize()
    Set spec = New Specification
    Set console = New ConsoleBox
End Sub

' Deconstructor
Private Sub UserForm_Terminate()
    Set spec = Nothing
    Set console = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub


