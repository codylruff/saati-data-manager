VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formWarpingSearch 
   Caption         =   "Specification Search"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390
   OleObjectBlob   =   "formWarpingSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formWarpingSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit

Private Sub cmdClear_Click()
'Clears the form
    ClearForm Me
End Sub

Private Sub cmdOptions_Click()
    Unload Me
    GoToMain
End Sub

Private Sub cmdSearch_Click()
    If SpecManager.ExecuteSearch(txtSAPcode.value) <> SM_SEARCH_FAILURE Then
        MsgBox "Specification not found!", , "Null Spec Exception"
        Exit Sub
    End If
    SpecManager.PrintSpecification Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub


