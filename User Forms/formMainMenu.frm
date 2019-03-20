VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMainMenu 
   Caption         =   "Data Manager Main Menu"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6300
   OleObjectBlob   =   "formMainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Option Explicit

Private Sub cmdConfig_Click()
    SpecManager.StopSpecManager
    Logger.SaveLog
    Unload Me
    ConfigControl
End Sub

Private Sub cmdCreateTemplate_Click()
    SpecManager.StopSpecManager
    SpecManager.StartSpecManager
    Unload Me
    On Error Resume Next
    formCreateGeneric.Show
End Sub

Private Sub cmdExit_Click()
    ExitApp
End Sub

Private Sub cmdWarping_Click()
    SpecManager.StopSpecManager
    SpecManager.StartSpecManager
    Unload Me
    On Error Resume Next
    formWarpingSearch.Show
End Sub

Private Sub CommandButton3_Click()
    SpecManager.StopSpecManager
    SpecManager.StartSpecManager
    Unload Me
    formSpecConfig.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
