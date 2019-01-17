VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMainMenu 
   Caption         =   "Data Manager Main Menu"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3465
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
    
    Unload Me
    ConfigControl
    
End Sub

Private Sub cmdExit_Click()

    ExitApp

End Sub

Private Sub cmdWarping_Click()

    Unload Me
    Warping.GoToMenu

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        Cancel = True
    End If

End Sub
