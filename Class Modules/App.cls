VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public server As Object
Public standard As Specification
Public current_spec As Specification
Public specs As Object
Public console As ConsoleBox

Private Sub Class_Initialize()
    Set server = CreateObject("DM_LIB.DmComServer")
    Set current_spec = New Specification
    Set standard = New Specification
    Set specs = CreateObject("Scripting.Dictionary")
    Set console = New ConsoleBox
End Sub

Private Sub Class_Terminate()
    Set server = Nothing
    Set current_spec = Nothing
    Set standard = Nothing
    Set specs = Nothing
    Set console = Nothing
End Sub


