VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Fields As Dictionary

Public Sub AddValue(name As String, value As Variant)
    Fields(name) = value
End Sub

Public Sub AddField(name As String)
    Fields.Add key:=name, item:=""
End Sub

'---------
' Events
'---------
' Constructor
Private Sub Class_Initialize()
    Set Me.Fields = New Dictionary
End Sub
' Deconstructor
Private Sub Class_Terminate()
End Sub
