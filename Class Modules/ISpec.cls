VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Interface for version control
' of standardized process
' specifications
Public Property Get MaterialNumber() As String
End Property
Public Property Get MaterialDescription() As String
End Property
Public Sub Push()
End Sub
Public Sub Commit()
End Sub
