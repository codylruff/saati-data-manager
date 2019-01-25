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
' specifications such as Weaving
Public Property Get Properties() As Object
End Property
Public Property Get IsDefaultSpec() As Boolean
End Property
Public Sub Push()
End Sub
Public Sub Commit()
End Sub
Private Sub SetProperties()
End Sub
Public Sub SetDefaultProperties()
End Sub
Private Sub SaveSpecification()
End Sub