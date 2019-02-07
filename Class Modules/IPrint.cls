VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Interfaces")
'====================================
'CLASS DESCRIPTION:
' Interface for Console Box printing
'====================================
Private PrettyProperties_       As Object

Public Property Get PrettyProperties() As Object
End Property

Public Property Set PrettyProperties(dict As Object)
    ' Example: Where items are the properties
    '          and the keys are print labels.
    ' dict.Add key:="Dent/cm", item:=DentsPerCm
    ' dict.Add key:="End/Dent", item:=EndsPerDent
    ' dict.Add key:="Beam Width", item:=BeamWidth
    ' Set PrettyProperties_ = dict
End Property

Public Sub SetPrettyProperties()
End Sub
