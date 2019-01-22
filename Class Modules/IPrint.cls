VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Interface for Console Box printing
'====================================
'------------
' COLLECTIONS
'------------
Private PrettyProperties_       As Object
'------------
' PROPERTIES
'------------
' Properties
Public Property Get PrettyProperties() As Object
    ' Set PrettyProperties = PrettyProperties_
End Property
Public Property Set PrettyProperties(dict As Object)
    ' Example: Where items are the properties
    '          and the keys are print labels.
    ' dict.Add key:="Dent/cm", item:=DentsPerCm
    ' dict.Add key:="End/Dent", item:=EndsPerDent
    ' dict.Add key:="Beam Width", item:=BeamWidth
    ' Set PrettyProperties_ = dict
End Property
'----------
' METHODS
'----------
Public Sub SetPrettyProperties()
End Sub
