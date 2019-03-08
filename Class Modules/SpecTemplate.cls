VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------
'CLASS VARIABLES
'----------------
Private Properties_     As Variant
Private SpecType_       As String
Private Revision_       As String
'----------------
' AUTO-PROPERTIES
'----------------
Public Property Let Revision(value As String): Revision_ = value: End Property
Public Property Get Revision() As String: Revision = Revision_: End Property
Public Property Let SpecType(value As String): SpecType_ = value: End Property
Public Property Get SpecType() As String: SpecType = SpecType_: End Property
Public Property Get Properties() As Variant: Properties = Properties_: End Property
'--METHODS--'
Public Sub AddProperty(property As String)
' Adds a property to the list of properties
    ExpandArray(Me.Properties, 2, 1, property)
End Sub

Public Function PropertiesJson() As String
    PropertiesJson = JsonConverter.ConvertToJson(Me.Properties)
End Function