VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Specification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Properties dictionary
Public Properties As Dictionary
' FIELDS and PROPERTIES---
Private SpecType_   As String
Private Revision_   As String
Private MaterialId_ As String
Private IsStandard_ As Boolean
Private IsLatest_   As Boolean
Public Property Let IsStandard(value As Boolean): IsStandard_ = value: End Property
Public Property Get IsStandard() As Boolean: IsStandard = IIf(IsStandard_ = Null, False, IsStandard_): End Property
Public Property Let IsLatest(value As Boolean): IsLatest_ = value: End Property
Public Property Get IsLatest() As Boolean: IsLatest = IIf(IsLatest_ = Null, False, IsLatest_): End Property
Public Property Let Revision(value As String): Revision_ = value: End Property
Public Property Get Revision() As String: Revision = Revision_: End Property
Public Property Let SpecType(value As String): SpecType_ = value: End Property
Public Property Get SpecType() As String: SpecType = SpecType_: End Property
Public Property Get MaterialId() As String: MaterialId = MaterialId_: End Property
Public Property Let MaterialId(value As String): MaterialId_ = value: End Property
' METHODS---
Public Sub JsonToObject(jsonText As String)
' Map json to the spec
    Dim key As Variant
    Set Properties = JsonConverter.ParseJson(jsonText)
    MaterialId = Properties.Item("MaterialId")
    SpecType = Properties.Item("SpecType")
    Revision = Properties.Item("Revision")
End Sub
Public Function ObjectToJson() As String
' Store spec Properties dictionary as json
    ObjectToJson = JsonConverter.ConvertToJson(Properties)
End Function

' Constructor
Private Sub Class_Initialize()
    Set Properties = New Dictionary
End Sub
' Deconstructor
Private Sub Class_Terminate()
    Set Properties = Nothing
End Sub
