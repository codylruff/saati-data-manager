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

Public Properties As Dictionary
Public PrettyProperties As Dictionary ' Properties in normal case instead of camel case.

' Interface for version control
' of standardized process
' specifications
Public Property Get MaterialId() As String: End Property ' sap code, style # etc...
Public Property Get Revision() As String: End Property
Public Sub JsonToObject(jsonText As String)
' Map json to the spec
    Dim key As Variant
    Set Properties = JsonConverter.ParseJson(jsonText)
    For Each key In Properties
        PrettyProperties.Add Key:= key, Item:= Properties.Item(key)
    Next key
End Sub 
Public Function ObjectToJson() As String
' Store spec Properties dictionary as json
    ObjectToJson = JsonConverter.ConvertToJson(Properties)
End Sub 
' Constructor
Private Sub Class_Initialize()
    Set Properties = New Dictionary
    Set PrettyProperties = New Dictionary
End Sub
' Deconstructor
Private Sub Class_Terminate()
    Set Properties = Nothing
    Set PrettyProperties = Nothing
End Sub