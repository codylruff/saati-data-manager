VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
'Interface for json converter library
'=====================================
'----------------
'CLASS VARIABLES
'----------------
Private JsonString_ As String
'------------
' COLLECTIONS
'------------
'------------
' PROPERTIES
'------------
' JsonString
Public Property Let JsonString(value As String)
    'JsonString_ = value
End Property
Public Property Get JsonString() As String
    'JsonString = JsonString_
End Property
'----------
' METHODS
'----------
Public Function ConvertToJson(Optional Whitespace As Integer = 2, Optional Indentation As Long = 1) As String
    'JsonString = JsonConverter.ConvertToJson( _
    '               JsonValue:= JsonProperties, _
    '               Whitespace:=Whitespace, _
    '               Indentation:=Indentation)
    'ConvertToJson = JsonString
End Function
Public Sub MapToObject(json_string As String)
    ' JsonString = json_string
    ' Set ISpec_Properties = JsonConverter.ParseJson(json_string)
End Sub