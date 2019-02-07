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
'@Folder("Interfaces")
' Interface for version control
' of standardized process
' specifications
Public Property Get TableName() As String: End Property
Public Property Get Properties() As Object: End Property
Public Property Get SpecType() As String: End Property
Public Property Get ParentSpec() As ISpec: End Property
' Public Property Get MaterialDescription() As String: End Property
' Public Property Get StyleSpec() As StyleSpecification: End Property
' Public Sub Push()
' End Sub

' Public Sub Commit()
' End Sub

Public Sub JsonToObject(jsonText As String)
End Sub

Public Sub SetDefaultProperties()
End Sub
