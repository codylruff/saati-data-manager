VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SlittingSpecification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
'CLASS DESCRIPTION:
' Current only valid for Heat Slitting
' This represents both the slitting
' spec and spec history for a
' particular code and width.
'=====================================
'Factory Module Constructor function:
'Function CreateSlittingSpecification() As SlittingSpecification
'----------------
'CLASS VARIABLES
'----------------
Private SpecPath_ As String
'------------
' COLLECTIONS
'------------
Private PartHistory_ As Collection
'------------
' PROPERTIES
'------------
Private MeshThickness_ As Double
Private AirPressure_ As Double
Private IsPowerOn_ As Boolean
Private CoreDiameter_ As Double
Private CoreMaterialType_ As String
Private BladeSize_ As String
Private UnwindingTension_ As Double
Private WindingTension_ As Double
Private BladeTemperature_ As Double
' BladeTemperature
Public Property Let BladeTemperature(value As Double)
  BladeTemperature_ = value
End Property
Public Property Get BladeTemperature() As Double
  BladeTemperature = BladeTemperature_
End Property
' WindingTension
Public Property Let WindingTension(value As Double)
  WindingTension_ = value
End Property
Public Property Get WindingTension() As Double
  WindingTension = WindingTension_
End Property
' UnwindingTension
Public Property Let UnwindingTension(value As Double)
  UnwindingTension_ = value
End Property
Public Property Get UnwindingTension() As Double
  UnwindingTension = UnwindingTension_
End Property
' BladeSize
Public Property Let BladeSize(value As String)
  BladeSize_ = value
End Property
Public Property Get BladeSize() As String
  BladeSize = BladeSize_
End Property
' CoreMaterialType
Public Property Let CoreMaterialType(value As String)
  CoreMaterialType_ = value
End Property
Public Property Get CoreMaterialType() As String
  CoreMaterialType = CoreMaterialType_
End Property
' CoreDiameter
Public Property Let CoreDiameter(value As Double)
  CoreDiameter_ = value
End Property
Public Property Get CoreDiameter() As Double
  CoreDiameter = CoreDiameter_
End Property
' IsPowerOn
Public Property Let IsPowerOn(value As Boolean)
  IsPowerOn_ = value
End Property
Public Property Get IsPowerOn() As Boolean
  IsPowerOn = IsPowerOn_
End Property
' AirPressure
Public Property Let AirPressure(value As Double)
  AirPressure_ = value
End Property
Public Property Get AirPressure() As Double
  AirPressure = AirPressure_
End Property
' MeshThickness
Public Property Let MeshThickness(value As Double)
  MeshThickness_ = value
End Property
Public Property Get MeshThickness() As Double
  MeshThickness = MeshThickness_
End Property
' SpecPath
Public Property Let SpecPath(value As String)
  SpecPath_ = value
End Property
Public Property Get SpecPath() As String
  SpecPath = SpecPath_
End Property
'----------
' METHODS
'----------
'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
End Sub
' Deconstructor
Private Sub Class_Terminate()
End Sub
