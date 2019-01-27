VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SlittingSpecification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================
' CLASS DESCRIPTION:
' Current only valid for Heat Slitting
' This represents both the slitting
' spec and spec history for a
' particular code and width.
'======================================
'------------
'INTERFACES
'------------
Implements IPrint

Private MaterialNumber_         As String
Private MaterialDescription_    As String
Private SpecPath_               As String
Private MeshThickness_          As Double
Private AirPressure_            As Double
Private IsPowerOn_              As Boolean
Private CoreDiameter_           As Double
Private CoreMaterialType_       As String
Private BladeSize_              As String
Private UnwindingTension_       As Double
Private WindingTension_         As Double
Private BladeTemperature_       As Double
'------------
' COLLECTIONS
'------------
Private Properties_             As Object
Private PrettyProperties_       As Object
Private PartHistory_            As Collection
' Properties
Public Property Get Properties() As Object
    Set Properties = Properties_
End Property
Public Property Set Properties(dict As Object)
    dict.Add key:="MaterialNumber", item:=MaterialNumber
    dict.Add key:="MaterialDescription", item:=MaterialDescription
    Set Properties_ = dict
End Property
' PrettyProperties
Public Property Get IPrint_PrettyProperties() As Object
    Set IPrint_PrettyProperties = IPrint_PrettyProperties_
End Property
Public Property Set IPrint_PrettyProperties(dict As Object)
    dict.Add key:="Material Number", item:=MaterialNumber
    dict.Add key:="Material Description", item:=MaterialDescription
    If IsPowerOn = True Then
        dict.Add key:="Power On/Off", item:="ON"
    Else
        dict.Add key:="Power On/Off", item:="OFF"
    End If
    Set IPrint_PrettyProperties_ = dict
End Property
'------------
' PROPERTIES
'------------
' Material Number
Public Property Let MaterialNumber(value As String): MaterialNumber_ = value: End Property
Public Property Get MaterialNumber() As String: MaterialNumber = MaterialNumber_: End Property
' Material Description
Public Property Let MaterialDescription(value As String): MaterialDescription_ = value: End Property
Public Property Get MaterialDescription() As String: MaterialDescription = MaterialDescription_: End Property
' BladeTemperature
Public Property Let BladeTemperature(value As Double): BladeTemperature_ = value: End Property
Public Property Get BladeTemperature() As Double: BladeTemperature = BladeTemperature_: End Property
' WindingTension
Public Property Let WindingTension(value As Double): WindingTension_ = value: End Property
Public Property Get WindingTension() As Double: WindingTension = WindingTension_: End Property
' UnwindingTension
Public Property Let UnwindingTension(value As Double): UnwindingTension_ = value: End Property
Public Property Get UnwindingTension() As Double: UnwindingTension = UnwindingTension_: End Property
' BladeSize
Public Property Let BladeSize(value As String): BladeSize_ = value: End Property
Public Property Get BladeSize() As String: BladeSize = BladeSize_: End Property
' CoreMaterialType
Public Property Let CoreMaterialType(value As String): CoreMaterialType_ = value: End Property
Public Property Get CoreMaterialType() As String: CoreMaterialType = CoreMaterialType_: End Property
' CoreDiameter
Public Property Let CoreDiameter(value As Double): CoreDiameter_ = value: End Property
Public Property Get CoreDiameter() As Double: CoreDiameter = CoreDiameter_: End Property
' IsPowerOn
Public Property Let IsPowerOn(value As Boolean): IsPowerOn_ = value: End Property
Public Property Get IsPowerOn() As Boolean: IsPowerOn = IsPowerOn_: End Property
' AirPressure
Public Property Let AirPressure(value As Double): AirPressure_ = value: End Property
Public Property Get AirPressure() As Double: AirPressure = AirPressure_: End Property
' MeshThickness
Public Property Let MeshThickness(value As Double): MeshThickness_ = value: End Property
Public Property Get MeshThickness() As Double: MeshThickness = MeshThickness_: End Property
' SpecPath
Public Property Let SpecPath(value As String): SpecPath_ = value: End Property
Public Property Get SpecPath() As String: SpecPath = SpecPath_: End Property
'----------
' METHODS
'----------
Public Sub SetProperties()
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
Public Sub IPrint_SetPrettyProperties()
    Set IPrint_PrettyProperties = CreateObject("Scripting.Dictionary")
End Sub
