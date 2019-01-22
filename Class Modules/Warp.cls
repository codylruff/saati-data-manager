VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Warp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Represents a specific warp order.
'====================================
'Factory Module Constructor function:
'Function CreateWarp(Specification As WarpingSpecification, NumberOfBobbins As Integer, _
'                     PackageWeightlbs As Double, WarpLengthYds As Double) As Warp
'----------------
'CONSTANTS
'----------------
Const MinimumPackageLengthyds = 10936
'----------------
'CLASS VARIABLES
'----------------

Private Specification_          As WarpingSpecification
Private NumberOfBobbins_        As Double
Private PackageWeightlbs_       As Double
Private WarpLengthYds_          As Double
'----------------
' COLLECTIONS
'----------------
'----------------
' PROPERTIES
'----------------
' Is Minimum Package Length
Public Property Get IsMinimumPackageLength() As Boolean
    ' Allows for less changes in weaving
    If ResidualLengthYds > MinimumPackageLengthyds Then
        IsMinimumPackageLength = True
    Else
        IsMinimumPackageLength = False
    End If
End Property
' WarpLengthYds
Public Property Let WarpLengthYds(value As Double)
    WarpLengthYds_ = value
End Property
Public Property Get WarpLengthYds() As Double
    WarpLengthYds = WarpLengthYds_
End Property
' Residual Length yards
Public Property Get ResidualLengthYds() As Double
    ResidualLengthYds = PackageLengthYds - WarpLengthYds
End Property
' PackageLengthYds
Public Property Get PackageLengthYds() As Double
    With Me.Specification
        ' 454 g / lb and 0.9144 m / yd
        PackageLengthYds = PackageWeightlbs / (.Dtex / 10000 / 454 * 0.9144)
    End With
End Property
' NumberOfSections
Public Property Get NumberOfSections() As Double
    With Me.Specification
        NumberOfSections = .EndsPerInch * .FinalWidthCm / 2.54 / NumberOfBobbins
    End With
End Property
' PackageWeightlbs
Public Property Let PackageWeightlbs(value As Double)
    PackageWeightlbs_ = value
End Property
Public Property Get PackageWeightlbs() As Double
    PackageWeightlbs = PackageWeightlbs_
End Property
' NumberOfBobbins
Public Property Let NumberOfBobbins(value As Double)
    NumberOfBobbins_ = value
End Property
Public Property Get NumberOfBobbins() As Double
    NumberOfBobbins = NumberOfBobbins_
End Property
' Specification
Public Property Set Specification(value As WarpingSpecification)
    Set Specification_ = value
End Property
Public Property Get Specification() As WarpingSpecification
    Set Specification = Specification_
End Property
'----------------
' METHODS
'----------------
'----------------
' EVENTS
'----------------
' Constructor
Private Sub Class_Initialize()
End Sub
' Deconstructor
Private Sub Class_Terminate()
End Sub
