Attribute VB_Name = "Factory"
Option Explicit

Function CreateConsoleBox(frm As UserForm) As ConsoleBox
    Dim Console As ConsoleBox
    Set Console = New ConsoleBox
    Set Console.FormId = frm
    Set CreateConsoleBox = Console
End Function
Function CreateWarp(Specification As WarpingSpecification, NumberOfBobbins As Integer, _
                     PackageWeightlbs As Double, WarpLength As Double) As Warp
    Dim w As Warp
    Set w = New Warp
    ' Create object
    Set w.Specification = Specification
    With w
        .NumberOfBobbins = NumberOfBobbins
        .PackageWeightlbs = PackageWeightlbs
        .WarpLengthYds = WarpLength
    End With
    ' Return object
    Set CreateWarp = w
End Function

Function CreateWarpingSpecification(MaterialNumber As String, MaterialDescription As String, _
                     Optional DentsPerCm As Double = 0, Optional EndsPerDent As Double = 0) As WarpingSpecification

    Dim spec As WarpingSpecification
    Dim Sty As Long
    Dim styleSpec As StyleSpecification
    Set spec = New WarpingSpecification
    Sty = CLng(Mid(MaterialNumber, 6, 3))
    Debug.Print Sty
    Set styleSpec = RetrieveStyleSpecification(Sty)
    
    ' Create Object
    With spec
        .MaterialNumber = MaterialNumber
        .MaterialDescription = MaterialDescription
        .DentsPerCm = DentsPerCm
        .EndsPerDent = EndsPerDent
    End With
    Set spec.styleSpec = styleSpec
    ' Return object
    Set CreateWarpingSpecification = spec

End Function

