Attribute VB_Name = "Factory"
Option Explicit
'----------------------------
' Object factory functions
' serve as pseduo class
' constructors.
'----------------------------
Function CreateSQLiteDb() As SQLiteDatabase
' Creates a SQLite Database object
    Dim sqlite: Set sqlite = New SQLiteDatabase
    Set CreateSQLiteDb = sqlite
End Function

Function CreateConsoleBox(frm As UserForm) As ConsoleBox
' Creates a console box object
    Dim Console: Set Console = New ConsoleBox
    Set Console.FormId = frm
    Set CreateConsoleBox = Console
End Function

Function CreateWarp(Specification As WarpingSpecification, NumberOfBobbins As Integer, _
                     PackageWeightlbs As Double, WarpLength As Double) As Warp
' Creates a warp object
    Dim w: Set w = New Warp
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
' Creates a Warping Specification object
    Dim spec: Set spec = New WarpingSpecification
    Dim Sty As Long
    Dim styleSpec As StyleSpecification
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

