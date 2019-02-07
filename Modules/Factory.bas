Attribute VB_Name = "Factory"
Option Explicit
'@Folder("Modules")
'----------------------------
' Object factory functions
' serve as pseduo class
' constructors.
'----------------------------
Function CreateDatabaseRecord() As DatabaseRecord
' Creates a database record object
    Dim record: Set record = New DatabaseRecord
    Set CreateDatabaseRecord = record
End Function

Function CreateSQLiteDatabase() As SQLiteDatabase
' Creates a SQLite Database object
    Dim sqlite: Set sqlite = New SQLiteDatabase
    Set CreateSQLiteDatabase = sqlite
End Function

Function CreateConsoleBox() As ConsoleBox
' Creates a console box object
    Dim Console: Set Console = New ConsoleBox
    Set CreateConsoleBox = Console
End Function

Function CreateWarp() As Warp
' Creates a warp object
    Dim w: Set w = New Warp
    Set CreateWarp = w
End Function

Function CreateWarpingSpecification() As WarpingSpecification
' Creates a Warping Specification object
    Dim spec: Set spec = New WarpingSpecification
    Set CreateWarpingSpecification = spec
End Function

Function CreateStyleSpecification() As StyleSpecification
' Creates a style specification object
    Dim Style: Set Style = New StyleSpecification
    Set CreateStyleSpecification = Style
End Function

Function CreateSlitterSpecification() As SlitterSpecification
    Dim spec: Set spec = New SlitterSpecification
    Set CreateSlitterSpecification = spec
End Function

Function CreateUltraSonicSpecification() As UltraSonicSpecification
    Dim spec: Set spec = New UltraSonicSpecification
    Set CreateUltraSonicSpecification = spec
End Function

Function CreateISpec(specType As Long) As ISpec
' Gets and returns a certain type of ISpec object
    Select Case specType
        Case ISPEC_WARPING
            Set CreateISpec = CreateWarpingSpecification
        Case ISPEC_STYLE
            Set CreateISpec = CreateStyleSpecification
        Case ISPEC_SLITTER
            Set CreateISpec = CreateSlitterSpecification
        Case ISPEC_ULTRASONIC
            Set CreateISpec = CreateUltraSonicSpecification
    End Select
End Function

