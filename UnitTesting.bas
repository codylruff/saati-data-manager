Attribute VB_Name = "UnitTesting"
Option Explicit

Public Function WarpingSpecificationTest()
    Dim spec As WarpingSpecification

    Set spec = New WarpingSpecification

    With spec
        .MaterialNumber = "OKE101011AP00CN"
        .MaterialDescription = "STY 101 K29 3300D 1F279 131CM"
        .FinalWidthCm = 103.5
        .WarpingSpeed = 300
        .BeamingSpeed = 80
        .CrossWinding = 10
        .DentsPerCm = 0
        .EndsPerDent = 0
        
    End With
    spec.SaveSpecification
End Function

Public Sub select_Test()
    Dim record As DatabaseRecord
    Dim field As Variant
    
    Set record = ExecuteSQLite3Select("SELECT * FROM tblWarpingSpecs WHERE spec_id = 1")
    
    For Each field In record.Fields
        Debug.Print field & " " & record.Fields(field)
    Next field

End Sub

Public Sub Retrieve_Test()
    Const Padding = 24

    Dim key As Variant
    Dim spec As WarpingSpecification
    
    Set spec = RetrieveWarpingSpecification("OKE101011AP00CN")
    'Set w = Factory.CreateWarp(spec, NumberOfBobbins, PackageWeightlbs, WarpLength)
    
    For Each key In spec.Properties
        Debug.Print Left$(key & ":" & Space(Padding), _
                    Padding) & spec.Properties(key)
    Next key
    

End Sub
