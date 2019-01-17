Attribute VB_Name = "Warping"
Option Explicit

Sub GoToMenu()
    
    Application.Visible = False
    UnloadAllForms
    formWarpingMainMenu.Show

End Sub

Function Main(frm As UserForm) As Long
' Entry point for the warping module.
    Const Padding = 25
    Dim key As Variant
    Dim spec As WarpingSpecification
    Dim w As Warp
    Dim Console As ConsoleBox
    Set Console = Factory.CreateConsoleBox(frm)
    ' Retrieve a specification
    With frm
        Set spec = RetrieveWarpingSpecification(.txtSAPcode.Text)
        Set w = Factory.CreateWarp( _
        spec, .txtNumberOfBobbins, .txtPackageWeightlbs, .txtWarpLength)
    End With
    For Each key In spec.PrettyProperties
        Console.PrintLine key, spec.PrettyProperties(key)
    Next key
    ' Variable parameters
    Console.PrintLine "Package Length [yds]", CStr(Round(w.PackageLengthYds, 2))
    Console.PrintLine "Number of Sections [-]", CStr(Round(w.NumberOfSections, 2))
    Console.PrintLine "Residual Length [yds]", CStr(Round(w.ResidualLengthYds, 2))
End Function

Function RetrieveStyleSpecification(Style As Long) As StyleSpecification
' Retrieves a style spec from the database
    Dim SQLstmt As String
    Dim record As DatabaseRecord
    Dim styleSpec As StyleSpecification
    
    Set styleSpec = New StyleSpecification
    Set record = New DatabaseRecord
    SQLstmt = "SELECT * FROM tblStyleSpecs " & _
              "WHERE Style = " & Style & ";"
    Debug.Print SQLstmt
    Set record = ExecuteSQLite3Select(SQLstmt)
    With styleSpec
        .Dtex = record.Fields("Dtex")
        .Style = record.Fields("Style")
        .WeaveType = record.Fields("WeaveType")
        .YarnType = record.Fields("YarnType")
        .Denier = record.Fields("Denier")
        .MeanWarpCount = record.Fields("MeanWarpCount")
        .MinWarpCount = record.Fields("MinWarpCount")
        .MaxWarpCount = record.Fields("MaxWarpCount")
        .MeanFillCount = record.Fields("MeanFillCount")
        .MinFillCount = record.Fields("MinFillCount")
        .MaxFillCount = record.Fields("MaxFillCount")
        .MeanDryWeight = record.Fields("MeanDryWeight")
        .MinDryWeight = record.Fields("MinDryWeight")
        .MaxDryWeight = record.Fields("MaxDryWeight")
        .MeanConditionedWeight = record.Fields("MeanConditionedWeight")
        .MinConditionedWeight = record.Fields("MinConditionedWeight")
        .MaxConditionedWeight = record.Fields("MaxConditionedWeight")
        .YarnFinish = record.Fields("YarnFinish")
        .YarnCode = record.Fields("YarnCode")
        .MoistureRegain = record.Fields("MoistureRegain")
        .Twisting = record.Fields("Twisting")
        .YarnColor = record.Fields("YarnColor")
        .Notes = record.Fields("Notes")
        .YarnMerge = record.Fields("YarnMerge")
    End With

    Set RetrieveStyleSpecification = styleSpec

End Function

Function RetrieveWarpingSpecification(MaterialNumber As String) As WarpingSpecification
' Retrieves a warping spec from the database
    Dim SQLstmt As String
    Dim record As DatabaseRecord
    Dim field As Variant
    Dim key As Variant
    Dim warpSpec As WarpingSpecification
    Dim styleSpec As StyleSpecification
    
    Set warpSpec = New WarpingSpecification
    Set styleSpec = RetrieveStyleSpecification(Mid(MaterialNumber, 6, 3))
    Set warpSpec.styleSpec = styleSpec
    Set record = New DatabaseRecord

    SQLstmt = "SELECT * FROM tblWarpingSpecs " & _
              "WHERE MaterialNumber = """ & MaterialNumber & """;"
    Debug.Print SQLstmt
    Set record = ExecuteSQLite3Select(SQLstmt)

    With warpSpec
        .MaterialNumber = record.Fields("MaterialNumber")
        .MaterialDescription = record.Fields("MaterialDescription")
        .FinalWidthCm = record.Fields("FinalWidthCm")
        .WarpingSpeed = record.Fields("WarpingSpeed")
        .BeamingSpeed = record.Fields("BeamingSpeed")
        .CrossWinding = record.Fields("CrossWinding")
        .DentsPerCm = record.Fields("DentsPerCm")
        .EndsPerDent = record.Fields("EndsPerDent")
        .IsSWrapped = record.Fields("IsSWrapped")
        .NumberOfEnds = record.Fields("NumberOfEnds")
        .BeamWidth = record.Fields("BeamWidth")
        .BeamingTension = record.Fields("BeamingTension")
        .WarpingTension = record.Fields("WarpingTension")
        .K1 = record.Fields("K1")
        .K2 = record.Fields("K2")
        .SetProperties
        .SetPrettyProperties
    End With
    
    Set RetrieveWarpingSpecification = warpSpec

End Function

Public Sub UpdateCurrentSpec()
' Updates a current spec in database with modifcations
    Dim SQLstmt As String
    Dim RetVal As Long
    ' Create SQL statement from object
    SQLstmt = ""
    RetVal = ExecuteSQLite3(SQLstmt)
    If Not RetVal = SQLITE_DONE Then Err.Raise Number:=1
End Sub

Public Sub AddNewSpec(MaterialNumber As String, MaterialDescription As String)
' Add a new spec to the database
    Dim spec As WarpingSpecification

    Set spec = CreateWarpingSpecification(MaterialNumber, MaterialDescription)
    spec.SetDefaultProperties
    spec.SaveSpecification
End Sub

Private Sub MassLoadSpecifications()
' One time use to mass upload specs to the database
    Dim i As Integer
    Dim code As Range
    Dim Description As Range
    Dim ws As Worksheet

    Set ws = Sheets("Materials")
    
    For i = 2 To 80
        Set code = ws.Cells(i, 1)
        Debug.Print code.value
        Set Description = ws.Cells(i, 2)
        AddNewSpec code.value, Description.value
    Next i

End Sub
