Attribute VB_Name = "Warping"
Option Explicit
'@Folder("Modules")

Sub GoToMenu()
    
    Application.Visible = False
    UnloadAllForms
    formWarpingMainMenu.Show

End Sub

Function PrintWarpingSpecification(frm As UserForm) As WarpingSpecification
    Dim spec As WarpingSpecification
    Dim Console As ConsoleBox
    Set Console = Factory.CreateConsoleBox(frm)
    ' Retrieve a specification
    Set spec = RetrieveWarpingSpecification(frm.txtSAPcode.Text)
    ' Print object to console
    Console.PrintObject spec
    Set PrintWarpingSpecification = spec
End Function

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
    ' Print object to console
    Console.PrintObject spec
    ' Variable parameters
    Console.PrintLine "Package Length [yds]", CStr(Round(w.PackageLengthYds, 2))
    Console.PrintLine "Number of Sections [-]", CStr(Round(w.NumberOfSections, 2))
    Console.PrintLine "Residual Length [yds]", CStr(Round(w.ResidualLengthYds, 2))
End Function

Function RetrieveStyleSpecification(Style As Long) As StyleSpecification
' Retrieves a style spec from the database
    Dim SQLstmt As String
    Dim record As DatabaseRecord
    Dim StyleSpec As StyleSpecification
    
    Set StyleSpec = New StyleSpecification
    Set record = New DatabaseRecord
    SQLstmt = "SELECT * FROM tblStyleSpecs " & _
              "WHERE Style = " & Style
    Debug.Print SQLstmt
    Set record = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, _
    "W:\App Development\Data Manager\Protection_Quality_Control.db3", SQLstmt)
    record.SetDictionary
    With StyleSpec
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

    Set RetrieveStyleSpecification = StyleSpec

End Function

Function RetrieveWarpingSpecification(MaterialNumber As String) As WarpingSpecification
' Retrieves a warping spec from the database
    Dim SQLstmt As String
    Dim record As DatabaseRecord
    Dim field As Variant
    Dim key As Variant
    Dim warpSpec As WarpingSpecification
    Dim StyleSpec As StyleSpecification
    
    Set warpSpec = New WarpingSpecification
    Set StyleSpec = RetrieveStyleSpecification(Mid(MaterialNumber, 6, 3))
    Set warpSpec.StyleSpec = StyleSpec
    Set record = New DatabaseRecord

    SQLstmt = "SELECT * FROM tblWarpingSpecs " & _
              "WHERE MaterialNumber = """ & MaterialNumber & """"
    Debug.Print SQLstmt
    Set record = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt)
    record.SetDictionary
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
        .IPrint_SetPrettyProperties
    End With
    
    Set RetrieveWarpingSpecification = warpSpec

End Function

