VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WarpingSpecification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Specification for sectional warper
'====================================
'------------
'INTERFACES
'------------
Implements IPrint
'Implements IDatabaseRecord
'------------------------------------
'Factory Module Constructor function:
'------------------------------------
'Function CreateWarpingSpecification(MaterialNumber As String, MaterialDescription As String, _
'                         DentsPerCm As Double,EndsPerDent As Double) As WarpingSpecification
'----------------
'CONSTANTS
'----------------
Const HighDtex = 3000
'----------------
'CLASS VARIABLES
'----------------
Private MaterialNumber_         As String
Private MaterialDescription_    As String
Private FinalWidthCm_           As Double
Private EndsPerInch_            As Double
Private Dtex_                   As Double
Private WarpingSpeed_           As Double
Private BeamingSpeed_           As Double
Private CrossWinding_           As Double
Private DentsPerCm_             As Double
Private EndsPerDent_            As Double
Private IsSWrapped_             As Boolean
Private NumberOfEnds_           As Double
Private BeamWidth_              As Double
Private BeamingTension_         As Double
Private WarpingTension_         As Double
Private K1_                     As Double
Private K2_                     As Double
Private StyleSpec_              As StyleSpecification
'------------
' COLLECTIONS
'------------
Private Properties_                    As Object
Private IPrint_PrettyProperties_       As Object
'------------
' PROPERTIES
'------------
Private WarpDensityStart_ As Double
Private WarpDensityMeasured_ As Double
Private CompactionFactor_ As Double
Private EvenerRoller_ As Boolean
Private MeasuringPhase_ As Boolean
Private ReedShape_ As Long
Private YarnLeaseFeed_ As Double
' YarnLeaseFeed
Public Property Let YarnLeaseFeed(value As Double)
    YarnLeaseFeed_ = value
End Property
Public Property Get YarnLeaseFeed() As Double
    YarnLeaseFeed = YarnLeaseFeed_
End Property
' ReedShape
Public Property Let ReedShape(value As Long)
    ReedShape_ = value
End Property
Public Property Get ReedShape() As Long
    ReedShape = ReedShape_
End Property
' MeasuringPhase
Public Property Let MeasuringPhase(value As Boolean)
    MeasuringPhase_ = value
End Property
Public Property Get MeasuringPhase() As Boolean
    MeasuringPhase = MeasuringPhase_
End Property
' EvenerRoller
Public Property Let EvenerRoller(value As Boolean)
    EvenerRoller_ = value
End Property
Public Property Get EvenerRoller() As Boolean
    EvenerRoller = EvenerRoller_
End Property
' CompactionFactor
Public Property Let CompactionFactor(value As Double)
    CompactionFactor_ = value
End Property
Public Property Get CompactionFactor() As Double
    CompactionFactor = CompactionFactor_
End Property
' WarpDensityMeasured
Public Property Let WarpDensityMeasured(value As Double)
    WarpDensityMeasured_ = value
End Property
Public Property Get WarpDensityMeasured() As Double
    WarpDensityMeasured = WarpDensityMeasured_
End Property
' WarpDensityStart
Public Property Let WarpDensityStart(value As Double)
    WarpDensityStart_ = value
End Property
Public Property Get WarpDensityStart() As Double
    WarpDensityStart = WarpDensityStart_
End Property
' YarnCode
Public Property Get YarnCode() As String
    YarnCode = styleSpec.YarnCode
End Property
' K2
Public Property Let K2(value As Double)
    K2_ = value
End Property
Public Property Get K2() As Double
    K2 = K2_
End Property
' K1
Public Property Let K1(value As Double)
    K1_ = value
End Property
Public Property Get K1() As Double
    K1 = K1_
End Property
' WarpingTension
Public Property Let WarpingTension(value As Double)
    WarpingTension_ = value
End Property
Public Property Get WarpingTension() As Double
    WarpingTension = WarpingTension_
End Property
' BeamingTension
Public Property Let BeamingTension(value As Double)
    BeamingTension_ = value
End Property
Public Property Get BeamingTension() As Double
    BeamingTension = BeamingTension_
End Property
' BeamWidth
Public Property Let BeamWidth(value As Double)
    BeamWidth_ = value
End Property
Public Property Get BeamWidth() As Double
    BeamWidth = BeamWidth_
End Property
' NumberOfEnds
Public Property Let NumberOfEnds(value As Double)
    NumberOfEnds_ = value
End Property
Public Property Get NumberOfEnds() As Double
    NumberOfEnds = NumberOfEnds_
End Property
' StyleSpec
Public Property Set styleSpec(value As StyleSpecification)
    Set StyleSpec_ = value
End Property
Public Property Get styleSpec() As StyleSpecification
    Set styleSpec = StyleSpec_
End Property
' PrettyProperties
Public Property Get IPrint_PrettyProperties() As Object
    Set IPrint_PrettyProperties = IPrint_PrettyProperties_
End Property
Public Property Set IPrint_PrettyProperties(dict As Object)
    dict.Add key:="Style", item:=Style
    dict.Add key:="Material Number", item:=MaterialNumber
    dict.Add key:="Material Description", item:=MaterialDescription
    dict.Add key:="Yarn Type", item:=YarnSupplier
    dict.Add key:="Yarn Code", item:=YarnCode
    dict.Add key:="Final Width [cm]", item:=FinalWidthCm
    dict.Add key:="Number Of Ends", item:=NumberOfEnds
    If IsSWrapped = True Then
        dict.Add key:="S Wrap On/Off", item:="ON"
    Else
        dict.Add key:="S Wrap On/Off", item:="OFF"
    End If
    dict.Add key:="Spring Color", item:=SpringColor
    dict.Add key:="Warping Speed [m/min]", item:=WarpingSpeed
    dict.Add key:="Beaming Speed [m/min]", item:=BeamingSpeed
    dict.Add key:="WarpingTension", item:=WarpingTension
    dict.Add key:="Beaming Tension", item:=BeamingTension
    dict.Add key:="Cross Winding", item:=CrossWinding
    dict.Add key:="Dent/cm", item:=DentsPerCm
    dict.Add key:="End/Dent", item:=EndsPerDent
    dict.Add key:="Beam Width", item:=BeamWidth
    Set IPrint_PrettyProperties_ = dict
End Property
' Properties
Public Property Get Properties() As Object
    Set Properties = Properties_
End Property
Public Property Set Properties(dict As Object)
    dict.Add key:="MaterialNumber", item:=MaterialNumber
    dict.Add key:="MaterialDescription", item:=MaterialDescription
    dict.Add key:="FinalWidthCm", item:=FinalWidthCm
    dict.Add key:="NumberOfEnds", item:=NumberOfEnds
    dict.Add key:="IsSWrapped", item:=IsSWrapped
    dict.Add key:="SpringColor", item:=SpringColor
    dict.Add key:="WarpingSpeed", item:=WarpingSpeed
    dict.Add key:="BeamingSpeed", item:=BeamingSpeed
    dict.Add key:="CrossWinding", item:=CrossWinding
    dict.Add key:="DentsPerCm", item:=DentsPerCm
    dict.Add key:="EndsPerDent", item:=EndsPerDent
    dict.Add key:="Style", item:=Style
    dict.Add key:="BeamWidth", item:=BeamWidth
    dict.Add key:="YarnSupplier", item:=YarnSupplier
    dict.Add key:="YarnCode", item:=YarnCode
    dict.Add key:="K1", item:=K1
    dict.Add key:="WarpingTension", item:=WarpingTension
    dict.Add key:="K2", item:=K2
    dict.Add key:="BeamingTension", item:=BeamingTension
    dict.Add key:="Dtex", item:=Dtex
    Set Properties_ = dict
End Property
' Style
Public Property Get Style() As Long
    Style = styleSpec.Style
End Property
' EndsPerDent
Public Property Let EndsPerDent(value As Double)
    EndsPerDent_ = value
End Property
Public Property Get EndsPerDent() As Double
    EndsPerDent = EndsPerDent_
End Property
' DentsPerCm
Public Property Let DentsPerCm(value As Double)
    DentsPerCm_ = value
End Property
Public Property Get DentsPerCm() As Double
    DentsPerCm = DentsPerCm_
End Property
' CrossWinding
Public Property Let CrossWinding(value As Double)
    CrossWinding_ = value
End Property
Public Property Get CrossWinding() As Double
    CrossWinding = CrossWinding_
End Property
' BeamingSpeed
Public Property Let BeamingSpeed(value As Double)
    BeamingSpeed_ = value
End Property
Public Property Get BeamingSpeed() As Double
    BeamingSpeed = BeamingSpeed_
End Property
' Yarn Supplier
Public Property Get YarnSupplier() As String
    Select Case Mid(MaterialNumber, 2, 2)
        Case "KE"
            YarnSupplier = "Dupont"
        Case "TW"
            YarnSupplier = "Twaron"
        Case "NM"
            YarnSupplier = "Invista"
        Case "DY"
            YarnSupplier = "DSM"
        Case "HY"
            YarnSupplier = "Hyosung"
        Case "IN"
            YarnSupplier = "Innegra"
        Case "SP"
            YarnSupplier = "Spectra"
        Case Else
            YarnSupplier = "Unknown"
    End Select
End Property
' WarpingSpeed
Public Property Let WarpingSpeed(value As Double)
    WarpingSpeed_ = value
End Property
Public Property Get WarpingSpeed() As Double
    WarpingSpeed = WarpingSpeed_
End Property
' Spring Color
Public Property Get SpringColor() As String
    Select Case True
        Case Dtex <= 440
            SpringColor = "Blue"
        Case Dtex >= 660 And Dtex < 1100
            SpringColor = "Black"
        Case Else
            SpringColor = "Yellow"
    End Select
End Property
' Material Number
Public Property Let MaterialNumber(value As String)
    MaterialNumber_ = value
End Property
Public Property Get MaterialNumber() As String
    MaterialNumber = MaterialNumber_
End Property
' Material Description
Public Property Let MaterialDescription(value As String)
    MaterialDescription_ = value
End Property
Public Property Get MaterialDescription() As String
    MaterialDescription = MaterialDescription_
End Property
' Final Width in cm
Public Property Let FinalWidthCm(value As Double)
    FinalWidthCm_ = value
End Property
Public Property Get FinalWidthCm() As Double
    FinalWidthCm = FinalWidthCm_
End Property
' Ends Per Inch
Public Property Get EndsPerInch() As Double
    EndsPerInch = styleSpec.MeanWarpCount
End Property
' Linear Density in Dtex or g / 10000 m
Public Property Get Dtex() As Double
    Dtex = styleSpec.Dtex
End Property
' S Wrapping
Public Property Let IsSWrapped(value As Boolean)
    IsSWrapped_ = value
End Property
Public Property Get IsSWrapped() As Boolean
    IsSWrapped = IsSWrapped_
End Property
'---------
' METHODS
'---------
Public Sub SetDefaultProperties()
' Bases property values on theoretical calcualtions
    Dim tmp As Double
    If Right(MaterialDescription, 2) = "CM" Then
        FinalWidthCm = CDbl(Left(Right(MaterialDescription, 5), 3))
    ElseIf Right(MaterialDescription, 1) = """" Then
        FinalWidthCm = Round(CDbl(Left(Right(MaterialDescription, 3), 2)) / 2.54, 2)
    Else
        Exit Sub
    End If
    tmp = Round(FinalWidthCm * EndsPerInch / 2.54, 0)
    If tmp Mod 2 = 0 Then
        NumberOfEnds = tmp
    Else
        NumberOfEnds = tmp + 1
    End If
    If Dtex >= 3000 Then IsSWrapped = True
    If Dtex <= 3000 Then
        BeamWidth = ((FinalWidthCm * 10#) - 3#) / 10#
    Else
        BeamWidth = ((FinalWidthCm * 10#) - 8#) / 10#
    End If
    If YarnSupplier = "Dupont" Then
        K1 = 0.25
    Else
        K1 = 0.15
    End If
    K2 = K1 + 1#
    WarpingTension = Round(Dtex * K1, 0)
    BeamingTension = Round(NumberOfEnds * WarpingTension * K2 / 100#, 0)
    WarpingSpeed = 300 ' meters / min
    If Dtex >= HighDtex Then
        BeamingSpeed = 80 ' meters / min
        CrossWinding = 10
    Else
        BeamingSpeed = 120 ' meters / min
        CrossWinding = 5
    End If
End Sub
Public Sub SetProperties()
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
Public Sub IPrint_SetPrettyProperties()
    Set IPrint_PrettyProperties = CreateObject("Scripting.Dictionary")
End Sub
Public Sub SaveSpecification()
' Adds a new specification to the database
    Dim key As Variant
    Dim SQLstmt As String
    Dim INSERTstmt, VALUESstmt As String
    Dim tbl As String: tbl = "tblWarpingSpecs"
    ' Set object properties
    SetProperties
    ' Create the insert portion of the statement
    INSERTstmt = "INSERT INTO " & tbl & " ("
    ' Create the values portion of the statement
    VALUESstmt = "VALUES ("

    For Each key In Properties
        INSERTstmt = INSERTstmt & key & ", "
        VALUESstmt = VALUESstmt & "'" & Properties(key) & "', "
    Next key

    INSERTstmt = INSERTstmt & "Time_Stamp) "
    VALUESstmt = VALUESstmt & "'" & Now() & "')"

    ' Create SQL statement from object
    SQLstmt = INSERTstmt & vbNewLine & VALUESstmt
    Debug.Print SQLstmt
    RetVal = ExecuteSQLite3(SQLstmt)
    If Not RetVal = SQLITE_DONE Then Err.Raise Number:=1
End Sub
