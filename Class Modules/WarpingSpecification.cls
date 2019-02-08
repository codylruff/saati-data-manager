VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WarpingSpecification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Classes")
'====================================
'CLASS DESCRIPTION:
' Specification for sectional warper
'====================================
'------------
'INTERFACES
'------------
Implements IPrint
Implements ISpec
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
Public MaterialNumber             As String
Public MaterialDescription        As String
Public FinalWidthCm               As Double
Public EndsPerInch                As Double
Public Dtex                       As Double
Public WarpingSpeed               As Double
Public BeamingSpeed               As Double
Public CrossWinding               As Double
Public DentsPerCm                 As Double
Public EndsPerDent                As Double
Public IsSWrapped                 As Boolean
Public NumberOfEnds               As Double
Public BeamWidth                  As Double
Public BeamingTension             As Double
Public WarpingTension             As Double
Public K1                         As Double
Public K2                         As Double
Public StyleSpec                  As StyleSpecification
Public Style                      As Long
Public WarpDensityStart           As Double
Public WarpDensityMeasured        As Double
Public CompactionFactor           As Double
Public EvenerRoller               As Boolean
Public MeasuringPhase             As Boolean
Public ReedShape                  As Long
Public YarnLeaseFeed              As Double
Public Properties_                As Object
Public IPrint_PrettyProperties_   As Object
Public JsonText_                  As String
'-------------
' COLLECTIONS
'-------------
' IPrint_PrettyProperties
Public Property Get IPrint_PrettyProperties() As Object
    Set IPrint_PrettyProperties = IPrint_PrettyProperties_
End Property
Public Property Set IPrint_PrettyProperties(dict As Object)
    Dim key As Variant
    For Each key In ISpec_Properties.keys
        dict.Add key:=Utility.SplitCamelCase(CStr(key)), Item:=ISpec_Properties.Item(key)
    Next key
    Set IPrint_PrettyProperties_ = dict
End Property
' Properties
Public Property Get ISpec_Properties() As Object
    Set ISpec_Properties = Properties_
End Property
Public Property Set ISpec_Properties(dict As Object)
    dict.Add key:="MaterialNumber", Item:=MaterialNumber
    dict.Add key:="MaterialDescription", Item:=MaterialDescription
    dict.Add key:="FinalWidthCm", Item:=FinalWidthCm
    dict.Add key:="NumberOfEnds", Item:=NumberOfEnds
    dict.Add key:="IsSWrapped", Item:=IsSWrapped
    dict.Add key:="SpringColor", Item:=SpringColor
    dict.Add key:="WarpingSpeed", Item:=WarpingSpeed
    dict.Add key:="BeamingSpeed", Item:=BeamingSpeed
    dict.Add key:="CrossWinding", Item:=CrossWinding
    dict.Add key:="DentsPerCm", Item:=DentsPerCm
    dict.Add key:="EndsPerDent", Item:=EndsPerDent
    dict.Add key:="Style", Item:=Style
    dict.Add key:="BeamWidth", Item:=BeamWidth
    dict.Add key:="YarnSupplier", Item:=YarnSupplier
    dict.Add key:="YarnCode", Item:=YarnCode
    dict.Add key:="K1", Item:=K1
    dict.Add key:="WarpingTension", Item:=WarpingTension
    dict.Add key:="K2", Item:=K2
    dict.Add key:="BeamingTension", Item:=BeamingTension
    dict.Add key:="Dtex", Item:=Dtex
    JsonString_ = JsonConverter.ConvertToJson(dict)
    dict.Add key:="jsonText", Item:=JsonString_
    Set Properties_ = dict
End Property
'----------------
' AUTO-PROPERTIES
'----------------
' Specification Type
Public Property Get ISpec_SpecType() As String: ISpec_SpecType = "warping": End Property
Public Property Get ISpec_MaterialId() As String: ISpec_MaterialId = Me.MaterialNumber: End Property
Public Property Get ISpec_ParentSpec() As ISpec: Set ISpec_ParentSpec = Me.StyleSpec: End Property
Public Property Let ISpec_JsonText(value As String) As String: JsonText_ = value: End Property
Public Property Get ISpec_JsonText() As String: ISpec_JsonText = JsonText_: End Property
'------------
' PROPERTIES
'------------
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
'---------
' METHODS
'---------
Public Sub ISpec_JsonToObject(jsonText As String)
' Convert a json string to a warp spec
    Dim dict: Set dict = JsonConverter.ParseJson(jsonText)
    On Error Resume Next
    With Me
        .MaterialNumber = dict.Item("MaterialNumber")
        .MaterialDescription = dict.Item("MaterialDescription")
        .FinalWidthCm = dict.Item("FinalWidthCm")
        .WarpingSpeed = dict.Item("WarpingSpeed")
        .BeamingSpeed = dict.Item("BeamingSpeed")
        .CrossWinding = dict.Item("CrossWinding")
        .DentsPerCm = dict.Item("DentsPerCm")
        .EndsPerDent = dict.Item("EndsPerDent")
        .IsSWrapped = dict.Item("IsSWrapped")
        .NumberOfEnds = dict.Item("NumberOfEnds")
        .BeamWidth = dict.Item("BeamWidth")
        .BeamingTension = dict.Item("BeamingTension")
        .WarpingTension = dict.Item("WarpingTension")
        .K1 = dict.Item("K1")
        .K2 = dict.Item("K2")
    End With
    Me.SetProperties
End Sub

Public Sub ISpec_ObjectToJson()
' Convert this object to a json string
    
End Sub

Public Sub ISpec_SetDefaultProperties()
' Bases property values on theoretical calcualtions
    Style = StyleSpec.Style
    Dtex = StyleSpec.Dtex
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
    K1 = IIf(YarnSupplier = "Dupont", 0.25, 0.15)
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
    Set ISpec_Properties = CreateObject("Scripting.Dictionary")
    ISpec_JsonText = JsonConverter.ConvertToJson(ISpec_Properties)
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
    Style = StyleSpec.Style
    Dtex = StyleSpec.Dtex
    SetProperties
    ' Create the insert portion of the statement
    INSERTstmt = "INSERT INTO " & tbl & " ("
    ' Create the values portion of the statement
    VALUESstmt = "VALUES ("

    For Each key In ISpec_Properties
        INSERTstmt = INSERTstmt & key & ", "
        VALUESstmt = VALUESstmt & "'" & ISpec_Properties(key) & "', "
    Next key

    INSERTstmt = INSERTstmt & "Time_Stamp) "
    VALUESstmt = VALUESstmt & "'" & Now() & "')"

    ' Create SQL statement from objects
    SQLstmt = INSERTstmt & vbNewLine & VALUESstmt
    'Debug.Print SQLstmt
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt
End Sub
