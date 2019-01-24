VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleSpecification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================
'CLASS DESCRIPTION:
' Represents the spec for a specific style
'=========================================
Implements ISpec
'----------------
'CLASS VARIABLES
'----------------
'----------------
'CLASS VARIABLES
'----------------
Private Dtex_                       As Double
Private Style_                      As Long
Private WeaveType_                  As String
Private YarnType_                   As String
Private Denier_                     As Double
Private MeanWarpCount_              As Double
Private MinWarpCount_               As Double
Private MaxWarpCount_               As Double
Private MeanFillCount_              As Double
Private MinFillCount_               As Double
Private MaxFillCount_               As Double
Private MeanDryWeight_              As Double
Private MinDryWeight_               As Double
Private MaxDryWeight_               As Double
Private MeanConditionedWeight_      As Double
Private MinConditionedWeight_       As Double
Private MaxConditionedWeight_       As Double
Private YarnFinish_                 As String
Private YarnCode_                   As String
Private MoistureRegain_             As Double
Private Twisting_                   As Double
Private YarnColor_                  As String
Private Notes_                      As String
Private YarnMerge_                  As String
Private Properties_                 As Object

'------------
' COLLECTIONS
'------------
'------------
' PROPERTIES
'------------
' Properties
Public Property Get Properties() As Object
    Set Properties = Properties_
End Property
Public Property Set Properties(dict As Object)
    dict.Add key:="Dtex", item:=Dtex
    dict.Add key:="Style", item:=Style
    dict.Add key:="WeaveType", item:=WeaveType
    dict.Add key:="YarnType", item:=YarnType
    dict.Add key:="Denier", item:=Denier
    dict.Add key:="MeanWarpCount", item:=MeanWarpCount
    dict.Add key:="MinWarpCount", item:=MinWarpCount
    dict.Add key:="MaxWarpCount", item:=MaxWarpCount
    dict.Add key:="MeanFillCount", item:=MeanFillCount
    dict.Add key:="MinFillCount", item:=MinFillCount
    dict.Add key:="MaxFillCount", item:=MaxFillCount
    dict.Add key:="MeanDryWeight", item:=MeanDryWeight
    dict.Add key:="MinDryWeight", item:=MinDryWeight
    dict.Add key:="MaxDryWeight", item:=MaxDryWeight
    dict.Add key:="MeanConditionedWeight", item:=MeanConditionedWeight
    dict.Add key:="MinConditionedWeight", item:=MinConditionedWeight
    dict.Add key:="MaxConditionedWeight", item:=MaxConditionedWeight
    dict.Add key:="YarnFinish", item:=YarnFinish
    dict.Add key:="YarnCode", item:=YarnCode
    dict.Add key:="MoistureRegain", item:=MoistureRegain
    dict.Add key:="Twisting", item:=Twisting
    dict.Add key:="YarnColor", item:=YarnColor
    dict.Add key:="Notes", item:=Notes
    dict.Add key:="YarnMerge", item:=YarnMerge
    Set Properties_ = dict
End Property
' MinFillCount
Public Property Let MinFillCount(value As Double)
    MinFillCount_ = value
End Property
Public Property Get MinFillCount() As Double
    MinFillCount = MinFillCount_
End Property
' YarnMerge
Public Property Let YarnMerge(value As String)
    YarnMerge_ = value
End Property
Public Property Get YarnMerge() As String
    YarnMerge = YarnMerge_
End Property
' Notes
Public Property Let Notes(value As String)
    Notes_ = value
End Property
Public Property Get Notes() As String
    Notes = Notes_
End Property
' Twisting
Public Property Let Twisting(value As Double)
    Twisting_ = value
End Property
Public Property Get Twisting() As Double
    Twisting = Twisting_
End Property
' YarnColor
Public Property Let YarnColor(value As String)
    YarnColor_ = value
End Property
Public Property Get YarnColor() As String
    YarnColor = YarnColor_
End Property
' MoistureRegain
Public Property Let MoistureRegain(value As Double)
    MoistureRegain_ = value
End Property
Public Property Get MoistureRegain() As Double
    MoistureRegain = MoistureRegain_
End Property
' YarnCode
Public Property Let YarnCode(value As String)
    YarnCode_ = value
End Property
Public Property Get YarnCode() As String
    YarnCode = YarnCode_
End Property
' YarnFinish
Public Property Let YarnFinish(value As String)
    YarnFinish_ = value
End Property
Public Property Get YarnFinish() As String
    YarnFinish = YarnFinish_
End Property
' MaxConditionedWeight
Public Property Let MaxConditionedWeight(value As Double)
    MaxConditionedWeight_ = value
End Property
Public Property Get MaxConditionedWeight() As Double
    MaxConditionedWeight = MaxConditionedWeight_
End Property
' MinConditionedWeight
Public Property Let MinConditionedWeight(value As Double)
    MinConditionedWeight_ = value
End Property
Public Property Get MinConditionedWeight() As Double
    MinConditionedWeight = MinConditionedWeight_
End Property
' MeanConditionedWeight
Public Property Let MeanConditionedWeight(value As Double)
    MeanConditionedWeight_ = value
End Property
Public Property Get MeanConditionedWeight() As Double
    MeanConditionedWeight = MeanConditionedWeight_
End Property
' MaxDryWeight
Public Property Let MaxDryWeight(value As Double)
    MaxDryWeight_ = value
End Property
Public Property Get MaxDryWeight() As Double
    MaxDryWeight = MaxDryWeight_
End Property
' MinDryWeight
Public Property Let MinDryWeight(value As Double)
    MinDryWeight_ = value
End Property
Public Property Get MinDryWeight() As Double
    MinDryWeight = MinDryWeight_
End Property
' MeanDryWeight
Public Property Let MeanDryWeight(value As Double)
    MeanDryWeight_ = value
End Property
Public Property Get MeanDryWeight() As Double
    MeanDryWeight = MeanDryWeight_
End Property
' MaxFillCount
Public Property Let MaxFillCount(value As Double)
    MaxFillCount_ = value
End Property
Public Property Get MaxFillCount() As Double
    MaxFillCount = MaxFillCount_
End Property
' MinWarpCount
Public Property Let MinWarpCount(value As Double)
    MinWarpCount_ = value
End Property
Public Property Get MinWarpCount() As Double
    MinWarpCount = MinWarpCount_
End Property
' MaxWarpCount
Public Property Let MaxWarpCount(value As Double)
    MaxWarpCount_ = value
End Property
Public Property Get MaxWarpCount() As Double
    MaxWarpCount = MaxWarpCount_
End Property
' MeanFillCount
Public Property Let MeanFillCount(value As Double)
    MeanFillCount_ = value
End Property
Public Property Get MeanFillCount() As Double
    MeanFillCount = MeanFillCount_
End Property
' MeanWarpCount
Public Property Let MeanWarpCount(value As Double)
    MeanWarpCount_ = value
End Property
Public Property Get MeanWarpCount() As Double
    MeanWarpCount = MeanWarpCount_
End Property
' Denier
Public Property Let Denier(value As Double)
    Denier_ = value
End Property
Public Property Get Denier() As Double
    Denier = Denier_
End Property
' YarnType
Public Property Let YarnType(value As String)
    YarnType_ = value
End Property
Public Property Get YarnType() As String
    YarnType = YarnType_
End Property
' WeaveType
Public Property Let WeaveType(value As String)
    WeaveType_ = value
End Property
Public Property Get WeaveType() As String
    WeaveType = WeaveType_
End Property
' Style
Public Property Let Style(value As Long)
    Style_ = value
End Property
Public Property Get Style() As Long
    Style = Style_
End Property
' Ends Per Inch
Public Property Let EndsPerInch(value As Double)
    EndsPerInch_ = value
End Property
Public Property Get EndsPerInch() As Double
    EndsPerInch = EndsPerInch_
End Property
' Linear Density in Dtex or g / 10000 m
Public Property Let Dtex(value As Double)
    Dtex_ = value
End Property
Public Property Get Dtex() As Double
    Dtex = Dtex_
End Property
'----------
' METHODS
'----------
Private Sub SetProperties()
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
Public Sub ISpec_Commit()
    SaveSpecification "tblStyleSpecCommits"
End Sub
Public Sub ISpec_Push()
    SaveSpecification "tblStyleSpecs"
End Sub
Private Sub SaveSpecification()
' Adds a new specification to the database
    Dim key As Variant
    Dim SQLstmt As String
    Dim INSERTstmt, VALUESstmt As String
    Dim tbl As String: tbl = "tblStyleSpecs"
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
'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
End Sub
' Deconstructor
Private Sub Class_Terminate()
End Sub
