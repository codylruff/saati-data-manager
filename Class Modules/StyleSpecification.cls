VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleSpecification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Classes")
'=========================================
'CLASS DESCRIPTION:
' Represents the spec for a specific style
'=========================================
Implements ISpec

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

'------------
' COLLECTIONS
'------------
Private Properties_                 As Object
' Properties
Public Property Get ISpec_Properties() As Object
    Set ISpec_Properties = Properties_
End Property
Public Property Set ISpec_Properties(dict As Object)
    dict.Add key:="Dtex", Item:=Dtex
    dict.Add key:="Style", Item:=Style
    dict.Add key:="WeaveType", Item:=WeaveType
    dict.Add key:="YarnType", Item:=YarnType
    dict.Add key:="Denier", Item:=Denier
    dict.Add key:="MeanWarpCount", Item:=MeanWarpCount
    dict.Add key:="MinWarpCount", Item:=MinWarpCount
    dict.Add key:="MaxWarpCount", Item:=MaxWarpCount
    dict.Add key:="MeanFillCount", Item:=MeanFillCount
    dict.Add key:="MinFillCount", Item:=MinFillCount
    dict.Add key:="MaxFillCount", Item:=MaxFillCount
    dict.Add key:="MeanDryWeight", Item:=MeanDryWeight
    dict.Add key:="MinDryWeight", Item:=MinDryWeight
    dict.Add key:="MaxDryWeight", Item:=MaxDryWeight
    dict.Add key:="MeanConditionedWeight", Item:=MeanConditionedWeight
    dict.Add key:="MinConditionedWeight", Item:=MinConditionedWeight
    dict.Add key:="MaxConditionedWeight", Item:=MaxConditionedWeight
    dict.Add key:="YarnFinish", Item:=YarnFinish
    dict.Add key:="YarnCode", Item:=YarnCode
    dict.Add key:="MoistureRegain", Item:=MoistureRegain
    dict.Add key:="Twisting", Item:=Twisting
    dict.Add key:="YarnColor", Item:=YarnColor
    dict.Add key:="Notes", Item:=Notes
    dict.Add key:="YarnMerge", Item:=YarnMerge
    Set Properties_ = dict
End Property
'-----------------
' AUTO-PROPERTIES
'-----------------
' Database table
Public Property Get ISpec_TableName() As String: ISpec_TableName = "style_specifications": End Property
' Specification Type
Public Property Get ISpec_SpecType() As String: ISpec_SpecType = "Style": End Property
' MinFillCount
Public Property Let MinFillCount(value As Double): MinFillCount_ = value: End Property
Public Property Get MinFillCount() As Double: MinFillCount = MinFillCount_: End Property
' YarnMerge
Public Property Let YarnMerge(value As String): YarnMerge_ = value: End Property
Public Property Get YarnMerge() As String: YarnMerge = YarnMerge_: End Property
' Notes
Public Property Let Notes(value As String): Notes_ = value: End Property
Public Property Get Notes() As String: Notes = Notes_: End Property
' Twisting
Public Property Let Twisting(value As Double): Twisting_ = value: End Property
Public Property Get Twisting() As Double: Twisting = Twisting_: End Property
' YarnColor
Public Property Let YarnColor(value As String): YarnColor_ = value: End Property
Public Property Get YarnColor() As String: YarnColor = YarnColor_: End Property
' MoistureRegain
Public Property Let MoistureRegain(value As Double): MoistureRegain_ = value: End Property
Public Property Get MoistureRegain() As Double: MoistureRegain = MoistureRegain_: End Property
' YarnCode
Public Property Let YarnCode(value As String): YarnCode_ = value: End Property
Public Property Get YarnCode() As String: YarnCode = YarnCode_: End Property
' YarnFinish
Public Property Let YarnFinish(value As String): YarnFinish_ = value: End Property
Public Property Get YarnFinish() As String: YarnFinish = YarnFinish_: End Property
' MaxConditionedWeight
Public Property Let MaxConditionedWeight(value As Double): MaxConditionedWeight_ = value: End Property
Public Property Get MaxConditionedWeight() As Double: MaxConditionedWeight = MaxConditionedWeight_: End Property
' MinConditionedWeight
Public Property Let MinConditionedWeight(value As Double): MinConditionedWeight_ = value: End Property
Public Property Get MinConditionedWeight() As Double: MinConditionedWeight = MinConditionedWeight_: End Property
' MeanConditionedWeight
Public Property Let MeanConditionedWeight(value As Double): MeanConditionedWeight_ = value: End Property
Public Property Get MeanConditionedWeight() As Double: MeanConditionedWeight = MeanConditionedWeight_: End Property
' MaxDryWeight
Public Property Let MaxDryWeight(value As Double): MaxDryWeight_ = value: End Property
Public Property Get MaxDryWeight() As Double: MaxDryWeight = MaxDryWeight_: End Property
' MinDryWeight
Public Property Let MinDryWeight(value As Double): MinDryWeight_ = value: End Property
Public Property Get MinDryWeight() As Double: MinDryWeight = MinDryWeight_: End Property
' MeanDryWeight
Public Property Let MeanDryWeight(value As Double): MeanDryWeight_ = value: End Property
Public Property Get MeanDryWeight() As Double: MeanDryWeight = MeanDryWeight_: End Property
' MaxFillCount
Public Property Let MaxFillCount(value As Double): MaxFillCount_ = value: End Property
Public Property Get MaxFillCount() As Double: MaxFillCount = MaxFillCount_: End Property
' MinWarpCount
Public Property Let MinWarpCount(value As Double): MinWarpCount_ = value: End Property
Public Property Get MinWarpCount() As Double: MinWarpCount = MinWarpCount_: End Property
' MaxWarpCount
Public Property Let MaxWarpCount(value As Double): MaxWarpCount_ = value: End Property
Public Property Get MaxWarpCount() As Double: MaxWarpCount = MaxWarpCount_: End Property
' MeanFillCount
Public Property Let MeanFillCount(value As Double): MeanFillCount_ = value: End Property
Public Property Get MeanFillCount() As Double: MeanFillCount = MeanFillCount_: End Property
' MeanWarpCount
Public Property Let MeanWarpCount(value As Double): MeanWarpCount_ = value: End Property
Public Property Get MeanWarpCount() As Double: MeanWarpCount = MeanWarpCount_: End Property
' Denier
Public Property Let Denier(value As Double): Denier_ = value: End Property
Public Property Get Denier() As Double: Denier = Denier_: End Property
' YarnType
Public Property Let YarnType(value As String): YarnType_ = value: End Property
Public Property Get YarnType() As String: YarnType = YarnType_: End Property
' WeaveType
Public Property Let WeaveType(value As String): WeaveType_ = value: End Property
Public Property Get WeaveType() As String: WeaveType = WeaveType_: End Property
' Style
Public Property Let Style(value As Long): Style_ = value: End Property
Public Property Get Style() As Long: Style = Style_: End Property
' Linear Density in Dtex or g / 10000 m
Public Property Let Dtex(value As Double): Dtex_ = value: End Property
Public Property Get Dtex() As Double: Dtex = Dtex_: End Property
'----------
' METHODS
'----------
Public Sub SetProperties()
    Set ISpec_Properties = CreateObject("Scripting.Dictionary")
End Sub
Public Sub SaveSpecification()
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

    For Each key In ISpec_Properties
        INSERTstmt = INSERTstmt & key & ", "
        VALUESstmt = VALUESstmt & "'" & ISpec_Properties(key) & "', "
    Next key

    INSERTstmt = INSERTstmt & "Time_Stamp) "
    VALUESstmt = VALUESstmt & "'" & Now() & "')"

    ' Create SQL statement from object
    SQLstmt = INSERTstmt & vbNewLine & VALUESstmt
    'Debug.Print SQLstmt
    ExecuteSQL Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt
End Sub

Private Property Get ISpec_StyleSpec() As StyleSpecification
    Err.Raise 5 'TODO implement interface member
End Property

Public Sub ISpec_JsonToObject(jsonText As String)
    ' Convert a json string to a warp spec
    Dim dict: Set dict = JsonConverter.ParseJson(jsonText)
    On Error Resume Next
    With Me
        .Dtex = dict.Item("Dtex")
        .Style = dict.Item("Style")
        .WeaveType = dict.Item("WeaveType")
        .YarnType = dict.Item("YarnType")
        .Denier = dict.Item("Denier")
        .MeanWarpCount = dict.Item("MeanWarpCount")
        .MinWarpCount = dict.Item("MinWarpCount")
        .MaxWarpCount = dict.Item("MaxWarpCount")
        .MeanFillCount = dict.Item("MeanFillCount")
        .MinFillCount = dict.Item("MinFillCount")
        .MaxFillCount = dict.Item("MaxFillCount")
        .MeanDryWeight = dict.Item("MeanDryWeight")
        .MinDryWeight = dict.Item("MinDryWeight")
        .MaxDryWeight = dict.Item("MaxDryWeight")
        .MeanConditionedWeight = dict.Item("MeanConditionedWeight")
        .MinConditionedWeight = dict.Item("MinConditionedWeight")
        .MaxConditionedWeight = dict.Item("MaxConditionedWeight")
        .YarnFinish = dict.Item("YarnFinish")
        .YarnCode = dict.Item("YarnCode")
        .MoistureRegain = dict.Item("MoistureRegain")
        .Twisting = dict.Item("Twisting")
        .YarnColor = dict.Item("YarnColor")
        .Notes = dict.Item("Notes")
        .YarnMerge = dict.Item("YarnMerge")
    End With
    Me.SetProperties
End Sub

Private Sub ISpec_SetDefaultProperties()
    Err.Raise 5 'TODO implement interface member
End Sub

