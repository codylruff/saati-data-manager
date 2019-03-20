VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_header    As Variant
Private m_data      As Variant
Public Fields       As Dictionary
Public records      As Collection

Public Property Get header() As Variant: header = m_header: End Property
Public Property Let header(ByVal value As Variant): m_header = value: End Property
Public Property Get data() As Variant: data = m_data: End Property
Public Property Let data(ByVal value As Variant): m_data = value: End Property

Public Function rows() As Long
    If NumberOfArrayDimensions = 1 Then
        rows = 1
    Else
        rows = UBound(Me.data, 2)
    End If
End Function

Public Function columns() As Long
    columns = UBound(Me.header)
End Function

Private Function NumberOfArrayDimensions() As Integer
    Dim Ndx As Integer
    Dim Res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
        Do
            Ndx = Ndx + 1
            Res = UBound(Me.data, Ndx)
        Loop Until Err.Number <> 0
    NumberOfArrayDimensions = Ndx - 1
End Function

Public Function SetDictionary()
' Returns a dictionary representation of this object
    Dim dict: Set dict = New Dictionary
    Dim coll: Set coll = New Collection
    Dim i, j As Long
    If Me.rows = 1 Then
        Set dict = Factory.CreateDictionary()
        For i = LBound(Me.data) To Me.columns
            dict.Add key:=Me.header(1, i), Item:=Me.data(1, i)
        Next i
    Else
        For i = LBound(Me.data) To Me.columns
            Set dict = Factory.CreateDictionary()
            For j = LBound(Me.data) To Me.rows
            Logger.Log "Column " & j & " : " & header(1, j) & ", Row " & i & " : " & data(i, j)
                dict.Add key:=Me.header(1, j), Item:=Me.data(i, j)
            Next j
            coll.Add dict
        Next i
    End If
    Set records = coll
End Function
