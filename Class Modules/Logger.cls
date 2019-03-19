VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Name_ As String
Private fso As FileSystemObject ' Declare a FileSystemObject.
Private stream As TextStream ' Declare a TextStream.
Private buffer As Dictionary

Public Property Get Name() As String: Name = Name_: End Property
Public Property Let Name(value As String): Name_ = value: End Property
Public Property Get FilePath() As String
    FilePath = "C:\Users\" & manager.current_user & "\Documents\DataManager\" & Me.Name & "_" & CStr(Now) & ".txt"
End Property


Public Sub Log(text As String)
    buffer.Add key:=CStr(Now), Item:=text
    Debug.Print Utils.printf("{0} : {1}", CStr(Now), text)
End Sub

Private Sub SaveLog()
    Dim key As Variant
    If Not fso.FolderExists(Me.FilePath) Then fso.CreateFolder Me.FilePath
    Set stream = fso.CreateTextFile(Me.FilePath, True)
    For Each key In buffer
      stream.WriteLine Utils.printf("{0} : {1}", key, buffer.Item(key))
    Next key
    stream.Close
End Sub

Private Sub Class_Initialize()
    Set fso = New FileSystemObject
    Set buffer = New Dictionary
End Sub
  
Private Sub Class_Terminate()
  If buffer.count > 0 Then SaveLog
  Set fso = Nothing
  Set buffer = Nothing
End Sub
