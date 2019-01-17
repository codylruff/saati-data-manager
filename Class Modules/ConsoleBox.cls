VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================
'CLASS DESCRIPTION:
' Takes a user form and text box name
' and uses it as a console output.
'====================================
'Factory Module Constructor function:
'Function CreateConsoleBox(frm As UserForm) As ConsoleBox
'----------------
'CLASS VARIABLES
'----------------
Private FormId_ As MSForms.UserForm
'------------
' COLLECTIONS
'------------
Private OutputHistory_ As Collection
'------------
' PROPERTIES
'------------
' OutputHistory
Public Property Set OutputHistory(value As Collection)
    Set OutputHistory_ = value
End Property
Public Property Get OutputHistory() As Collection
    Set OutputHistory = OutputHistory_
End Property
' FormId
Public Property Set FormId(value As MSForms.UserForm)
    Set FormId_ = value
End Property
Public Property Get FormId() As MSForms.UserForm
    Set FormId = FormId_
End Property
'----------
' METHODS
'----------
Public Sub PrintLine(ParamArray var() As Variant)
    Const Padding = 25
    Dim i As Integer
    Dim s As String
    s = ""
    'If FormId.txtConsole = Nothing Then Exit Sub
    For i = LBound(var) To UBound(var)
         If (i + 1) Mod 2 = 0 Then
             s = s & var(i)
         Else
             s = s & Left$(var(i) & ":" & Space(Padding), Padding)
         End If
    Next
    OutputHistory.Add s
    FormId.txtConsole.Text = FormId.txtConsole.Text & s & vbNewLine
End Sub
'---------
' EVENTS
'---------
' Constructor
Private Sub Class_Initialize()
    Set OutputHistory = New Collection
End Sub
' Deconstructor
Private Sub Class_Terminate()
End Sub
