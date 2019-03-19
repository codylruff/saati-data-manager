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

Public fso As FileSystemObject ' Declare a FileSystemObject.
Public stream As TextStream ' Declare a TextStream.

Public Sub Log(text As String)
    stream.WriteLine text
End Sub

stream.Close

public sub class_initialize()
    Set fso = New FileSystemObject
    Set stream = fso.CreateTextFile(LocalDir, True)
end sub
  
  
