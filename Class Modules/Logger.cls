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
Set fso = New FileSystemObject ' Create a FileSystemObject.
Dim stream As TextStream ' Declare a TextStream.

Set stream = fso.CreateTextFile("C:\output.txt", True)
stream.WriteLine "abc"
stream.WriteLine "def"
stream.Close

public sub class_initialize()
    Set fso = New FileSystemObject
end sub
  
  
