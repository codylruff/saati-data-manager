Dim fso As FileSystemObject ' Declare a FileSystemObject.
Set fso = New FileSystemObject ' Create a FileSystemObject.
Dim stream As TextStream ' Declare a TextStream.

Set stream = fso.CreateTextFile("C:\output.txt", True)
stream.WriteLine "abc"
stream.WriteLine "def"
stream.Close
