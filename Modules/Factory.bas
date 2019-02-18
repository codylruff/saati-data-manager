Attribute VB_Name = "Factory"

Function CreateSpecification() As Specification
    Set CreateSpecification = New Specification
End Function

Function CreateConsoleBox(frm As UserForm) As ConsoleBox
    Dim console As ConsoleBox
    Set console = New ConsoleBox
    Set console.FormId = frm
    Set CreateConsoleBox = console
End Function
