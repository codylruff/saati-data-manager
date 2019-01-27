Attribute VB_Name = "Utility"
Option Explicit
'=================================
' DESCRIPTION: Util Module holds
' miscellenous helper functions.
'=================================
Sub CreateNewSheet(shtName As String)
' Creates a new worksheet with the given name
    Dim exists As Boolean, i As Integer
    With ThisWorkbook
        For i = 1 To Worksheets.count
            If Worksheets(i).name = shtName Then
                exists = True
            End If
        Next i
        If exists = True Then
            .Sheets(shtName).Delete
        End If
        .Sheets.Add(After:=.Sheets(.Sheets.count)).name = shtName
    End With
Function CheckForEmpties(frm) As Boolean
'Clears the values from a user form.
    Dim ctl As Control
    For Each ctl In frm.Controls
        Select Case VBA.TypeName(ctl)
            Case "TextBox"
                If ctl.value = "" Then
                    MsgBox "All boxes must be filed.", vbExclamation, "Input Error"
                    ctl.SetFocus
                    CheckForEmpties = True
                    Exit Function
                End If
            Case "ComboBox"
                If ctl.value = "" Then
                    MsgBox "Make a selection from the drop down menu.", vbExclamation, "Input Error"
                    ctl.SetFocus
                    CheckForEmpties = True
                    Exit Function
                End If
        End Select
    Next ctl
    CheckForEmpties = False
End Function

Sub CreateZipFile(folderToZipPath As Variant, _
                        zippedFileFullName As Variant)
' Zips files given path and filename with extension.
    Dim ShellApp As Object
    'Create an empty zip file
    Open zippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    'Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).Items
    'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until ShellApp.Namespace(zippedFileFullName).Items.count = ShellApp.Namespace(folderToZipPath).Items.count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0
End Sub

Sub UnloadAllForms(Optional dummyVariable As Byte)
'Unloads all open user forms
    Dim i As Long
    For i = VBA.UserForms.count - 1 To 0 Step -1
        Unload VBA.UserForms(i)
    Next
End Sub

Sub UpdateTable(shtName As String, tblName As String, header As String, val)
'Adds an entry at the bottom of specified column header.
    Dim rng As Range
    Set rng = Sheets(shtName).Range(tblName & "[" & header & "]")
    rng.End(xlDown).offset(1, 0).value = val
End Sub

Sub Update(rng As Range, val)
'Adds an entry at the bottom of specified column header.
    rng.End(xlDown).offset(1, 0).value = val
        
End Sub

Sub Insert(rng As Range, val)
'Inserts an entry into a specific named cell.
    rng.value = val
End Sub
