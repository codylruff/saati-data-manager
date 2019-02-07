Attribute VB_Name = "Utility"
Option Explicit
'@Folder("Modules")
'=================================
' DESCRIPTION: Util Module holds
' miscellenous helper functions.
'=================================

'---------------------------------------------------------------------------------------
' Procedure : SplitCamelCase
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Split/Break a Camel Case string
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sString   : Camel Case string to break/split
' sDelim    : Character to use as a spcaer, if omitted will use a space
'
' Usage:
' ~~~~~~
' ?SplitCamelCase("SplitCamelCase")
'       Returns Split Camel Case
' ?SplitCamelCase("SplitCamelCase", "_")
'       Returns Split_Camel_Case
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2008-May-03             Initial Release
'---------------------------------------------------------------------------------------
Function SplitCamelCase(sString As String, Optional sDelim As String = " ") As String
' Converts camel case to sentence case
On Error GoTo Error_Handler
    Dim oRegEx          As Object
 
    Set oRegEx = CreateObject("vbscript.regexp")
    With oRegEx
        .Pattern = "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))"
        .Global = True
        SplitCamelCase = .Replace(sString, "$1" & sDelim)
    End With
 
Error_Handler_Exit:
    On Error Resume Next
    Set oRegEx = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: SplitCamelCase" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Function GetLine(ParamArray var() As Variant) As String
    Const Padding = 25
    Dim i As Integer
    Dim s As String
    s = vbNullString
    'If FormId.txtConsole = Nothing Then Exit Sub
    For i = LBound(var) To UBound(var)
         If (i + 1) Mod 2 = 0 Then
             s = s & var(i)
         Else
             s = s & Left$(var(i) & ":" & Space(Padding), Padding)
         End If
    Next
    GetLine = s & vbNewLine
End Function

Function CreateNewSheet(shtName As String) As String
' Creates a new worksheet with the given name
    Dim exists As Boolean, i As Integer
    With ThisWorkbook
        For i = 1 To Worksheets.count
            If Worksheets(i).Name = shtName Then
                exists = True
            End If
        Next i
        If exists = True Then
            .Sheets(shtName).Delete
        End If
        .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = shtName
    End With
    CreateNewSheet = shtName
End Function

Function CheckForEmpties(frm) As Boolean
'Clears the values from a user form.
    Dim ctl As Control
    For Each ctl In frm.Controls
        Select Case VBA.TypeName(ctl)
            Case "TextBox"
                If ctl.value = vbNullString Then
                    MsgBox "All boxes must be filed.", vbExclamation, "Input Error"
                    ctl.SetFocus
                    CheckForEmpties = True
                    Exit Function
                End If
            Case "ComboBox"
                If ctl.value = vbNullString Then
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

