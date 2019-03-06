Attribute VB_Name = "Utils"
Option Explicit
'@Folder("Modules")
'=================================
' DESCRIPTION: Util Module holds
' miscellenous helper functions.
'=================================
Sub ToggleExcelGui(b As Boolean)
' Disables unpleasent ui effects
    Application.ScreenUpdating = b
    Application.DisplayAlerts = b
End Sub

Function ConvertToCamelCase(s As String) As String
' Converts sentence case to Camel Case
    With CreateObject("VBScript.RegExp")
        .Pattern = "[^a-zA-Z]"
        .Global = True
        ConvertToCamelCase = replace(StrConv(.replace(s, " "), vbProperCase), " ", "")
    End With
End Function

Function SplitCamelCase(sString As String, Optional sDelim As String = " ") As String
' Converts camel case to sentence case
On Error GoTo Error_Handler
    Dim oRegEx          As Object
 
    Set oRegEx = CreateObject("vbscript.regexp")
    With oRegEx
        .Pattern = "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))"
        .Global = True
        SplitCamelCase = .replace(sString, "$1" & sDelim)
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

Function StringBuilder(ParamArray args() As Variant) As String
    builder.AppendFormat("Warp Count : {0} ({1} to {2})\n", MeanWarpCount, MinWarpCount, MaxWarpCount)
    builder.AppendFormat("Fill Count : {0} ({1} to {2})\n", MeanFillCount, MinFillCount, MaxFillCount)
    builder.AppendFormat("Dry Weight : {0} ({1} to {2})\n", MeanDryWeight, MinDryWeight, MaxDryWeight)
    builder.AppendFormat("Conditioned Weight : {0} ({1} to {2})\n", _
        MeanConditionedWeight, MinConditionedWeight, MaxConditionedWeight)
End Function

Public Function printf(mask As String, ParamArray tokens()) As String
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function