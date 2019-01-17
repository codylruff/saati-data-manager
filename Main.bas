Attribute VB_Name = "Main"

Option Explicit
'=================================
'NOTES: Working on early verison
'of Main. Any buttons shown on a
'worksheet should be controlled in
'Main.bas, Developer.bas, or This-
'Workbook. Public functions are
'only allowed in this module.
'=================================

Public Function CheckForEmpties(frm) As Boolean
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

Public Sub CreateZipFile(folderToZipPath As Variant, _
                        zippedFileFullName As Variant)

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

Public Sub GoToMain()
'Opens the main menu form.
    Application.Visible = False
    UnloadAllForms
    formMainMenu.Show

End Sub

Public Sub UnloadAllForms(Optional dummyVariable As Byte)
'Unloads all open user forms

    Dim i As Long

    For i = VBA.UserForms.count - 1 To 0 Step -1
        Unload VBA.UserForms(i)
    Next

End Sub

Public Sub ExitApp()
'This exits the application after saving the thisworkbook.
    ThisWorkbook.Save
    Application.Quit
    
End Sub

Public Sub UpdateTable(shtName As String, tblName As String, header As String, val)
'Adds an entry at the bottom of specified column header.
    Dim rng As Range
    Set rng = Sheets(shtName).Range(tblName & "[" & header & "]")
    rng.End(xlDown).offset(1, 0).value = val
End Sub

Public Sub Update(rng As Range, val)
'Adds an entry at the bottom of specified column header.
    rng.End(xlDown).offset(1, 0).value = val
    
End Sub

Public Sub Insert(rng As Range, val)
'Inserts an entry into a specific named cell.
    rng.value = val

End Sub

Public Sub ClearForm(frm)
'Clears the values from a user form.
    Dim ctl As Control

    For Each ctl In frm.Controls
        Select Case VBA.TypeName(ctl)
            Case "TextBox"
                ctl.Text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
            Case Else
                End Select
    Next ctl

End Sub

'=================================================
'NOTES: Consider moving these to Developer.bas
'=================================================
Public Sub ConfigControl()
'Initializes the password form for config access.
    Dim bPriv As Boolean
    
    sUserName = Environ("UserName")
    
    If sUserName = "CRuff" Then
        bIsAdmin = True
    Else
        bIsAdmin = False
    End If
    
    bPriv = bIsAdmin
    
    If Not bPriv Then
        formPassword.Show
    Else
        Application.DisplayAlerts = True
        shtDeveloper.Visible = xlSheetVisible
        Application.Visible = True
        Application.VBE.MainWindow.Visible = True
        Application.SendKeys ("^r")
        IsTesting True
    End If
    
End Sub

Public Sub Open_Config(Password As String)
'Performs a password check and opens config.

    If Password = "@Wmp9296bm4ddw" Then
        Application.DisplayAlerts = True
        shtDeveloper.Visible = xlSheetVisible
        Application.Visible = True
        Application.VBE.MainWindow.Visible = True
        Application.SendKeys ("^r")
        Unload formPassword
        IsTesting True
    Else
        MsgBox "Access Denied", vbExclamation
        Exit Sub
    End If
    
End Sub

Public Sub CloseConfig()
'Performs actions needed to close config.
    
    ThisWorkbook.Save
    shtDeveloper.Visible = xlSheetVeryHidden
    Application.VBE.MainWindow.Visible = False
    Application.DisplayAlerts = False
    IsTesting False
    bDebugMessages = False
    GoToMain
        
End Sub
'==================================================
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'==================================================
