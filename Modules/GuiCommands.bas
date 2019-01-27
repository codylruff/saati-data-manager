Attribute VB_Name = "GuiCommands"
Option Explicit
'=================================
' DESCRIPTION: Holds commands used
' through the GUI with exception
' of the import function.
'=================================

Public Sub GoToMain()
'Opens the main menu form.
    Application.Visible = False
    Utility.UnloadAllForms
    formMainMenu.Show
End Sub

Public Sub ExportAll()
' Exports the codebase to a project folder as text files
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As Object
    Dim count As Integer
    Dim Path As String
    Dim directory As String
    Dim extension As String
    Dim lngCounter As Long
    Dim lngNumberOfTasks As Long

    lngNumberOfTasks = 4
    lngCounter = 0

    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Creating a New Version...", _
        False)
        
    directory = GitRepo & "\"
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        1, _
        lngNumberOfTasks, _
        "Saving...", _
        False, _
        "Spec Manager")
    
    count = 0
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Creating Directory...", _
        False)
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Exporting Code Modules...", _
        False)

    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        
        If VBComponent.Type <> Document Then
            Select Case VBComponent.Type
                Case ClassModule
                    extension = ".cls"
                    Path = directory & "Class Modules\" & VBComponent.name & extension
                Case Form
                    extension = ".frm"
                    Path = directory & "User Forms\" & VBComponent.name & extension
                    
                Case Module
                    extension = ".bas"
                    Path = directory & "Modules\" & VBComponent.name & extension
                    
                Case Else
                    extension = ".txt"
            End Select
            
            On Error Resume Next
            Err.Clear
            
            
            Call VBComponent.Export(Path)
            
            If Err.Number <> 0 Then
                Debug.Print "Failed to export " & VBComponent.name & " to " & Path
            Else
                count = count + 1
                Debug.Print "Exported " & Left$(VBComponent.name & ":" & Space(Padding), Padding) & Path
            End If

            On Error GoTo 0
        End If

    Next
    
    lngCounter = lngCounter + 1
    Call modProgress.ShowProgress( _
        lngCounter, _
        lngNumberOfTasks, _
        "Finishing...", _
        False)
    
End Sub

Public Sub ConfigControl()
'Initializes the password form for config access.
    If Environ("UserName") <> "CRuff" Then
        formPassword.Show
    Else
        Application.DisplayAlerts = True
        shtDeveloper.Visible = xlSheetVisible
        Application.Visible = True
        Application.VBE.MainWindow.Visible = True
        Application.SendKeys ("^r")
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
    GuiCommands.GoToMain
End Sub

Public Sub ExitApp()
'This exits the application after saving the thisworkbook.
    ThisWorkbook.Save
    Application.Quit
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

Public Sub DB2W_tblWarpingSpecs()
' Dumps warping specs to a new worksheet
    DataAccess.DatabaseToWorksheet Factory.CreateSQLiteDatabase, SQLITE_PATH, "tblWarpingSpecs"
End Sub

Public Sub DB2W_tblStyleSpecs()
' Dumps style specs to a new worksheet
    DataAccess.DatabaseToWorksheet Factory.CreateSQLiteDatabase, SQLITE_PATH, "tblStyleSpecs"
End Sub