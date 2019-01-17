Attribute VB_Name = "Developer"
Option Explicit

Public Const GitBashExe As String = "C:\Users\cruff\AppData\Local\Programs\Git\git-bash.exe"
Public Const GitRepo As String = "C:\Users\cruff\source\DataManager\DataManager"

Public bIsVS As Boolean
Public bDebugMessages As Boolean
Public bIsTesting As Boolean

Public Sub IsTesting(TurnOn As Boolean)
' Switches to the local database when for debug/testing if IsTesting(True)
    If TurnOn Then
        bIsTesting = True
    Else
        bIsTesting = False
    End If
    
End Sub

Public Sub DebugBox(sText As String)

    If bDebugMessages Then MsgBox _
        Prompt:=sText, _
        Title:="Debug Message"
    
End Sub

Sub VSExport()

    ExportAll True, True
    
End Sub

Sub ExportAll(IsVS As Boolean, IsTest As Boolean, Optional VCTable As String)

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
    Dim VerNum As String
    Dim NewVersion As String
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
        
        DebugBox "Successfully exported " & CStr(count) & " VBA files to " & directory
    
End Sub

