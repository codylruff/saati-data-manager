Attribute VB_Name = "modCover"
Option Explicit

'========================================================================
'Cover Macro to Test the Progress Bar
'========================================================================
Sub TestTheBar()
Attribute TestTheBar.VB_ProcData.VB_Invoke_Func = "Q\n14"

    'Declaring Sub Level Variables
    Dim lngCounter As Long
    Dim lngNumberOfTasks As Long

    'Initilaizing Variables
    lngNumberOfTasks = 10000

    'Calling the ShowProgress sub with ActionNumber = 0, to let the
    'user know we are going to work on the 1st task. Also, set a
    'title for the form
    Call modProgress.ShowProgress( _
                        0, _
                        lngNumberOfTasks, _
                        "Excel is working on Task Number 1", _
                        False, _
                        "Progress Bar Test")

    For lngCounter = 1 To lngNumberOfTasks
        'The code for each task goes here
        
        'You can add your code here

        'Call the ShowProgress sub each time a task is finished to
        'the user know that X out of Y tasks are over, and that
        'the X+1'th task is in progress.
        Call modProgress.ShowProgress( _
                    lngCounter, _
                    lngNumberOfTasks, _
                    "Excel is working on Task Number " & lngCounter + 1, _
                    False)
        
    Next lngCounter

End Sub

'========================================================================
'A Macro to Illustrate the use of Application.Statusbar property
'========================================================================
Sub StatusBarExample()

    'Declaring Sub Level Variables
    Dim lngCounter As Long
    Dim lngNumberOfTasks As Long

    'Initilaizing Variables
    lngNumberOfTasks = 1000

    For lngCounter = 1 To lngNumberOfTasks
        'Altering the Statusbar Property
        Application.StatusBar = "Executing " & lngCounter & _
            " of " & lngNumberOfTasks & " | " & _
            "Custom Message " & lngCounter
    Next lngCounter

    'Letting Excel Take over the status bar
    Application.StatusBar = False

End Sub
