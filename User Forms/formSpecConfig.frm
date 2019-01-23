VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formSpecConfig 
   Caption         =   "Specification Control"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735
   OleObjectBlob   =   "formSpecConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formSpecConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim spec As WarpingSpecification

Private Sub cmdBack_Click()
    Unload Me
    GoToMain
End Sub

Private Sub cmdSaveChanges_Click()
    spec.SaveSpecification
End Sub

Private Sub cmdSubmit_Click()
    Dim Console As ConsoleBox
    With spec
        Select Case cboSelectProperty.value
            Case "Style"
                MsgBox "Style is a read-only property"
                Exit Sub
            Case "Material Number"
                MsgBox "Material Number is a read-only property"
                Exit Sub
            Case "Material Description"
                MsgBox "Material Description is a read-only property"
                Exit Sub
            Case "Yarn Code"
                MsgBox "Yarn Code is a read-only property"
                Exit Sub
            Case "Final Width [cm]"
                .FinalWidthCm = txtPropertyValue
            Case "Number Of Ends"
                .NumberOfEnds = txtPropertyValue
            Case "Spring Color"
                MsgBox "Spring Color is a read-only property"
                Exit Sub
            Case "Warping Speed [m/min]"
                .WarpingSpeed = txtPropertyValue
            Case "Beaming Speed [m/min]"
                .BeamingSpeed = txtPropertyValue
            Case "WarpingTension"
                .WarpingTension = txtPropertyValue
            Case "Beaming Tension"
                .BeamingTension = txtPropertyValue
            Case "Cross Winding"
                .CrossWinding = txtPropertyValue
            Case "Dent/cm"
                .DentsPerCm = txtPropertyValue
            Case "End/Dent"
                .EndsPerDent = txtPropertyValue
            Case "Beam Width"
                .BeamWidth = txtPropertyValue
            Case "S Wrap On/Off"
                If txtPropertyValue = "ON" Then
                    .IsSWrapped = True
                Else
                    .IsSWrapped = False
                End If
            End Select
    End With
    spec.IPrint_SetPrettyProperties
    Set Console = Factory.CreateConsoleBox(Me)
    ' Print object to console
    Console.PrintObject spec
End Sub

Private Sub cmdSearch_Click()
    Dim key As Variant
    Set spec = PrintWarpingSpecification(Me)
    With cboSelectProperty
        For Each key In spec.IPrint_PrettyProperties
            .AddItem key
        Next key
    End With
End Sub

Private Sub cmdClear_Click()
'Clears the form
    ClearForm Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Set spec = New WarpingSpecification
End Sub
