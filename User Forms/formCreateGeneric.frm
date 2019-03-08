VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCreateGeneric 
   Caption         =   "Create New Spec Type"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9285
   OleObjectBlob   =   "formCreateGeneric.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCreateGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private template_name As String

Private Sub UserForm_Initialize()
    template_name = InputBox("Enter a template name :", "Custom Template Name")
    lblInstructions.Caption = " Instructions :" & vbNewLine & _
         " Create the template parameters one at a time," & _
         " selecting a parameter type (text, number, True/False)," & _
         " entering the parameter name, and clicking the 'Set Property'" & _
         " button. The template can be saved by clicking the save button."
End Sub
