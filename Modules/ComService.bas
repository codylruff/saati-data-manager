Attribute VB_Name = "ComService"
Option Explicit

Public Function GetSpecJson(material_id As String) As String
' Calls to the COM server for an object returned as a json string
' This json string will need to be unpacked to build the spec object
    Dim server As Object
    Set server = CreateObject("DM_LIB.DmComServer")
    On Error Goto NullSpecException
    GetSpecJson = server.GetSpecJson(material_id)
NullSpecException:
    GetSpecJson = vbNullString
End Function

Public Function SendSpecJson(spec As Specification) As Long
' Sends a json string object to DM.NET for update
    Dim server As Object
    Set server = CreateObject("DM_LIB.DmComServer")
    SendSpecJson = server.SendSpecJson(spec.ObjectToJson, spec.SpecType, spec.MaterialId)
End Function
