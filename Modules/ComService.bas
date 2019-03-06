Attribute VB_Name = "ComService"
Option Explicit

Public Function GetStandardJson(material_id As String) As String
    
    On Error GoTo NullSpecException
    GetStandardJson = App.server.GetStandardJson(material_id)
    Exit Function
    
NullSpecException:
    GetStandardJson = vbNullString
End Function

Public Function GetSpecJson(material_id As String) As String
' Calls to the DM.NET COM App.server for an object returned as a json string
' This json string will need to be unpacked to build the spec object
    Set App.server = CreateObject("DM_LIB.DmComServer")
    On Error GoTo NullSpecException
    GetSpecJson = App.server.GetSpecJson(material_id)
    Exit Function

NullSpecException:
    GetSpecJson = vbNullString
End Function

Public Function PushSpecJson(spec As Specification, Optional is_default As Boolean = False) As Long
' Sends a json string object to DM.NET for update
    PushSpecJson = App.server.PushSpecJson( _
        spec.ObjectToJson, spec.SpecType, spec.MaterialId, _
        IIf(is_default, spec.Revision + 1, spec.Revision + 0.1), _
        is_default)
        
End Function

Private Function DeserializeComPackage(json_text As String)
' Takes the json data transmitted from the com App.server and un-packs it
' return ???
End Function
