Attribute VB_Name = "ComService"
Option Explicit

Public Function GetStandardJson(material_id As String) As String
' Calls to the COM Server for a specification standard in json ie rev X.0
    GetStandardJson = manager.server.GetStandardJson(material_id)
End Function

Public Function GetSpecJson(material_id As String) As String
' Calls to the DM.NET COM manager.server for an object returned as a json string
' This json string will need to be unpacked to build the spec object
    GetSpecJson = manager.server.GetSpecJson(material_id)
End Function

Public Function GetSpecTemplate(spec_type As String) As String
' Calls to the COM server for an object representing a custom spec template.
    GetSpecTemplate = manager.server.GetSpecTemplate(spec_type)
End Function

Public Function PushSpecJson(spec As Specification, Optional is_default As Boolean = False) As Long
' Sends a json string object to DM.NET for update
    PushSpecJson = manager.server.PushSpecJson( _
        spec.ObjectToJson, spec.SpecType, spec.MaterialId, _
        IIf(is_default, spec.Revision + 1, spec.Revision + 0.1), _
        is_default)
        
End Function

Public Function PushSpecTemplate(template As SpecTemplate) As Long
' Sends a json string representing a specification template to be stored in the database
    PushSpecTemplate = manager.server.PushSpecTemplate( _
        template.PropertiesJson, template.SpecType, template.Revision + 1)

End Function

