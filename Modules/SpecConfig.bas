Attribute VB_Name = "SpecConfig"
Option Explicit
'@Folder("Modules")

Function SpecSearch(ByVal materialId As String, ByVal specType As String) As ISpec
' Sets up a query to the database and returns an ISpec object
    Dim spec As WarpingSpecification
    Dim Style As StyleSpecification

    Select Case specType
        Case "Warping"
            Set Style = New StyleSpecification
            Set spec = New WarpingSpecification
            Style = GetISpec Style, Mid(materialId, 6, 3), "style_specification"
            GetISpec spec, Me.txtSAPcode, "warping_specifications"
            Set SpecSearch = spec
        Case "Style"
            Set Style = New StyleSpecification
            Style = GetISpec(Style, materialId, "style_specification")
    End Select
End Function