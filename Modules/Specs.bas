Attribute VB_Name = "Specs"

Public Function Specs()
    'Set Specs = New SpecSuite
    'Specs.Description = "Data Manger"
    'I would like to use real unit testing eventually
    On Error Resume Next
    ''
    ' Convert JSON string to object (Dictionary/Collection)
    '
    ' @method ParseJson
    ' @param {String} json_String
    ' @return {Object} (Dictionary or Collection)
    ' @throws 10001 - JSON parse error
    ''

    ''
    ' Convert object (Dictionary/Collection/Array) to JSON
    '
    ' @method ConvertToJson
    ' @param {Variant} JsonValue (Dictionary, Collection, or Array)
    ' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
    ' @return {String}
    ''
    Dim SQLstmt As String
    Dim record As DatabaseRecord
    Dim dict As Dictionary
    Dim key As Variant
    Dim warpSpec As WarpingSpecification
    Dim styleSpec As StyleSpecification
    Dim MaterialNumber As String

    MaterialNumber = "NKE003279MP00DF"
    Set warpSpec = New WarpingSpecification
    Set styleSpec = RetrieveStyleSpecification(Mid(MaterialNumber, 6, 3))
    Set warpSpec.styleSpec = styleSpec
    Set record = New DatabaseRecord

    SQLstmt = "SELECT * FROM tblWarpingSpecs " & _
              "WHERE MaterialNumber = """ & MaterialNumber & """;"
    'Debug.Print SQLstmt
    Set record = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, SQLITE_PATH, SQLstmt)
    Set dict = record.GetDictionary
End Function