Attribute VB_Name = "Json"
' dict.Add key:="MaterialNumber", item:=MaterialNumber
' dict.Add key:="MaterialDescription", item:=MaterialDescription
' dict.Add key:="FabricWidth", item:=FabricWidth
' dict.Add key:="EndsPerInch", item:=EndsPerInch
' dict.Add key:="Dtex", item:=Dtex
' dict.Add key:="IsSWrapped", item:=IsSWrapped
' dict.Add key:="SpringColor", item:=SpringColor
' dict.Add key:="WarpingSpeed", item:=WarpingSpeed
' dict.Add key:="BeamingSpeed", item:=BeamingSpeed
' dict.Add key:="CrossWinding", item:=CrossWinding
' dict.Add key:="DentsPerCm", item:=DentsPerCm
' dict.Add key:="EndsPerDent", item:=EndsPerDent
' dict.Add key:="Style", item:=Style

'     Dim key As Variant
'     Dim SQLstmt As String
'     Dim INSERTstmt, VALUESstmt As String
'     Dim tbl As String

'     ' Create the insert portion of the statement
'     tbl = "tblWarpingSpecs"
'     INSERTstmt = "INSERT INTO " & tbl & " (, "
'     ' Create the values portion of the statement
'     VALUESstmt = "VALUES ("
'     For Each key In Properties
'         INSERTstmt = INSERTstmt & key & ", "
'         VALUESstmt = VALUESstmt & "'" & Properties(key) "', "
'     Next key
'     INSERTstmt = INSERTstmt & "Time_Stamp)"
'     VALUESstmt = VALUESstmt & "'" & Now() & "')"
'     ' Create SQL statement from object
'     SQLstmt = INSERTstmt & VALUESstmt
'     Debug.Print SQLstmt
'     RetVal = ExecuteSQLite3(SQLstmt)
'     If Not RetVal = SQLITE_DONE Then Err.Raise Number:=1
