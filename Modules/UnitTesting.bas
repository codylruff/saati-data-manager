Attribute VB_Name = "UnitTesting"
Option Explicit

Public Sub SQLite3Connection_Test()
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    Dim RetVal As Long
    Dim recordsAffected As Long
    Dim InitReturn As Long
    Dim ErrorVal As Long
    ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    InitReturn = SQLite3Initialize
    
    If InitReturn <> SQLITE_INIT_OK Then
        MsgBox "Error Initializing SQLite. Error: " & Err.LastDllError & "Contact Admin"
        Exit Sub
    End If

    Dim stepMsg As String
    
    ' Open the database - getting a DbHandle back
    RetVal = SQLite3Open(PathToSQLite3Database, myDbHandle)
    If RetVal = SQLITE_OK Then
        MsgBox "Connected Succesfully"
    Else
        MsgBox "SQLite3Open() returned: " & RetVal
    End If
End Sub

Public Function WarpingSpecificationTest()
    Dim spec As WarpingSpecification

    Set spec = New WarpingSpecification

    With spec
        .MaterialNumber = "OKE101011AP00CN"
        .MaterialDescription = "STY 101 K29 3300D 1F279 131CM"
        .FinalWidthCm = 103.5
        .WarpingSpeed = 300
        .BeamingSpeed = 80
        .CrossWinding = 10
        .DentsPerCm = 0
        .EndsPerDent = 0
        
    End With
    spec.SaveSpecification
End Function

Public Sub select_Test()
    Dim record As DatabaseRecord
    Dim i, j As Long
    Set record = ExecuteSQLSelect(Factory.CreateSQLiteDatabase, SQLITE_PATH, "SELECT * FROM tblWarpingSpecs WHERE spec_id = 15")
    If record.rows = 1 Then
        For i = LBound(record.data) To record.columns
            Debug.Print record.header(1, i), record.data(1, i)
        Next i
    Else
        For i = LBound(record.data) To record.columns
            For j = LBound(record.data) To record.rows
                Debug.Print record.header(1, j), record.data(i, j)
            Next j
        Next i
    End If
End Sub

Public Sub Retrieve_Test()
    Const Padding = 24

    Dim key As Variant
    Dim spec As WarpingSpecification
    
    Set spec = RetrieveWarpingSpecification("OKE101011AP00CN")
    'Set w = Factory.CreateWarp(spec, NumberOfBobbins, PackageWeightlbs, WarpLength)
    
    For Each key In spec.Properties
        Debug.Print Left$(key & ":" & Space(Padding), _
                    Padding) & spec.Properties(key)
    Next key
    

End Sub
