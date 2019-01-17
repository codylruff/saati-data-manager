Attribute VB_Name = "PublicVars"
Option Explicit

' Global Constants
Public Const GsmToOsy               As Double = 0.0294935
Public Const PublicDir              As String = "S:\Data Manager"
Public Const LocalDir               As String = "W:\App Development\Spec Manager"
Public Const sDHConStr              As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=S:\Data Manager\DataHistory.accdb; Persist Security Info=False;"
Public Const PathToSQLite3Database  As String = "W:\App Development\Data Manager\Protection_Quality_Control.db3"
Public Const VCConStr               As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=S:\Data Manager\VersionControl\VersionControl.accdb; Persist Security Info=False;"
Public Const sQualityInfoConStr     As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=S:\Data Manager\QualityInfo.accdb; Persist Security Info=False;"


'======================================
'NOTES: A List of all product variables
'that are initialized through the UI.
'======================================


' Data Structure Objects

' System Objects
Public bIsAdmin                     As Boolean
Public sUserName                    As String




