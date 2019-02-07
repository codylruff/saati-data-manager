Attribute VB_Name = "Constants"
Option Explicit
'@Folder("Modules")

Public Const PublicDir          As String = vbNullString
Public Const LocalDir           As String = "W:\App Development\Spec Manager"
Public Const SQLITE_PATH        As String = "S:\Data Manager\SAATI_Spec_Manager.db3"
Public Const GitBashExe         As String = "C:\Users\cruff\AppData\Local\Programs\Git\git-bash.exe"
Public Const GitRepo            As String = "C:\Users\cruff\source\DataManager"
Public Const SlitterPath        As String = "S:\Public\04 Division - Filtration\Slitter set up parameters"

Public Const ISPEC_WARPING      As Long = 1 ' warper spec identifer for ispec factory
Public Const ISPEC_STYLE        As Long = 2 ' fabric style spec identifier for ispec factory
Public Const ISPEC_SLITTER      As Long = 3 ' heat slitter spec identifier for ispec factory
Public Const ISPEC_ULTRASONIC   As Long = 4 ' ultra sonic welder spec identifier for ispec factory