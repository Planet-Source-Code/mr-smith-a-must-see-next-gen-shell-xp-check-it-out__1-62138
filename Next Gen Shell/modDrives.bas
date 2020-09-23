Attribute VB_Name = "modDrives"
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_UNKNOWN = 0
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Function FindDrives() As String
Dim i As Integer
Dim drv As Long
Dim X As String
Dim r As String
drv = 65
Do Until drv = 91
X = Chr$(drv)
If X = "C" Then r = r & Chr$(drv) & 3
i = GetDriveType(X & ":")
DoEvents
Select Case i
Case 0
r = r & Chr$(drv) & 0
Case 2
r = r & Chr$(drv) & 2
Case 3
r = r & Chr$(drv) & 3
Case 4
r = r & Chr$(drv) & 4
Case 5
r = r & Chr$(drv) & 5
Case 6
r = r & Chr$(drv) & 6
End Select
drv = drv + 1
Loop
FindDrives = r
End Function
