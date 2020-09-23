Attribute VB_Name = "modDriveVol"
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" ( _
           ByVal lpRootPathName As String, _
           ByVal lpVolumeNameBuffer As String, _
           ByVal nVolumeNameSize As Long, _
           lpVolumeSerialNumber As Long, _
           lpMaximumComponentLength As Long, _
           lpFileSystemFlags As Long, _
           ByVal lpFileSystemNameBuffer As String, _
           ByVal nFileSystemNameSize As Long _
) As Long


Public Function GetDriveName(a As String) As String
Dim Serial As Long, VName As String, FSName As String
VName = String$(255, Chr$(0))
FSName = String$(255, Chr$(0))
GetVolumeInformation a & ":\", VName, 255, Serial, 0, 0, FSName, 255
VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
GetDriveName = VName
End Function
