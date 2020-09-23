Attribute VB_Name = "modUser"
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetWinComputerName() As String
Dim sBuffer As String
Dim lBufSize As Long
Dim lStatus As Long
lBufSize = 255
sBuffer = String$(lBufSize, " ")
lStatus = GetComputerName(sBuffer, lBufSize)
GetWinComputerName = ""
If lStatus <> 0 Then
GetWinComputerName = Left(sBuffer, lBufSize)
End If
End Function

Function CurUserName$()
    Dim sTmp1$
    sTmp1 = Space$(512)
    GetUserName sTmp1, Len(sTmp1)
    CurUserName = Trim$(sTmp1)
End Function

