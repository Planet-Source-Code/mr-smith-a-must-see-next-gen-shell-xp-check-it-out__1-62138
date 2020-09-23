Attribute VB_Name = "modDirs"
Option Explicit
    Public Const EWX_FORCE = 4
  
    Public Const EWX_REBOOT = 2
    Public Const EWX_SHUTDOWN = 1

Public Enum CSIDL
CSIDL_PROGRAMS = &H2 'yes!
CSIDL_PERSONAL = &H5
CSIDL_FAVORITES = &H6 'yea
CSIDL_STARTMENU = &HB 'yea

End Enum
Public Enum siCSIDL_VALUES
    CSIDL_FLAG_CREATE = &H8000 ' (Version 5.0)
    CSIDL_ADMINTOOLS = &H30 ' (Version 5.0)
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_APPDATA = &H1A ' (Version 4.71)
    CSIDL_BITBUCKET = &HA
    CSIDL_COMMON_ADMINTOOLS = &H2F  ' Version 5
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_APPDATA = &H23  ' Version 5
    CSIDL_COMMON_DOCUMENTS = &H2E
   
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_CONTROLS = &H3
    CSIDL_COOKIES = &H21
    CSIDL_desktop = &H0 'yes
    CSIDL_DESKTOPDIRECTORY = &H10 'yes
    CSIDL_DRIVES = &H11
    CSIDL_FONTS = &H14
    CSIDL_HISTORY = &H22
    CSIDL_INTERNET = &H1
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_LOCAL_APPDATA = &H1C      ' Version 5
    CSIDL_MYPICTURES = &H27  ' Version 5
    CSIDL_NETHOOD = &H13
    CSIDL_NETWORK = &H12
    CSIDL_PRINTERS = &H4
    CSIDL_PRINTHOOD = &H1B
    CSIDL_PROFILE = &H28  ' Version 5
    CSIDL_PROGRAM_FILES = &H2A  ' Version 5
    CSIDL_PROGRAM_FILES_COMMON = &H2B  ' Version 5

    CSIDL_RECENT = &H8 'yea!
    CSIDL_SENDTO = &H9 'yea!

    CSIDL_STARTUP = &H7 'sorta
    CSIDL_SYSTEM = &H25  ' Version 5
    CSIDL_TEMPLATES = &H15
    CSIDL_WINDOWS = &H24    ' Version 5.0.
End Enum
Private Enum SpecialFolderIDs
 desktop = &H0
 Programs = &H2
 Personal = &H5
 Favorites = &H6
 StartUp = &H7
 Recent = &H8
 SendTo = &H9
 StartMenu = &HB
 DesktopDirectory = &H10
 NetHood = &H13
 Fonts = &H14
 Templates = &H15
 Common_StartMenu = &H16
 Common_Programs = &H17
 Common_StartUp = &H18
 Common_DesktopDirectory = &H19
 AppData = &H1A
 PrintHood = &H1B
End Enum
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As SpecialFolderIDs, ByRef pIdl As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long


Public Declare Function SystemParametersInfoA Lib "user32" (ByVal a As Long, ByVal b As Long, ByVal C As Long, ByVal d As Long) As Long
Public Declare Function DeleteFileA Lib "kernel32" (ByVal a As String) As Long
Public Declare Function Sleep Lib "kernel32" (ByVal a As Long) As Long
Public Declare Function mciSendStringA Lib "winmm" (ByVal a As String, ByVal b As String, ByVal C As Long, ByVal d As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal a As Long, ByVal b As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal a As Long) As Long
Public Declare Function FindWindowA Lib "user32" (ByVal a As String, ByVal b As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal a As Long, ByVal b As Long) As Long
Public Declare Function FindWindowExA Lib "user32" (ByVal a As Long, ByVal b As Long, ByVal C As String, ByVal d As String) As Long
Private Const SW_SHOWNORMAL As Long = 1
Private Declare Function SHGetFolderPath Lib "SHFolder" Alias "SHGetFolderPathA" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, _
ByVal dwFlags As Long, ByVal sPath As String) As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hwnd As Long, ByVal Idk1 As Long, ByVal Idk2 As Long, ByVal dTitle As String, ByVal dPrompt As String, ByVal uFlags As Long) As Long

Declare Function FindExecutable _
    Lib "shell32.dll" _
    Alias "FindExecutableA" ( _
        ByVal lpFile As String, _
        ByVal lpDirectory As String, _
        ByVal lpResult As String) _
    As Long

Private Const S_OK = &H0

Public Function sDir(dir As CSIDL) As String

    Dim sPath As String * 255, L As Long
    Dim SW As Long

   sDir = ""

    SW = SHGetFolderPath(0, dir, 0&, 0&, sPath)
    
    If SW = S_OK Then
        L = InStr(sPath, Chr$(0))
        If (L > 0) And (L <= 255) Then sDir = Left$(sPath, L - 1)
    End If

End Function
Public Sub ShowRunDialog(Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal BrowseButton As Boolean = True, Optional ByVal OwnerFormhWnd As Long)
 Call SHRunDialog(OwnerFormhWnd, 0, 0, Title, Text, IIf(BrowseButton, 2, 1))
End Sub

Public Sub ShowFindDialog(Optional InitialDirectory As String)

ShellExecute 0, "find", _
  IIf(InitialDirectory = "", "", InitialDirectory), _
  vbNullString, vbNullString, SW_SHOW

End Sub
Public Function GetFolder(ByVal FolderID As Integer) As String
 Dim path As String, IDL As Long
 If SHGetSpecialFolderLocation(0, FolderID, IDL) = 0 Then
  path = String$(255, 0)
  Call SHGetPathFromIDListA(IDL, path)
  GetFolder = Left(path, InStr(1, path, Chr(0)) - 1)
 End If
End Function



Public Function ShowRecycleBin() As Boolean
      Dim lRet As Long
     'if using from a form, you can use me.hwnd instead of 0&
     'for the first argument
       lRet = ShellExecute(0&, "Open", "explorer.exe", _
       "/root,::{645FF040-5081-101B-9F08-00AA002F954E}", 0&, _
        SW_SHOWNORMAL)
        ShowRecycleBin = lRet > 32
End Function
