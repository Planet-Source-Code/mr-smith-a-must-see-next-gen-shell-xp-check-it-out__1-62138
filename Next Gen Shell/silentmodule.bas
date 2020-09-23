Attribute VB_Name = "silent"
'This is a module I compiled for you peeps TO LEARN from, IT IS NOT one I use I made my own totally, but I compiled this for you guys to leard from. Some was coded by me and some was coded by Unsakred... Have Fun - i_silent_i
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function venkymd5crypt Lib "venky.dll" (ByVal pass As String, ByVal salt As String, ByVal Ret As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Option Explicit
      #If Win32 Then

      #Else
        Private Declare Function sndPlaySound Lib "MMSYSTEM" ( _
                           lpszSoundName As Any, ByVal uFlags%) As Integer
      #End If
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Const EM_UNDO = &HC7
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Public Const SND_MEMORY = &H4

Public Const WM_SETFOCUS = &H7
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101

Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_TAB = &H9
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = 1

Public Const SW_ERASE = &H4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_SEPARATOR = &H800&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const ENTA = 13
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const EM_LINESCROLL = &HB6
Private Const SPI_SCREENSAVERRUNNING = 97
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)




Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


' <VB WATCH>
Const VBWMODULE = "Module1"
' </VB WATCH>

Public Sub runmenu(lngwindow As Long, strmenutext As String)
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "Module1.runmenu"
3          If vbwTraceProc Then
4              Dim vbwParameterString As String
5              If vbwTraceParameters Then
6                  vbwParameterString = "(" & vbwReportParameter("lngwindow", lngwindow) & ", "
7                  vbwParameterString = vbwParameterString & vbwReportParameter("strmenutext", strmenutext) & ") "
8              End If
9              vbwTraceIn VBWPROCNAME, vbwParameterString
10         End If
' </VB WATCH>
11     Dim intLoop As Integer, intSubLoop As Integer, intSub2Loop As Integer, intSub3Loop As Integer, intSub4Loop As Integer
12     Dim lngmenu(1 To 5) As Long
13     Dim lngcount(1 To 5) As Long
14     Dim lngSubMenuID(1 To 4) As Long
15     Dim strcaption(1 To 4) As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "16         lngmenu(1) = GetMenu(lngwindow&)"
' </VB WATCH>
16         lngmenu(1) = GetMenu(lngwindow&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "17         lngcount(1) = GetMenuItemCount(lngmenu(1))"
' </VB WATCH>
17         lngcount(1) = GetMenuItemCount(lngmenu(1))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "18             For intLoop% = 0 To lngcount(1) - 1"
' </VB WATCH>
18             For intLoop% = 0 To lngcount(1) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "19                 DoEvents"
' </VB WATCH>
19                 DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "20                 lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)"
' </VB WATCH>
20                 lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "21                 lngcount(2) = GetMenuItemCount(lngmenu(2))"
' </VB WATCH>
21                 lngcount(2) = GetMenuItemCount(lngmenu(2))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "22                     For intSubLoop% = 0 To lngcount(2) - 1"
' </VB WATCH>
22                     For intSubLoop% = 0 To lngcount(2) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "23                         DoEvents"
' </VB WATCH>
23                         DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "24                         lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)"
' </VB WATCH>
24                         lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "25                         strcaption(1) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
25                         strcaption(1) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "26                         Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)"
' </VB WATCH>
26                         Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "27                             If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then"
' </VB WATCH>
27                             If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "28                                 Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)"
' </VB WATCH>
28                                 Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)
' <VB WATCH>
29         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "30                                 Exit Sub"
' </VB WATCH>
30                                 Exit Sub
31                             End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "31                             End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "32                         lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)"
' </VB WATCH>
32                         lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "33                         lngcount(3) = GetMenuItemCount(lngmenu(3))"
' </VB WATCH>
33                         lngcount(3) = GetMenuItemCount(lngmenu(3))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "34                             If lngcount(3) > 0 Then"
' </VB WATCH>
34                             If lngcount(3) > 0 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "35                                 For intSub2Loop% = 0 To lngcount(3) - 1"
' </VB WATCH>
35                                 For intSub2Loop% = 0 To lngcount(3) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "36                                     DoEvents"
' </VB WATCH>
36                                     DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "37                                     lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
37                                     lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "38                                     strcaption(2) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
38                                     strcaption(2) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "39                                     Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)"
' </VB WATCH>
39                                     Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "40                                         If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then"
' </VB WATCH>
40                                         If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "41                                             Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)"
' </VB WATCH>
41                                             Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)
' <VB WATCH>
42         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "43                                             Exit Sub"
' </VB WATCH>
43                                             Exit Sub
44                                         End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "44                                         End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "45                                     lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
45                                     lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "46                                     lngcount(4) = GetMenuItemCount(lngmenu(4))"
' </VB WATCH>
46                                     lngcount(4) = GetMenuItemCount(lngmenu(4))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "47                                         If lngcount(4) > 0 Then"
' </VB WATCH>
47                                         If lngcount(4) > 0 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "48                                             For intSub3Loop% = 0 To lngcount(4) - 1"
' </VB WATCH>
48                                             For intSub3Loop% = 0 To lngcount(4) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "49                                                 DoEvents"
' </VB WATCH>
49                                                 DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "50                                                 lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
50                                                 lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "51                                                 strcaption(3) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
51                                                 strcaption(3) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "52                                                 Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)"
' </VB WATCH>
52                                                 Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "53                                                     If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then"
' </VB WATCH>
53                                                     If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "54                                                         Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)"
' </VB WATCH>
54                                                         Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)
' <VB WATCH>
55         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "56                                                         Exit Sub"
' </VB WATCH>
56                                                         Exit Sub
57                                                     End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "57                                                     End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "58                                                 lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
58                                                 lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "59                                                 lngcount(5) = GetMenuItemCount(lngmenu(5))"
' </VB WATCH>
59                                                 lngcount(5) = GetMenuItemCount(lngmenu(5))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "60                                                     If lngcount(5) > 0 Then"
' </VB WATCH>
60                                                     If lngcount(5) > 0 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "61                                                         For intSub4Loop% = 0 To lngcount(5) - 1"
' </VB WATCH>
61                                                         For intSub4Loop% = 0 To lngcount(5) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "62                                                             DoEvents"
' </VB WATCH>
62                                                             DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "63                                                             lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)"
' </VB WATCH>
63                                                             lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "64                                                             strcaption(4) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
64                                                             strcaption(4) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "65                                                             Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)"
' </VB WATCH>
65                                                             Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "66                                                                 If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then"
' </VB WATCH>
66                                                                 If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "67                                                                     Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)"
' </VB WATCH>
67                                                                     Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)
' <VB WATCH>
68         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "69                                                                     Exit Sub"
' </VB WATCH>
69                                                                     Exit Sub
70                                                                 End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "70                                                                 End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "71                                                         Next intSub4Loop%"
' </VB WATCH>
71                                                         Next intSub4Loop%
72                                                     End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "72                                                     End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "73                                             Next intSub3Loop%"
' </VB WATCH>
73                                             Next intSub3Loop%
74                                         End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "74                                         End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "75                                 Next intSub2Loop%"
' </VB WATCH>
75                                 Next intSub2Loop%
76                             End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "76                             End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "77                     Next intSubLoop%"
' </VB WATCH>
77                     Next intSubLoop%
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "78             Next intLoop%"
' </VB WATCH>
78             Next intLoop%
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
79         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
80         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "runmenu"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub mypmboot()
' <VB WATCH>
81         On Error GoTo vbwErrHandler
82         Const VBWPROCNAME = "Module1.mypmboot"
83         If vbwTraceProc Then
84             Dim vbwParameterString As String
85             If vbwTraceParameters Then
86                 vbwParameterString = "()"
87             End If
88             vbwTraceIn VBWPROCNAME, vbwParameterString
89         End If
' </VB WATCH>
90     Dim imclass As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "91     imclass = FindWindow(" & Chr(34) & "IMCLASS" & Chr(34) & ", vbNullString)"
' </VB WATCH>
91     imclass = FindWindow("IMCLASS", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "92     Call runmenu(imclass&, " & Chr(34) & "file &Send" & Chr(34) & ")"
' </VB WATCH>
92     Call runmenu(imclass&, "file &Send")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
93         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
94         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "mypmboot"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub Pause(interval)
' <VB WATCH>
95         On Error GoTo vbwErrHandler
96         Const VBWPROCNAME = "Module1.pause"
97         If vbwTraceProc Then
98             Dim vbwParameterString As String
99             If vbwTraceParameters Then
100                vbwParameterString = "(" & vbwReportParameter("interval", interval) & ") "
101            End If
102            vbwTraceIn VBWPROCNAME, vbwParameterString
103        End If
' </VB WATCH>
104    Dim Current
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "105    Current = Timer"
' </VB WATCH>
105    Current = Timer
' <VBW_LINE>Do While Timer - Current < Val(interval)
106    Do While vbwExecuteLine(False, "106    Do While Timer - Current < Val(interval)") Or _
        Timer - Current < Val(interval)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "107    DoEvents"
' </VB WATCH>
107    DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "108    Loop"
' </VB WATCH>
108    Loop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
109        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
110        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "pause"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Sub WindowClose(lngHwnd As Long)
' <VB WATCH>
111        On Error GoTo vbwErrHandler
112        Const VBWPROCNAME = "Module1.WindowClose"
113        If vbwTraceProc Then
114            Dim vbwParameterString As String
115            If vbwTraceParameters Then
116                vbwParameterString = "(" & vbwReportParameter("lngHwnd", lngHwnd) & ") "
117            End If
118            vbwTraceIn VBWPROCNAME, vbwParameterString
119        End If
' </VB WATCH>
120      Dim imclass As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "121    imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
121    imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "122    Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
122    Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
123        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
124        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WindowClose"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Sub WindowHide(lngHwnd As Long)
' <VB WATCH>
125        On Error GoTo vbwErrHandler
126        Const VBWPROCNAME = "Module1.WindowHide"
127        If vbwTraceProc Then
128            Dim vbwParameterString As String
129            If vbwTraceParameters Then
130                vbwParameterString = "(" & vbwReportParameter("lngHwnd", lngHwnd) & ") "
131            End If
132            vbwTraceIn VBWPROCNAME, vbwParameterString
133        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "134        Call ShowWindow(lngHwnd&, SW_HIDE)"
' </VB WATCH>
134        Call ShowWindow(lngHwnd&, SW_HIDE)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
135        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
136        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WindowHide"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Sub WindowShow(lngHwnd As Long)
' <VB WATCH>
137        On Error GoTo vbwErrHandler
138        Const VBWPROCNAME = "Module1.WindowShow"
139        If vbwTraceProc Then
140            Dim vbwParameterString As String
141            If vbwTraceParameters Then
142                vbwParameterString = "(" & vbwReportParameter("lngHwnd", lngHwnd) & ") "
143            End If
144            vbwTraceIn VBWPROCNAME, vbwParameterString
145        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "146        Call ShowWindow(lngHwnd&, SW_SHOW)"
' </VB WATCH>
146        Call ShowWindow(lngHwnd&, SW_SHOW)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
147        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
148        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WindowShow"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub Buddy()
' <VB WATCH>
149        On Error GoTo vbwErrHandler
150        Const VBWPROCNAME = "Module1.Buddy"
151        If vbwTraceProc Then
152            Dim vbwParameterString As String
153            If vbwTraceParameters Then
154                vbwParameterString = "()"
155            End If
156            vbwTraceIn VBWPROCNAME, vbwParameterString
157        End If
' </VB WATCH>
158    Dim imclass As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "159    imclass = FindWindow(" & Chr(34) & "IMCLASS" & Chr(34) & ", vbNullString)"
' </VB WATCH>
159    imclass = FindWindow("IMCLASS", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "160    Call runmenu(imclass&, " & Chr(34) & "&Add as Friend" & Chr(34) & ")"
' </VB WATCH>
160    Call runmenu(imclass&, "&Add as Friend")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
161        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
162        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Buddy"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub ClickMenu(lngwindow As Long, strmenutext As String)
       'This is from Andymaul one of my closest friends
       'Thanks man.
' <VB WATCH>
163        On Error GoTo vbwErrHandler
164        Const VBWPROCNAME = "Module1.ClickMenu"
165        If vbwTraceProc Then
166            Dim vbwParameterString As String
167            If vbwTraceParameters Then
168                vbwParameterString = "(" & vbwReportParameter("lngwindow", lngwindow) & ", "
169                vbwParameterString = vbwParameterString & vbwReportParameter("strmenutext", strmenutext) & ") "
170            End If
171            vbwTraceIn VBWPROCNAME, vbwParameterString
172        End If
' </VB WATCH>
173    Dim intLoop As Integer, intSubLoop As Integer, intSub2Loop As Integer, intSub3Loop As Integer, intSub4Loop As Integer
174    Dim lngmenu(1 To 5) As Long
175    Dim lngcount(1 To 5) As Long
176    Dim lngSubMenuID(1 To 4) As Long
177    Dim strcaption(1 To 4) As String

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "178        lngmenu(1) = GetMenu(lngwindow&)"
' </VB WATCH>
178        lngmenu(1) = GetMenu(lngwindow&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "179        lngcount(1) = GetMenuItemCount(lngmenu(1))"
' </VB WATCH>
179        lngcount(1) = GetMenuItemCount(lngmenu(1))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "180            For intLoop% = 0 To lngcount(1) - 1"
' </VB WATCH>
180            For intLoop% = 0 To lngcount(1) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "181                DoEvents"
' </VB WATCH>
181                DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "182                lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)"
' </VB WATCH>
182                lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "183                lngcount(2) = GetMenuItemCount(lngmenu(2))"
' </VB WATCH>
183                lngcount(2) = GetMenuItemCount(lngmenu(2))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "184                    For intSubLoop% = 0 To lngcount(2) - 1"
' </VB WATCH>
184                    For intSubLoop% = 0 To lngcount(2) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "185                        DoEvents"
' </VB WATCH>
185                        DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "186                        lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)"
' </VB WATCH>
186                        lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "187                        strcaption(1) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
187                        strcaption(1) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "188                        Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)"
' </VB WATCH>
188                        Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "189                            If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then"
' </VB WATCH>
189                            If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "190                                Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)"
' </VB WATCH>
190                                Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)

' <VB WATCH>
191        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "192                                Exit Sub"
' </VB WATCH>
192                                Exit Sub

193                            End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "193                            End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "194                        lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)"
' </VB WATCH>
194                        lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "195                        lngcount(3) = GetMenuItemCount(lngmenu(3))"
' </VB WATCH>
195                        lngcount(3) = GetMenuItemCount(lngmenu(3))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "196                            If lngcount(3) > 0 Then"
' </VB WATCH>
196                            If lngcount(3) > 0 Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "197                                For intSub2Loop% = 0 To lngcount(3) - 1"
' </VB WATCH>
197                                For intSub2Loop% = 0 To lngcount(3) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "198                                    DoEvents"
' </VB WATCH>
198                                    DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "199                                    lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
199                                    lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "200                                    strcaption(2) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
200                                    strcaption(2) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "201                                    Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)"
' </VB WATCH>
201                                    Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "202                                        If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then"
' </VB WATCH>
202                                        If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "203                                            Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)"
' </VB WATCH>
203                                            Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)

' <VB WATCH>
204        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "205                                            Exit Sub"
' </VB WATCH>
205                                            Exit Sub

206                                        End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "206                                        End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "207                                    lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
207                                    lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "208                                    lngcount(4) = GetMenuItemCount(lngmenu(4))"
' </VB WATCH>
208                                    lngcount(4) = GetMenuItemCount(lngmenu(4))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "209                                        If lngcount(4) > 0 Then"
' </VB WATCH>
209                                        If lngcount(4) > 0 Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "210                                            For intSub3Loop% = 0 To lngcount(4) - 1"
' </VB WATCH>
210                                            For intSub3Loop% = 0 To lngcount(4) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "211                                                DoEvents"
' </VB WATCH>
211                                                DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "212                                                lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
212                                                lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "213                                                strcaption(3) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
213                                                strcaption(3) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "214                                                Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)"
' </VB WATCH>
214                                                Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "215                                                    If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then"
' </VB WATCH>
215                                                    If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "216                                                        Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)"
' </VB WATCH>
216                                                        Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)

' <VB WATCH>
217        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "218                                                        Exit Sub"
' </VB WATCH>
218                                                        Exit Sub

219                                                    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "219                                                    End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "220                                                lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
220                                                lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "221                                                lngcount(5) = GetMenuItemCount(lngmenu(5))"
' </VB WATCH>
221                                                lngcount(5) = GetMenuItemCount(lngmenu(5))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "222                                                    If lngcount(5) > 0 Then"
' </VB WATCH>
222                                                    If lngcount(5) > 0 Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "223                                                        For intSub4Loop% = 0 To lngcount(5) - 1"
' </VB WATCH>
223                                                        For intSub4Loop% = 0 To lngcount(5) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "224                                                            DoEvents"
' </VB WATCH>
224                                                            DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "225                                                            lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)"
' </VB WATCH>
225                                                            lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "226                                                            strcaption(4) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
226                                                            strcaption(4) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "227                                                            Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)"
' </VB WATCH>
227                                                            Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "228                                                                If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then"
' </VB WATCH>
228                                                                If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "229                                                                    Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)"
' </VB WATCH>
229                                                                    Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)

' <VB WATCH>
230        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "231                                                                    Exit Sub"
' </VB WATCH>
231                                                                    Exit Sub

232                                                                End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "232                                                                End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "233                                                        Next intSub4Loop%"
' </VB WATCH>
233                                                        Next intSub4Loop%

234                                                    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "234                                                    End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "235                                            Next intSub3Loop%"
' </VB WATCH>
235                                            Next intSub3Loop%

236                                        End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "236                                        End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "237                                Next intSub2Loop%"
' </VB WATCH>
237                                Next intSub2Loop%

238                            End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "238                            End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "239                    Next intSubLoop%"
' </VB WATCH>
239                    Next intSubLoop%

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "240            Next intLoop%"
' </VB WATCH>
240            Next intLoop%

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
241        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
242        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ClickMenu"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Function GetCaption(hwnd)
' <VB WATCH>
243        On Error GoTo vbwErrHandler
244        Const VBWPROCNAME = "Module1.GetCaption"
245        If vbwTraceProc Then
246            Dim vbwParameterString As String
247            If vbwTraceParameters Then
248                vbwParameterString = "(" & vbwReportParameter("hwnd", hwnd) & ") "
249            End If
250            vbwTraceIn VBWPROCNAME, vbwParameterString
251        End If
' </VB WATCH>
252    Dim hWndlength As Integer, hWndTitle As String, a As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "253    hWndlength% = GetWindowTextLength(hwnd)"
' </VB WATCH>
253    hWndlength% = GetWindowTextLength(hwnd)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "254    hWndTitle$ = String$(hWndlength%, 0)"
' </VB WATCH>
254    hWndTitle$ = String$(hWndlength%, 0)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "255    a% = GetWindowText(hwnd, hWndTitle$, (hWndlength% + 1))"
' </VB WATCH>
255    a% = GetWindowText(hwnd, hWndTitle$, (hWndlength% + 1))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "256    GetCaption = hWndTitle$"
' </VB WATCH>
256    GetCaption = hWndTitle$
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
257        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
258        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetCaption"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function FindPMWnd()
' <VB WATCH>
259        On Error GoTo vbwErrHandler
260        Const VBWPROCNAME = "Module1.FindPMWnd"
261        If vbwTraceProc Then
262            Dim vbwParameterString As String
263            If vbwTraceParameters Then
264                vbwParameterString = "()"
265            End If
266            vbwTraceIn VBWPROCNAME, vbwParameterString
267        End If
' </VB WATCH>
268    Dim imclass As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "269    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
269    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "270    If InStr(GetCaption(imclass&), LCase(" & Chr(34) & " -- instant message" & Chr(34) & ")) Then"
' </VB WATCH>
270    If InStr(GetCaption(imclass&), LCase(" -- instant message")) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "271"
' </VB WATCH>
271
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "272        FindPMWnd = imclass&"
' </VB WATCH>
272        FindPMWnd = imclass&
273    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "273    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
274        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
275        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindPMWnd"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function FindMainWnd()
' <VB WATCH>
276        On Error GoTo vbwErrHandler
277        Const VBWPROCNAME = "Module1.FindMainWnd"
278        If vbwTraceProc Then
279            Dim vbwParameterString As String
280            If vbwTraceParameters Then
281                vbwParameterString = "()"
282            End If
283            vbwTraceIn VBWPROCNAME, vbwParameterString
284        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "285    FindMainWnd = FindWindow(" & Chr(34) & "Yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
285    FindMainWnd = FindWindow("Yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
286        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
287        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindMainWnd"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function FindChatWnd()
' <VB WATCH>
288        On Error GoTo vbwErrHandler
289        Const VBWPROCNAME = "Module1.FindChatWnd"
290        If vbwTraceProc Then
291            Dim vbwParameterString As String
292            If vbwTraceParameters Then
293                vbwParameterString = "()"
294            End If
295            vbwTraceIn VBWPROCNAME, vbwParameterString
296        End If
' </VB WATCH>
297    Dim imclass As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "298    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
298    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "299    If InStr(GetCaption(imclass&), LCase(" & Chr(34) & " -- chat" & Chr(34) & ")) Then"
' </VB WATCH>
299    If InStr(GetCaption(imclass&), LCase(" -- chat")) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "300"
' </VB WATCH>
300
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "301        FindChatWnd = imclass&"
' </VB WATCH>
301        FindChatWnd = imclass&
302    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "302    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
303        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
304        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindChatWnd"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function ClickButton(Button As Long)
' <VB WATCH>
305        On Error GoTo vbwErrHandler
306        Const VBWPROCNAME = "Module1.ClickButton"
307        If vbwTraceProc Then
308            Dim vbwParameterString As String
309            If vbwTraceParameters Then
310                vbwParameterString = "(" & vbwReportParameter("Button", Button) & ") "
311            End If
312            vbwTraceIn VBWPROCNAME, vbwParameterString
313        End If
' </VB WATCH>
314    Dim Click As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "315    Click& = SendMessageByNum(Button, WM_LBUTTONDOWN, &HD, 0)"
' </VB WATCH>
315    Click& = SendMessageByNum(Button, WM_LBUTTONDOWN, &HD, 0)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "316    Click& = SendMessageByNum(Button, WM_LBUTTONUP, &HD, 0)"
' </VB WATCH>
316    Click& = SendMessageByNum(Button, WM_LBUTTONUP, &HD, 0)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
317        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
318        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ClickButton"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Sub WindowDisable(Window As Long)
' <VB WATCH>
319        On Error GoTo vbwErrHandler
320        Const VBWPROCNAME = "Module1.WindowDisable"
321        If vbwTraceProc Then
322            Dim vbwParameterString As String
323            If vbwTraceParameters Then
324                vbwParameterString = "(" & vbwReportParameter("Window", Window) & ") "
325            End If
326            vbwTraceIn VBWPROCNAME, vbwParameterString
327        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "328    Call EnableWindow(Window&, 0)"
' </VB WATCH>
328    Call EnableWindow(Window&, 0)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
329        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
330        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WindowDisable"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub WindowEnable(Window As Long)
' <VB WATCH>
331        On Error GoTo vbwErrHandler
332        Const VBWPROCNAME = "Module1.WindowEnable"
333        If vbwTraceProc Then
334            Dim vbwParameterString As String
335            If vbwTraceParameters Then
336                vbwParameterString = "(" & vbwReportParameter("Window", Window) & ") "
337            End If
338            vbwTraceIn VBWPROCNAME, vbwParameterString
339        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "340    Call EnableWindow(Window&, 1)"
' </VB WATCH>
340    Call EnableWindow(Window&, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
341        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
342        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WindowEnable"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Sub SendText(what$)
' <VB WATCH>
343        On Error GoTo vbwErrHandler
344        Const VBWPROCNAME = "Module1.SendText"
345        If vbwTraceProc Then
346            Dim vbwParameterString As String
347            If vbwTraceParameters Then
348                vbwParameterString = "(" & vbwReportParameter("what$", what$) & ") "
349            End If
350            vbwTraceIn VBWPROCNAME, vbwParameterString
351        End If
' </VB WATCH>
352    Dim imc As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "353    imc& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
353    imc& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "354    RichEdit& = FindWindowEx(imc&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
354    RICHEDIT& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "355    Call SendMessageByString(RichEdit, WM_SETTEXT, 0&, what$)"
' </VB WATCH>
355    Call SendMessageByString(RICHEDIT, WM_SETTEXT, 0&, what$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "356    Call pause(0.2)"
' </VB WATCH>
356    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "357    Call ClickMenu(imc&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
357    Call ClickMenu(imc&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
358        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
359        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendText"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SendBoot(code$)
' <VB WATCH>
360        On Error GoTo vbwErrHandler
361        Const VBWPROCNAME = "Module1.SendBoot"
362        If vbwTraceProc Then
363            Dim vbwParameterString As String
364            If vbwTraceParameters Then
365                vbwParameterString = "(" & vbwReportParameter("Code$", code$) & ") "
366            End If
367            vbwTraceIn VBWPROCNAME, vbwParameterString
368        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "369    Anti2"
' </VB WATCH>
369    Anti2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "370    ClosedaWindow"
' </VB WATCH>
370    ClosedaWindow
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "371    Closeewindow"
' </VB WATCH>
371    Closeewindow
372    Dim parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "373    parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
373    parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "374    If InStr(GetCaption(parent&), LCase(" & Chr(34) & "-- instant message" & Chr(34) & ")) Then"
' </VB WATCH>
374    If InStr(GetCaption(parent&), LCase("-- instant message")) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "375"
' </VB WATCH>
375
' <VB WATCH>
376        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "377        Exit Sub"
' </VB WATCH>
377        Exit Sub
378    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "378    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "379    Call SetFocusApi(parent&)"
' </VB WATCH>
379    Call SetFocusApi(parent&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "380    Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
380    Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "381    Call SendMessageByString(Child2&, WM_SETTEXT, 0, Code$)"
' </VB WATCH>
381    Call SendMessageByString(Child2&, WM_SETTEXT, 0, code$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "382    Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
382    Call ClickMenu(parent&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
383        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
384        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendBoot"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SendTextScroll(what As String, Times As Integer)
' <VB WATCH>
385        On Error GoTo vbwErrHandler
386        Const VBWPROCNAME = "Module1.SendTextScroll"
387        If vbwTraceProc Then
388            Dim vbwParameterString As String
389            If vbwTraceParameters Then
390                vbwParameterString = "(" & vbwReportParameter("what", what) & ", "
391                vbwParameterString = vbwParameterString & vbwReportParameter("Times", Times) & ") "
392            End If
393            vbwTraceIn VBWPROCNAME, vbwParameterString
394        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "395    Do"
' </VB WATCH>
395    Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "396    SendText (what$)"
' </VB WATCH>
396    SendText (what$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "397    Times = Times% - 1"
' </VB WATCH>
397    Times = Times% - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "398    Call pause(0.3)"
' </VB WATCH>
398    Call Pause(0.3)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "399    Loop Until Times% = 0"
' </VB WATCH>
399    Loop Until Times% = 0
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
400        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
401        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendTextScroll"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendPM(who$, what$, Follow As Boolean)
' <VB WATCH>
402        On Error GoTo vbwErrHandler
403        Const VBWPROCNAME = "Module1.SendPM"
404        If vbwTraceProc Then
405            Dim vbwParameterString As String
406            If vbwTraceParameters Then
407                vbwParameterString = "(" & vbwReportParameter("Who$", who$) & ", "
408                vbwParameterString = vbwParameterString & vbwReportParameter("what$", what$) & ", "
409                vbwParameterString = vbwParameterString & vbwReportParameter("Follow", Follow) & ") "
410            End If
411            vbwTraceIn VBWPROCNAME, vbwParameterString
412        End If
' </VB WATCH>
413    Dim yahoobuddymain As Long, parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "414    yahoobuddymain = FindWindow(" & Chr(34) & "yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
414    yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "415    Call ClickMenu(yahoobuddymain&, " & Chr(34) & "Send a &Message" & Chr(34) & ")"
' </VB WATCH>
415    Call ClickMenu(yahoobuddymain&, "Send a &Message")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "416    Call pause(0.2)"
' </VB WATCH>
416    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "417    parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
417    parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "418    If InStr(GetCaption(parent&), " & Chr(34) & "Chat" & Chr(34) & ") Then"
' </VB WATCH>
418    If InStr(GetCaption(parent&), "Chat") Then
' <VB WATCH>
419        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "420         Exit Sub"
' </VB WATCH>
420         Exit Sub
421    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "421    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "422    Child1& = FindWindowEx(parent&, 0&, " & Chr(34) & "Edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
422    Child1& = FindWindowEx(parent&, 0&, "Edit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "423    Call SetFocusApi(Child1&)"
' </VB WATCH>
423    Call SetFocusApi(Child1&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "424    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, Who$)"
' </VB WATCH>
424    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, who$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "425    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)"
' </VB WATCH>
425    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "426    Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
426    Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "427    Call SendMessageByString(Child2&, WM_SETTEXT, 0&, what$)"
' </VB WATCH>
427    Call SendMessageByString(Child2&, WM_SETTEXT, 0&, what$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "428    Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
428    Call ClickMenu(parent&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "429    If Follow = False Then"
' </VB WATCH>
429    If Follow = False Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "430"
' </VB WATCH>
430
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "431        Call WindowClose(parent&)"
' </VB WATCH>
431        Call WindowClose(parent&)
432    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "432    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
433        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
434        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendPM"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SendChat(what$)
' <VB WATCH>
435        On Error GoTo vbwErrHandler
436        Const VBWPROCNAME = "Module1.SendChat"
437        If vbwTraceProc Then
438            Dim vbwParameterString As String
439            If vbwTraceParameters Then
440                vbwParameterString = "(" & vbwReportParameter("what$", what$) & ") "
441            End If
442            vbwTraceIn VBWPROCNAME, vbwParameterString
443        End If
' </VB WATCH>
444    Dim parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "445    parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
445    parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "446    If InStr(GetCaption(parent&), LCase(" & Chr(34) & "-- instant message" & Chr(34) & ")) Then"
' </VB WATCH>
446    If InStr(GetCaption(parent&), LCase("-- instant message")) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "447"
' </VB WATCH>
447
' <VB WATCH>
448        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "449        Exit Sub"
' </VB WATCH>
449        Exit Sub
450    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "450    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "451    Call SetFocusApi(parent&)"
' </VB WATCH>
451    Call SetFocusApi(parent&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "452    Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
452    Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "453    Call SendMessageByString(Child2&, WM_SETTEXT, 0, what$)"
' </VB WATCH>
453    Call SendMessageByString(Child2&, WM_SETTEXT, 0, what$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "454    Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
454    Call ClickMenu(parent&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
455        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
456        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendChat"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SendChatBoot(code$, ANTI As Boolean, StayIn As Boolean)
' <VB WATCH>
457        On Error GoTo vbwErrHandler
458        Const VBWPROCNAME = "Module1.SendChatBoot"
459        If vbwTraceProc Then
460            Dim vbwParameterString As String
461            If vbwTraceParameters Then
462                vbwParameterString = "(" & vbwReportParameter("Code$", code$) & ", "
463                vbwParameterString = vbwParameterString & vbwReportParameter("anti", ANTI) & ", "
464                vbwParameterString = vbwParameterString & vbwReportParameter("StayIn", StayIn) & ") "
465            End If
466            vbwTraceIn VBWPROCNAME, vbwParameterString
467        End If
' </VB WATCH>
468    Dim imc As Long, Rich As Long, Button As Long
469    Dim RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "470    imc& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
470    imc& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "471    If InStr(GetCaption(imc&), " & Chr(34) & "Chat" & Chr(34) & ") Then"
' </VB WATCH>
471    If InStr(GetCaption(imc&), "Chat") Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "472"
' </VB WATCH>
472
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "473        SetFocusApi (imc&)"
' </VB WATCH>
473        SetFocusApi (imc&)
474    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "474    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "475    imc& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
475    imc& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "476    RichEdit& = FindWindowEx(imc&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
476    RICHEDIT& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "477    RichEdit& = FindWindowEx(imc&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
477    RICHEDIT& = FindWindowEx(imc&, RICHEDIT&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "478    If anti = True Then"
' </VB WATCH>
478    If ANTI = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "479"
' </VB WATCH>
479
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "480        Call PostMessage(RichEdit&, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
480        Call PostMessage(RICHEDIT&, WM_CLOSE, 0&, 0&)
481    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "481    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "482    Rich& = FindWindowEx(imc&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
482    Rich& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "483    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Code$)"
' </VB WATCH>
483    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, code$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "484    Call pause(0.2)"
' </VB WATCH>
484    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "485    Call ClickMenu(imc&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
485    Call ClickMenu(imc&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "486    Call pause(0.2)"
' </VB WATCH>
486    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "487    If StayIn = False Then"
' </VB WATCH>
487    If StayIn = False Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "488"
' </VB WATCH>
488
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "489        WindowClose (imc&)"
' </VB WATCH>
489        WindowClose (imc&)
490    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "490    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
491        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
492        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendChatBoot"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendPMBoot(UserName$, code$, ANTI As Boolean, Follow As Boolean)
' <VB WATCH>
493        On Error GoTo vbwErrHandler
494        Const VBWPROCNAME = "Module1.SendPMBoot"
495        If vbwTraceProc Then
496            Dim vbwParameterString As String
497            If vbwTraceParameters Then
498                vbwParameterString = "(" & vbwReportParameter("UserName$", UserName$) & ", "
499                vbwParameterString = vbwParameterString & vbwReportParameter("Code$", code$) & ", "
500                vbwParameterString = vbwParameterString & vbwReportParameter("anti", ANTI) & ", "
501                vbwParameterString = vbwParameterString & vbwReportParameter("Follow", Follow) & ") "
502            End If
503            vbwTraceIn VBWPROCNAME, vbwParameterString
504        End If
' </VB WATCH>
505    Dim parent As Long, Child1 As Long, Child2 As Long, Button As Long
506    Dim yahoo As Long
507    Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "508    Yahoo& = FindWindow(" & Chr(34) & "YahooBuddyMain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
508    yahoo& = FindWindow("YahooBuddyMain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "509    Call ClickMenu(Yahoo&, " & Chr(34) & "Send a &Message" & Chr(34) & ")"
' </VB WATCH>
509    Call ClickMenu(yahoo&, "Send a &Message")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "510    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
510    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "511    If InStr(GetCaption(imclass&), " & Chr(34) & "Chat" & Chr(34) & ") Then"
' </VB WATCH>
511    If InStr(GetCaption(imclass&), "Chat") Then
' <VB WATCH>
512        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "513         Exit Sub"
' </VB WATCH>
513         Exit Sub
514    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "514    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "515    RichEdit& = FindWindowEx(imclass&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
515    RICHEDIT& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "516    RichEdit& = FindWindowEx(imclass&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
516    RICHEDIT& = FindWindowEx(imclass&, RICHEDIT&, "RICHEDIT", vbNullString)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "517    If anti = True Then"
' </VB WATCH>
517    If ANTI = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "518"
' </VB WATCH>
518
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "519        Call PostMessage(RichEdit&, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
519        Call PostMessage(RICHEDIT&, WM_CLOSE, 0&, 0&)
520    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "520    End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "521    parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
521    parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "522    RichEdit& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
522    RICHEDIT& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "523    RichEdit& = FindWindowEx(parent&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
523    RICHEDIT& = FindWindowEx(parent&, RICHEDIT&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "524    Child1& = FindWindowEx(parent&, 0&, " & Chr(34) & "Edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
524    Child1& = FindWindowEx(parent&, 0&, "Edit", vbNullString)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "525    Call SetFocusApi(Child1&)"
' </VB WATCH>
525    Call SetFocusApi(Child1&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "526    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, UserName$)"
' </VB WATCH>
526    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, UserName$)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "527    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)"
' </VB WATCH>
527    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "528    Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
528    Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)


' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "529    Call SetFocusApi(Child2&)"
' </VB WATCH>
529    Call SetFocusApi(Child2&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "530    Call SendMessageByString(Child2&, WM_SETTEXT, 0&, Code$)"
' </VB WATCH>
530    Call SendMessageByString(Child2&, WM_SETTEXT, 0&, code$)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "531    Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
531    Call ClickMenu(parent&, "Sen&d")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "532    Call pause(0.3)"
' </VB WATCH>
532    Call Pause(0.3)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "533    If Follow = False Then"
' </VB WATCH>
533    If Follow = False Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "534"
' </VB WATCH>
534
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "535        WindowClose (parent&)"
' </VB WATCH>
535        WindowClose (parent&)
536    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "536    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
537        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
538        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendPMBoot"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SendPMLagg(UserName$, code$, ANTI As Boolean, Times As Integer)
' <VB WATCH>
539        On Error GoTo vbwErrHandler
540        Const VBWPROCNAME = "Module1.SendPMLagg"
541        If vbwTraceProc Then
542            Dim vbwParameterString As String
543            If vbwTraceParameters Then
544                vbwParameterString = "(" & vbwReportParameter("UserName$", UserName$) & ", "
545                vbwParameterString = vbwParameterString & vbwReportParameter("Code$", code$) & ", "
546                vbwParameterString = vbwParameterString & vbwReportParameter("anti", ANTI) & ", "
547                vbwParameterString = vbwParameterString & vbwReportParameter("Times", Times) & ") "
548            End If
549            vbwTraceIn VBWPROCNAME, vbwParameterString
550        End If
' </VB WATCH>
551    Dim parent As Long, Child1 As Long, Child2 As Long, Button As Long
552    Dim yahoo As Long
553    Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "554    Yahoo& = FindWindow(" & Chr(34) & "YahooBuddyMain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
554    yahoo& = FindWindow("YahooBuddyMain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "555    Call ClickMenu(Yahoo&, " & Chr(34) & "Send a &Message" & Chr(34) & ")"
' </VB WATCH>
555    Call ClickMenu(yahoo&, "Send a &Message")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "556    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
556    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "557    If InStr(GetCaption(imclass&), " & Chr(34) & "Chat" & Chr(34) & ") Then"
' </VB WATCH>
557    If InStr(GetCaption(imclass&), "Chat") Then
' <VB WATCH>
558        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "559         Exit Sub"
' </VB WATCH>
559         Exit Sub
560    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "560    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "561    RichEdit& = FindWindowEx(imclass&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
561    RICHEDIT& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "562    RichEdit& = FindWindowEx(imclass&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
562    RICHEDIT& = FindWindowEx(imclass&, RICHEDIT&, "RICHEDIT", vbNullString)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "563    If anti = True Then"
' </VB WATCH>
563    If ANTI = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "564"
' </VB WATCH>
564
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "565        Call PostMessage(RichEdit&, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
565        Call PostMessage(RICHEDIT&, WM_CLOSE, 0&, 0&)
566    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "566    End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "567    parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
567    parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "568    RichEdit& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
568    RICHEDIT& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "569    RichEdit& = FindWindowEx(parent&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
569    RICHEDIT& = FindWindowEx(parent&, RICHEDIT&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "570    Child1& = FindWindowEx(parent&, 0&, " & Chr(34) & "Edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
570    Child1& = FindWindowEx(parent&, 0&, "Edit", vbNullString)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "571    Call SetFocusApi(Child1&)"
' </VB WATCH>
571    Call SetFocusApi(Child1&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "572    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, UserName$)"
' </VB WATCH>
572    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, UserName$)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "573    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)"
' </VB WATCH>
573    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "574    Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
574    Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)


' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "575    Call SetFocusApi(Child2&)"
' </VB WATCH>
575    Call SetFocusApi(Child2&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "576    Do"
' </VB WATCH>
576    Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "577    Call SendMessageByString(Child2&, WM_SETTEXT, 0&, Code$)"
' </VB WATCH>
577    Call SendMessageByString(Child2&, WM_SETTEXT, 0&, code$)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "578    Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
578    Call ClickMenu(parent&, "Sen&d")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "579    Call pause(0.3)"
' </VB WATCH>
579    Call Pause(0.3)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "580    Times% = Times% - 1"
' </VB WATCH>
580    Times% = Times% - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "581    Loop Until Times% = 0"
' </VB WATCH>
581    Loop Until Times% = 0
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "582    Call WindowClose(parent&)"
' </VB WATCH>
582    Call WindowClose(parent&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
583        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
584        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendPMLagg"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Function GetYahooText()
' <VB WATCH>
585        On Error GoTo vbwErrHandler
586        Const VBWPROCNAME = "Module1.GetYahooText"
587        If vbwTraceProc Then
588            Dim vbwParameterString As String
589            If vbwTraceParameters Then
590                vbwParameterString = "()"
591            End If
592            vbwTraceIn VBWPROCNAME, vbwParameterString
593        End If
' </VB WATCH>
594    Dim imc As Long, Rich As Long
595    Dim texts As String, thetextlen As Long

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "596    imc& = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
596    imc& = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "597    Rich& = FindWindowEx(imc&, 0&, " & Chr(34) & "richedit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
597    Rich& = FindWindowEx(imc&, 0&, "richedit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "598    Rich& = FindWindowEx(imc&, Rich, " & Chr(34) & "richedit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
598    Rich& = FindWindowEx(imc&, Rich, "richedit", vbNullString)
599    Dim TheText As String, TL As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "600    TL = SendMessageLong(Rich&, WM_GETTEXTLENGTH, 0&, 0&)"
' </VB WATCH>
600    TL = SendMessageLong(Rich&, WM_GETTEXTLENGTH, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "601    TheText = String(TL + 1, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
601    TheText = String(TL + 1, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "602    Call SendMessageByString(Rich&, WM_gettext, TL + 1, TheText)"
' </VB WATCH>
602    Call SendMessageByString(Rich&, WM_GETTEXT, TL + 1, TheText)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "603    TheText = Left(TheText, TL)"
' </VB WATCH>
603    TheText = Left(TheText, TL)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "604    If TheText = " & Chr(34) & "" & Chr(34) & " Then"
' </VB WATCH>
604    If TheText = "" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "605         GoTo NoText"
' </VB WATCH>
605         GoTo NoText
606    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "606    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "607            thetextlen& = (Len(TheText) - 2)"
' </VB WATCH>
607            thetextlen& = (Len(TheText) - 2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "608            TheText$ = Left$(TheText, thetextlen&)"
' </VB WATCH>
608            TheText$ = Left$(TheText, thetextlen&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "609    GetYahooText = TheText"
' </VB WATCH>
609    GetYahooText = TheText
610 NoText:
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
611        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
612        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetYahooText"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Sub SendPMScroll(who$, what$, Times As Integer, Follow As Boolean)
' <VB WATCH>
613        On Error GoTo vbwErrHandler
614        Const VBWPROCNAME = "Module1.SendPMScroll"
615        If vbwTraceProc Then
616            Dim vbwParameterString As String
617            If vbwTraceParameters Then
618                vbwParameterString = "(" & vbwReportParameter("Who$", who$) & ", "
619                vbwParameterString = vbwParameterString & vbwReportParameter("what$", what$) & ", "
620                vbwParameterString = vbwParameterString & vbwReportParameter("Times", Times) & ", "
621                vbwParameterString = vbwParameterString & vbwReportParameter("Follow", Follow) & ") "
622            End If
623            vbwTraceIn VBWPROCNAME, vbwParameterString
624        End If
' </VB WATCH>
625    Dim yahoobuddymain As Long, parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "626    yahoobuddymain = FindWindow(" & Chr(34) & "yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
626    yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "627    Call ClickMenu(yahoobuddymain&, " & Chr(34) & "Send a &Message" & Chr(34) & ")"
' </VB WATCH>
627    Call ClickMenu(yahoobuddymain&, "Send a &Message")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "628    Call pause(0.2)"
' </VB WATCH>
628    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "629    parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
629    parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "630    If InStr(GetCaption(parent&), LCase(" & Chr(34) & "chat" & Chr(34) & ")) Then"
' </VB WATCH>
630    If InStr(GetCaption(parent&), LCase("chat")) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "631"
' </VB WATCH>
631
' <VB WATCH>
632        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "633        Exit Sub"
' </VB WATCH>
633        Exit Sub
634    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "634    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "635    Call SetFocusApi(parent&)"
' </VB WATCH>
635    Call SetFocusApi(parent&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "636    Child1& = FindWindowEx(parent&, 0&, " & Chr(34) & "Edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
636    Child1& = FindWindowEx(parent&, 0&, "Edit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "637    Call SetFocusApi(Child1&)"
' </VB WATCH>
637    Call SetFocusApi(Child1&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "638    Call SendMessageByString(Child1&, WM_SETTEXT, 0, Who$)"
' </VB WATCH>
638    Call SendMessageByString(Child1&, WM_SETTEXT, 0, who$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "639    Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
639    Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "640    Do"
' </VB WATCH>
640    Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "641    Call SendMessageByString(Child2&, WM_SETTEXT, 0, what$)"
' </VB WATCH>
641    Call SendMessageByString(Child2&, WM_SETTEXT, 0, what$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "642    Times% = Times% - 1"
' </VB WATCH>
642    Times% = Times% - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "643    Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
643    Call ClickMenu(parent&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "644    Call pause(0.3)"
' </VB WATCH>
644    Call Pause(0.3)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "645    Loop Until Times% = 0"
' </VB WATCH>
645    Loop Until Times% = 0
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "646    If Follow = False Then"
' </VB WATCH>
646    If Follow = False Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "647"
' </VB WATCH>
647
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "648        Call WindowClose(parent&)"
' </VB WATCH>
648        Call WindowClose(parent&)
649    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "649    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
650        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
651        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendPMScroll"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SendFile(who$, file$, Message$)
' <VB WATCH>
652        On Error GoTo vbwErrHandler
653        Const VBWPROCNAME = "Module1.SendFile"
654        If vbwTraceProc Then
655            Dim vbwParameterString As String
656            If vbwTraceParameters Then
657                vbwParameterString = "(" & vbwReportParameter("Who$", who$) & ", "
658                vbwParameterString = vbwParameterString & vbwReportParameter("FilE$", file$) & ", "
659                vbwParameterString = vbwParameterString & vbwReportParameter("Message$", Message$) & ") "
660            End If
661            vbwTraceIn VBWPROCNAME, vbwParameterString
662        End If
' </VB WATCH>
663    Dim yahoo As Long, imclass As Long, RICHEDIT As Long, editx As Long, Button As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "664    Yahoo = FindWindow(" & Chr(34) & "yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
664    yahoo = FindWindow("yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "665    Call ClickMenu(Yahoo&, " & Chr(34) & "Send a &File..." & Chr(34) & ")"
' </VB WATCH>
665    Call ClickMenu(yahoo&, "Send a &File...")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "666    Call pause(0.2)"
' </VB WATCH>
666    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "667    imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", " & Chr(34) & "Send a File..." & Chr(34) & ")"
' </VB WATCH>
667    imclass = FindWindow("imclass", "Send a File...")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "668    RichEdit = FindWindowEx(imclass, 0&, " & Chr(34) & "richedit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
668    RICHEDIT = FindWindowEx(imclass, 0&, "richedit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "669    Call SetFocusApi(RichEdit&)"
' </VB WATCH>
669    Call SetFocusApi(RICHEDIT&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "670    Call SendMessageByString(RichEdit&, WM_SETTEXT, 0, Who$)"
' </VB WATCH>
670    Call SendMessageByString(RICHEDIT&, WM_SETTEXT, 0, who$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "671    Call pause(0.2)"
' </VB WATCH>
671    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "672    editx = FindWindowEx(imclass, 0&, " & Chr(34) & "edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
672    editx = FindWindowEx(imclass, 0&, "edit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "673    Call SetFocusApi(editx&)"
' </VB WATCH>
673    Call SetFocusApi(editx&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "674    Call SendMessageByString(editx&, WM_SETTEXT, 0, FilE$)"
' </VB WATCH>
674    Call SendMessageByString(editx&, WM_SETTEXT, 0, file$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "675    Call pause(0.2)"
' </VB WATCH>
675    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "676    editx = FindWindowEx(imclass, editx, " & Chr(34) & "edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
676    editx = FindWindowEx(imclass, editx, "edit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "677    Call SetFocusApi(editx&)"
' </VB WATCH>
677    Call SetFocusApi(editx&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "678    Call SendMessageByString(editx, WM_SETTEXT, 0&, Message$)"
' </VB WATCH>
678    Call SendMessageByString(editx, WM_SETTEXT, 0&, Message$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "679    Button = FindWindowEx(imclass, 0&, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
679    Button = FindWindowEx(imclass, 0&, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "680    Button = FindWindowEx(imclass, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
680    Button = FindWindowEx(imclass, Button, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "681    Call SetFocusApi(Button&)"
' </VB WATCH>
681    Call SetFocusApi(Button&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "682    Call ClickButton(Button&)"
' </VB WATCH>
682    Call ClickButton(Button&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
683        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
684        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendFile"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub YahooClose()
' <VB WATCH>
685        On Error GoTo vbwErrHandler
686        Const VBWPROCNAME = "Module1.YahooClose"
687        If vbwTraceProc Then
688            Dim vbwParameterString As String
689            If vbwTraceParameters Then
690                vbwParameterString = "()"
691            End If
692            vbwTraceIn VBWPROCNAME, vbwParameterString
693        End If
' </VB WATCH>
694    Dim yahoobuddymain As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "695    yahoobuddymain = FindWindow(" & Chr(34) & "yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
695    yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "696    Call ClickMenu(yahoobuddymain&, " & Chr(34) & "C&lose" & Chr(34) & ")"
' </VB WATCH>
696    Call ClickMenu(yahoobuddymain&, "C&lose")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
697        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
698        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "YahooClose"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub chatclear()
' <VB WATCH>
699        On Error GoTo vbwErrHandler
700        Const VBWPROCNAME = "Module1.ChatClear"
701        If vbwTraceProc Then
702            Dim vbwParameterString As String
703            If vbwTraceParameters Then
704                vbwParameterString = "()"
705            End If
706            vbwTraceIn VBWPROCNAME, vbwParameterString
707        End If
' </VB WATCH>
708    Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "709    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
709    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "710    RichEdit& = FindWindowEx(imclass&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
710    RICHEDIT& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "711    RichEdit& = FindWindowEx(imclass&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
711    RICHEDIT& = FindWindowEx(imclass&, RICHEDIT&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "712    If InStr(LCase(GetCaption(imclass&)), " & Chr(34) & "chat" & Chr(34) & ") Then"
' </VB WATCH>
712    If InStr(LCase(GetCaption(imclass&)), "chat") Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "713    Call SendMessageByString(RichEdit&, WM_SETTEXT, 0&, " & Chr(34) & "" & Chr(34) & ")"
' </VB WATCH>
713    Call SendMessageByString(RICHEDIT&, WM_SETTEXT, 0&, "")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
714        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
715        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChatClear"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End If
End Sub
Sub ChatHide()
' <VB WATCH>
716        On Error GoTo vbwErrHandler
717        Const VBWPROCNAME = "Module1.ChatHide"
718        If vbwTraceProc Then
719            Dim vbwParameterString As String
720            If vbwTraceParameters Then
721                vbwParameterString = "()"
722            End If
723            vbwTraceIn VBWPROCNAME, vbwParameterString
724        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "725    Call WindowHide(FindChatWnd)"
' </VB WATCH>
725    Call WindowHide(FindChatWnd)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
726        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
727        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChatHide"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub ChatShow()
' <VB WATCH>
728        On Error GoTo vbwErrHandler
729        Const VBWPROCNAME = "Module1.ChatShow"
730        If vbwTraceProc Then
731            Dim vbwParameterString As String
732            If vbwTraceParameters Then
733                vbwParameterString = "()"
734            End If
735            vbwTraceIn VBWPROCNAME, vbwParameterString
736        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "737    Call WindowShow(FindChatWnd)"
' </VB WATCH>
737    Call WindowShow(FindChatWnd)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
738        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
739        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChatShow"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub ChatClearTheirs()
' <VB WATCH>
740        On Error GoTo vbwErrHandler
741        Const VBWPROCNAME = "Module1.ChatClearTheirs"
742        If vbwTraceProc Then
743            Dim vbwParameterString As String
744            If vbwTraceParameters Then
745                vbwParameterString = "()"
746            End If
747            vbwTraceIn VBWPROCNAME, vbwParameterString
748        End If
' </VB WATCH>
749    Dim Text As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "750    Text$ = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf"
' </VB WATCH>
750    Text$ = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "751    Call SendChat(Text$)"
' </VB WATCH>
751    Call SendChat(Text$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
752        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
753        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChatClearTheirs"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub PMClear()
       'Clears PM
' <VB WATCH>
754        On Error GoTo vbwErrHandler
755        Const VBWPROCNAME = "Module1.PMClear"
756        If vbwTraceProc Then
757            Dim vbwParameterString As String
758            If vbwTraceParameters Then
759                vbwParameterString = "()"
760            End If
761            vbwTraceIn VBWPROCNAME, vbwParameterString
762        End If
' </VB WATCH>
763    Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "764    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
764    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "765    RichEdit& = FindWindowEx(imclass&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
765    RICHEDIT& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "766    RichEdit& = FindWindowEx(imclass&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
766    RICHEDIT& = FindWindowEx(imclass&, RICHEDIT&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "767    If InStr(LCase(GetCaption(imclass&)), " & Chr(34) & "instant message" & Chr(34) & ") Then"
' </VB WATCH>
767    If InStr(LCase(GetCaption(imclass&)), "instant message") Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "768    Call SendMessageByString(RichEdit&, WM_SETTEXT, 0&, " & Chr(34) & "" & Chr(34) & ")"
' </VB WATCH>
768    Call SendMessageByString(RICHEDIT&, WM_SETTEXT, 0&, "")
769    Else
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "769    Else" 'B
' </VB WATCH>
770    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "770    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
771        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
772        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PMClear"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub PMHide()
' <VB WATCH>
773        On Error GoTo vbwErrHandler
774        Const VBWPROCNAME = "Module1.PMHide"
775        If vbwTraceProc Then
776            Dim vbwParameterString As String
777            If vbwTraceParameters Then
778                vbwParameterString = "()"
779            End If
780            vbwTraceIn VBWPROCNAME, vbwParameterString
781        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "782    Call WindowHide(FindPMWnd)"
' </VB WATCH>
782    Call WindowHide(FindPMWnd)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
783        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
784        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PMHide"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub PMShow()
' <VB WATCH>
785        On Error GoTo vbwErrHandler
786        Const VBWPROCNAME = "Module1.PMShow"
787        If vbwTraceProc Then
788            Dim vbwParameterString As String
789            If vbwTraceParameters Then
790                vbwParameterString = "()"
791            End If
792            vbwTraceIn VBWPROCNAME, vbwParameterString
793        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "794    Call WindowShow(FindPMWnd)"
' </VB WATCH>
794    Call WindowShow(FindPMWnd)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
795        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
796        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PMShow"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Function PMFrom()
       'Get's the open pm user
' <VB WATCH>
797        On Error GoTo vbwErrHandler
798        Const VBWPROCNAME = "Module1.PMFrom"
799        If vbwTraceProc Then
800            Dim vbwParameterString As String
801            If vbwTraceParameters Then
802                vbwParameterString = "()"
803            End If
804            vbwTraceIn VBWPROCNAME, vbwParameterString
805        End If
' </VB WATCH>
806    Dim imclass As Long, Str As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "807    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
807    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "808    If InStr(LCase(GetCaption(imclass&)), " & Chr(34) & "instant message" & Chr(34) & ") Then"
' </VB WATCH>
808    If InStr(LCase(GetCaption(imclass&)), "instant message") Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "809    Str$ = GetCaption(imclass&)"
' </VB WATCH>
809    Str$ = GetCaption(imclass&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "810    Str$ = Replace(Str$, " & Chr(34) & " -- Instant Message" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ")"
' </VB WATCH>
810    Str$ = Replace(Str$, " -- Instant Message", "")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "811    PMFrom = Str$"
' </VB WATCH>
811    PMFrom = Str$
812    Else
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "812    Else" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "813    PMFrom = " & Chr(34) & "" & Chr(34) & ""
' </VB WATCH>
813    PMFrom = ""
814    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "814    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
815        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
816        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PMFrom"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function PMIgnore()
       'Ignores current user.
' <VB WATCH>
817        On Error GoTo vbwErrHandler
818        Const VBWPROCNAME = "Module1.PMIgnore"
819        If vbwTraceProc Then
820            Dim vbwParameterString As String
821            If vbwTraceParameters Then
822                vbwParameterString = "()"
823            End If
824            vbwTraceIn VBWPROCNAME, vbwParameterString
825        End If
' </VB WATCH>
826    Dim imclass As Long, Str As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "827    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
827    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "828    If InStr(LCase(GetCaption(imclass&)), " & Chr(34) & "instant message" & Chr(34) & ") Then"
' </VB WATCH>
828    If InStr(LCase(GetCaption(imclass&)), "instant message") Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "829    Call ClickMenu(imclass&, " & Chr(34) & "&Ignore User..." & Chr(34) & ")"
' </VB WATCH>
829    Call ClickMenu(imclass&, "&Ignore User...")
830    Else
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "830    Else" 'B
' </VB WATCH>
831    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "831    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
832        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
833        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PMIgnore"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function PMVoiceOnOff()
' <VB WATCH>
834        On Error GoTo vbwErrHandler
835        Const VBWPROCNAME = "Module1.PMVoiceOnOff"
836        If vbwTraceProc Then
837            Dim vbwParameterString As String
838            If vbwTraceParameters Then
839                vbwParameterString = "()"
840            End If
841            vbwTraceIn VBWPROCNAME, vbwParameterString
842        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "843    Call ClickMenu(FindPMWnd, " & Chr(34) & "Enable &Voice" & Chr(34) & ")"
' </VB WATCH>
843    Call ClickMenu(FindPMWnd, "Enable &Voice")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
844        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
845        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PMVoiceOnOff"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function Ignore(User As String)
       'Ignores specific user
' <VB WATCH>
846        On Error GoTo vbwErrHandler
847        Const VBWPROCNAME = "Module1.Ignore"
848        If vbwTraceProc Then
849            Dim vbwParameterString As String
850            If vbwTraceParameters Then
851                vbwParameterString = "(" & vbwReportParameter("User", User) & ") "
852            End If
853            vbwTraceIn VBWPROCNAME, vbwParameterString
854        End If
' </VB WATCH>
855    Dim yahoobuddymain As Long, parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "856    yahoobuddymain = FindWindow(" & Chr(34) & "yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
856    yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "857    Call ClickMenu(yahoobuddymain&, " & Chr(34) & "Send a &Message" & Chr(34) & ")"
' </VB WATCH>
857    Call ClickMenu(yahoobuddymain&, "Send a &Message")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "858    Call pause(0.2)"
' </VB WATCH>
858    Call Pause(0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "859    parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
859    parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "860    Child1& = FindWindowEx(parent&, 0&, " & Chr(34) & "Edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
860    Child1& = FindWindowEx(parent&, 0&, "Edit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "861    Call SetFocusApi(Child1&)"
' </VB WATCH>
861    Call SetFocusApi(Child1&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "862    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, User$)"
' </VB WATCH>
862    Call SendMessageByString(Child1&, WM_SETTEXT, 0&, User$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "863    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)"
' </VB WATCH>
863    Call SendMessageByNum(Child1&, WM_CHAR, 13, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "864    Call ClickMenu(parent&, " & Chr(34) & "&Ignore User..." & Chr(34) & ")"
' </VB WATCH>
864    Call ClickMenu(parent&, "&Ignore User...")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
865        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
866        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Ignore"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function

Sub PMClose()
       'Closes PM
' <VB WATCH>
867        On Error GoTo vbwErrHandler
868        Const VBWPROCNAME = "Module1.PMClose"
869        If vbwTraceProc Then
870            Dim vbwParameterString As String
871            If vbwTraceParameters Then
872                vbwParameterString = "()"
873            End If
874            vbwTraceIn VBWPROCNAME, vbwParameterString
875        End If
' </VB WATCH>
876    Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "877    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
877    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "878    RichEdit& = FindWindowEx(imclass&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
878    RICHEDIT& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "879    RichEdit& = FindWindowEx(imclass&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
879    RICHEDIT& = FindWindowEx(imclass&, RICHEDIT&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "880    If InStr(LCase(GetCaption(imclass&)), " & Chr(34) & "instant message" & Chr(34) & ") Then"
' </VB WATCH>
880    If InStr(LCase(GetCaption(imclass&)), "instant message") Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "881    Call WindowClose(imclass&)"
' </VB WATCH>
881    Call WindowClose(imclass&)
882    Else
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "882    Else" 'B
' </VB WATCH>
883    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "883    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
884        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
885        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PMClose"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub NewAnti()
' <VB WATCH>
886        On Error GoTo vbwErrHandler
887        Const VBWPROCNAME = "Module1.NewAnti"
888        If vbwTraceProc Then
889            Dim vbwParameterString As String
890            If vbwTraceParameters Then
891                vbwParameterString = "()"
892            End If
893            vbwTraceIn VBWPROCNAME, vbwParameterString
894        End If
' </VB WATCH>
895    Dim imclass As Long, atleeb As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "896    imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
896    imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "897    atleeb = FindWindowEx(imclass, 0&, " & Chr(34) & "atl:004eeb68" & Chr(34) & ", vbNullString)"
' </VB WATCH>
897    atleeb = FindWindowEx(imclass, 0&, "atl:004eeb68", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "898    Call SendMessageLong(atleeb, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
898    Call SendMessageLong(atleeb, WM_CLOSE, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
899        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
900        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NewAnti"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub ChatClose()
       'Closes Chat
' <VB WATCH>
901        On Error GoTo vbwErrHandler
902        Const VBWPROCNAME = "Module1.ChatClose"
903        If vbwTraceProc Then
904            Dim vbwParameterString As String
905            If vbwTraceParameters Then
906                vbwParameterString = "()"
907            End If
908            vbwTraceIn VBWPROCNAME, vbwParameterString
909        End If
' </VB WATCH>
910    Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "911    imclass& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
911    imclass& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "912    RichEdit& = FindWindowEx(imclass&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
912    RICHEDIT& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "913    RichEdit& = FindWindowEx(imclass&, RichEdit&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
913    RICHEDIT& = FindWindowEx(imclass&, RICHEDIT&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "914    If InStr(LCase(GetCaption(imclass&)), " & Chr(34) & "chat" & Chr(34) & ") Then"
' </VB WATCH>
914    If InStr(LCase(GetCaption(imclass&)), "chat") Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "915    Call WindowClose(imclass&)"
' </VB WATCH>
915    Call WindowClose(imclass&)
916    Else
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "916    Else" 'B
' </VB WATCH>
917    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "917    End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
918        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
919        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChatClose"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Function ChatVoiceOnOff()
' <VB WATCH>
920        On Error GoTo vbwErrHandler
921        Const VBWPROCNAME = "Module1.ChatVoiceOnOff"
922        If vbwTraceProc Then
923            Dim vbwParameterString As String
924            If vbwTraceParameters Then
925                vbwParameterString = "()"
926            End If
927            vbwTraceIn VBWPROCNAME, vbwParameterString
928        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "929    Call ClickMenu(FindChatWnd, " & Chr(34) & "Enable &Voice" & Chr(34) & ")"
' </VB WATCH>
929    Call ClickMenu(FindChatWnd, "Enable &Voice")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
930        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
931        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChatVoiceOnOff"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Sub SendChatScroll(what As String, Times As Integer)
       'Sends a Chat Scroll
' <VB WATCH>
932        On Error GoTo vbwErrHandler
933        Const VBWPROCNAME = "Module1.SendChatScroll"
934        If vbwTraceProc Then
935            Dim vbwParameterString As String
936            If vbwTraceParameters Then
937                vbwParameterString = "(" & vbwReportParameter("what", what) & ", "
938                vbwParameterString = vbwParameterString & vbwReportParameter("Times", Times) & ") "
939            End If
940            vbwTraceIn VBWPROCNAME, vbwParameterString
941        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "942    Do"
' </VB WATCH>
942    Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "943    Call SendChat(what$)"
' </VB WATCH>
943    Call SendChat(what$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "944    Call pause(0.3)"
' </VB WATCH>
944    Call Pause(0.3)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "945    Times% = Times% - 1"
' </VB WATCH>
945    Times% = Times% - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "946    Loop Until Times% = 0"
' </VB WATCH>
946    Loop Until Times% = 0
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
947        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
948        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendChatScroll"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SendChatLagg(code$, Times%)
       'Send's a Chat Lagg
' <VB WATCH>
949        On Error GoTo vbwErrHandler
950        Const VBWPROCNAME = "Module1.SendChatLagg"
951        If vbwTraceProc Then
952            Dim vbwParameterString As String
953            If vbwTraceParameters Then
954                vbwParameterString = "(" & vbwReportParameter("Code$", code$) & ", "
955                vbwParameterString = vbwParameterString & vbwReportParameter("Times%", Times%) & ") "
956            End If
957            vbwTraceIn VBWPROCNAME, vbwParameterString
958        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "959    Do"
' </VB WATCH>
959    Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "960    Call SendChat(Code$)"
' </VB WATCH>
960    Call SendChat(code$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "961    Call pause(0.3)"
' </VB WATCH>
961    Call Pause(0.3)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "962    Call ChatClear"
' </VB WATCH>
962    Call chatclear
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "963    Times% = Times% - 1"
' </VB WATCH>
963    Times% = Times% - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "964    Loop Until Times% = 0"
' </VB WATCH>
964    Loop Until Times% = 0
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
965        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
966        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendChatLagg"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub SignIn(UserName$, password$, SaveID As Boolean, AutoLogin As Boolean, Invisible As Boolean)
' <VB WATCH>
967        On Error GoTo vbwErrHandler
968        Const VBWPROCNAME = "Module1.SignIn"
969        If vbwTraceProc Then
970            Dim vbwParameterString As String
971            If vbwTraceParameters Then
972                vbwParameterString = "(" & vbwReportParameter("UserName$", UserName$) & ", "
973                vbwParameterString = vbwParameterString & vbwReportParameter("password$", password$) & ", "
974                vbwParameterString = vbwParameterString & vbwReportParameter("SaveID", SaveID) & ", "
975                vbwParameterString = vbwParameterString & vbwReportParameter("AutoLogin", AutoLogin) & ", "
976                vbwParameterString = vbwParameterString & vbwReportParameter("Invisible", Invisible) & ") "
977            End If
978            vbwTraceIn VBWPROCNAME, vbwParameterString
979        End If
' </VB WATCH>
980    Dim X As Long, editx As Long, Button As Long
981    Dim yahoobuddymain As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "982    yahoobuddymain = FindWindow(" & Chr(34) & "yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
982    yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "983    Call ClickMenu(yahoobuddymain&, " & Chr(34) & "C&lose" & Chr(34) & ")"
' </VB WATCH>
983    Call ClickMenu(yahoobuddymain&, "C&lose")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "984    X = FindWindow(" & Chr(34) & "#32770" & Chr(34) & ", " & Chr(34) & "Login" & Chr(34) & ")"
' </VB WATCH>
984    X = FindWindow("#32770", "Login")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "985    If X& = True Then"
' </VB WATCH>
985    If X& = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "986         GoTo SetText"
' </VB WATCH>
986         GoTo SetText
987    Else
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "987    Else" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "988"
' </VB WATCH>
988
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "989        Call ClickMenu(yahoobuddymain&, " & Chr(34) & "&Login..." & Chr(34) & ")"
' </VB WATCH>
989        Call ClickMenu(yahoobuddymain&, "&Login...")
990    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "990    End If" 'B
' </VB WATCH>

991 SetText:
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "992    editx = FindWindowEx(X, 0&, " & Chr(34) & "edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
992    editx = FindWindowEx(X, 0&, "edit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "993    Call SetFocusApi(editx)"
' </VB WATCH>
993    Call SetFocusApi(editx)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "994    Call SendMessageByString(editx, WM_SETTEXT, 0&, UserName$)"
' </VB WATCH>
994    Call SendMessageByString(editx, WM_SETTEXT, 0&, UserName$)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "995    editx = FindWindowEx(X, editx, " & Chr(34) & "edit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
995    editx = FindWindowEx(X, editx, "edit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "996    Call SetFocusApi(editx)"
' </VB WATCH>
996    Call SetFocusApi(editx)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "997    Call SendMessageByString(editx, WM_SETTEXT, 0&, password$)"
' </VB WATCH>
997    Call SendMessageByString(editx, WM_SETTEXT, 0&, password$)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "998    Button& = FindWindowEx(X&, 0&, " & Chr(34) & "Button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
998    Button& = FindWindowEx(X&, 0&, "Button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "999    If SaveID = True Then"
' </VB WATCH>
999    If SaveID = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1000"
' </VB WATCH>
1000
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1001       Call SendMessage(Button&, BM_SETCHECK, True, 0&)"
' </VB WATCH>
1001       Call SendMessage(Button&, BM_SETCHECK, True, 0&)
1002   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1002   End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1003   Button = FindWindowEx(X, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1003   Button = FindWindowEx(X, Button, "button", vbNullString)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1004   If AutoLogin = True Then"
' </VB WATCH>
1004   If AutoLogin = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1005"
' </VB WATCH>
1005
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1006       Call SendMessageLong(Button, BM_SETCHECK, True, 0&)"
' </VB WATCH>
1006       Call SendMessageLong(Button, BM_SETCHECK, True, 0&)
1007   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1007   End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1008   Button = FindWindowEx(X, 0&, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1008   Button = FindWindowEx(X, 0&, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1009   Button = FindWindowEx(X, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1009   Button = FindWindowEx(X, Button, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1010   Button = FindWindowEx(X, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1010   Button = FindWindowEx(X, Button, "button", vbNullString)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1011   If Invisible = True Then"
' </VB WATCH>
1011   If Invisible = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1012"
' </VB WATCH>
1012
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1013       Call SendMessageLong(Button, BM_SETCHECK, True, 0&)"
' </VB WATCH>
1013       Call SendMessageLong(Button, BM_SETCHECK, True, 0&)
1014   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1014   End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1015   Button = FindWindowEx(X, 0&, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1015   Button = FindWindowEx(X, 0&, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1016   Button = FindWindowEx(X, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1016   Button = FindWindowEx(X, Button, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1017   Button = FindWindowEx(X, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1017   Button = FindWindowEx(X, Button, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1018   Button = FindWindowEx(X, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1018   Button = FindWindowEx(X, Button, "button", vbNullString)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1019   Call SetFocusApi(Button&)"
' </VB WATCH>
1019   Call SetFocusApi(Button&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1020   Call pause(0.3)"
' </VB WATCH>
1020   Call Pause(0.3)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1021   Call ClickButton(Button&)"
' </VB WATCH>
1021   Call ClickButton(Button&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1022       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1023       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SignIn"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Function AntiError()
' <VB WATCH>
1024       On Error GoTo vbwErrHandler
1025       Const VBWPROCNAME = "Module1.AntiError"
1026       If vbwTraceProc Then
1027           Dim vbwParameterString As String
1028           If vbwTraceParameters Then
1029               vbwParameterString = "()"
1030           End If
1031           vbwTraceIn VBWPROCNAME, vbwParameterString
1032       End If
' </VB WATCH>
1033   Dim child As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1034   child& = FindWindow(" & Chr(34) & "#32770" & Chr(34) & ", " & Chr(34) & "Chat Error" & Chr(34) & ")"
' </VB WATCH>
1034   child& = FindWindow("#32770", "Chat Error")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1035   Call SendMessage(child&, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
1035   Call SendMessage(child&, WM_CLOSE, 0&, 0&)
       'Closes Chat Error
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1036       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1037       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AntiError"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function AntiLagg()
' <VB WATCH>
1038       On Error GoTo vbwErrHandler
1039       Const VBWPROCNAME = "Module1.AntiLagg"
1040       If vbwTraceProc Then
1041           Dim vbwParameterString As String
1042           If vbwTraceParameters Then
1043               vbwParameterString = "()"
1044           End If
1045           vbwTraceIn VBWPROCNAME, vbwParameterString
1046       End If
' </VB WATCH>
1047   Dim t As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1048   Do"
' </VB WATCH>
1048   Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1049   DoEvents"
' </VB WATCH>
1049   DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1050   t% = t% + 1"
' </VB WATCH>
1050   t% = t% + 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1051   If t% = 50 Then"
' </VB WATCH>
1051   If t% = 50 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1052        Exit Do"
' </VB WATCH>
1052        Exit Do
1053   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1053   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1054   Loop"
' </VB WATCH>
1054   Loop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1055       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1056       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AntiLagg"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Public Sub MassPM(List As ListBox, Message As String, Follow As Boolean)
' <VB WATCH>
1057       On Error GoTo vbwErrHandler
1058       Const VBWPROCNAME = "Module1.MassPM"
1059       If vbwTraceProc Then
1060           Dim vbwParameterString As String
1061           If vbwTraceParameters Then
1062               vbwParameterString = "(" & vbwReportParameter("List", List) & ", "
1063               vbwParameterString = vbwParameterString & vbwReportParameter("Message", Message) & ", "
1064               vbwParameterString = vbwParameterString & vbwReportParameter("Follow", Follow) & ") "
1065           End If
1066           vbwTraceIn VBWPROCNAME, vbwParameterString
1067       End If
' </VB WATCH>

1068   Dim Scrll As Integer, Num As Integer, Str As String

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1069   Num% = 0"
' </VB WATCH>
1069   Num% = 0

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1070   For Scrll% = 0 To List.ListCount - 1"
' </VB WATCH>
1070   For Scrll% = 0 To List.ListCount - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1071       Str$ = List.List(Scrll%)"
' </VB WATCH>
1071       Str$ = List.List(Scrll%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1072           If Num% >= 5 Then"
' </VB WATCH>
1072           If Num% >= 5 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1073               pause (3)"
' </VB WATCH>
1073               Pause (3)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1074               Num% = 0"
' </VB WATCH>
1074               Num% = 0
1075           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1075           End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1076           If Follow = True Then"
' </VB WATCH>
1076           If Follow = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1077"
' </VB WATCH>
1077
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1078               Call SendPM(Str$, Message$, True)"
' </VB WATCH>
1078               Call SendPM(Str$, Message$, True)
1079           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1079           End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1080           If Follow = False Then"
' </VB WATCH>
1080           If Follow = False Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1081"
' </VB WATCH>
1081
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1082               Call SendPM(Str$, Message$, False)"
' </VB WATCH>
1082               Call SendPM(Str$, Message$, False)
1083           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1083           End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1084           pause (0.2)"
' </VB WATCH>
1084           Pause (0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1085       Num% = Num% + 1"
' </VB WATCH>
1085       Num% = Num% + 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1086       DoEvents"
' </VB WATCH>
1086       DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1087   Next"
' </VB WATCH>
1087   Next

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1088       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1089       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "MassPM"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Sub MassPMBoot(List As ListBox, Message As String, ANTI As Boolean)
' <VB WATCH>
1090       On Error GoTo vbwErrHandler
1091       Const VBWPROCNAME = "Module1.MassPMBoot"
1092       If vbwTraceProc Then
1093           Dim vbwParameterString As String
1094           If vbwTraceParameters Then
1095               vbwParameterString = "(" & vbwReportParameter("List", List) & ", "
1096               vbwParameterString = vbwParameterString & vbwReportParameter("Message", Message) & ", "
1097               vbwParameterString = vbwParameterString & vbwReportParameter("anti", ANTI) & ") "
1098           End If
1099           vbwTraceIn VBWPROCNAME, vbwParameterString
1100       End If
' </VB WATCH>

1101   Dim Scrll As Integer, Num As Integer, Str As String

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1102   Num% = 0"
' </VB WATCH>
1102   Num% = 0

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1103   For Scrll% = 0 To List.ListCount - 1"
' </VB WATCH>
1103   For Scrll% = 0 To List.ListCount - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1104       Str$ = List.List(Scrll%)"
' </VB WATCH>
1104       Str$ = List.List(Scrll%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1105           If Num% >= 5 Then"
' </VB WATCH>
1105           If Num% >= 5 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1106               pause (3)"
' </VB WATCH>
1106               Pause (3)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1107               Num% = 0"
' </VB WATCH>
1107               Num% = 0
1108           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1108           End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1109           If anti = True Then"
' </VB WATCH>
1109           If ANTI = True Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1110"
' </VB WATCH>
1110
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1111               Call SendPMBoot(Str$, Message$, True, False)"
' </VB WATCH>
1111               Call SendPMBoot(Str$, Message$, True, False)
1112           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1112           End If" 'B
' </VB WATCH>
               'Determins what the boolean's are set to.
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1113           If anti = False Then"
' </VB WATCH>
1113           If ANTI = False Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1114"
' </VB WATCH>
1114
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1115               Call SendPMBoot(Str$, Message$, False, False)"
' </VB WATCH>
1115               Call SendPMBoot(Str$, Message$, False, False)
1116           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1116           End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1117           pause (0.2)"
' </VB WATCH>
1117           Pause (0.2)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1118       Num% = Num% + 1"
' </VB WATCH>
1118       Num% = Num% + 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1119       DoEvents"
' </VB WATCH>
1119       DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1120   Next"
' </VB WATCH>
1120   Next

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1121       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1122       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "MassPMBoot"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Function GetChatName()
' <VB WATCH>
1123       On Error GoTo vbwErrHandler
1124       Const VBWPROCNAME = "Module1.GetChatName"
1125       If vbwTraceProc Then
1126           Dim vbwParameterString As String
1127           If vbwTraceParameters Then
1128               vbwParameterString = "()"
1129           End If
1130           vbwTraceIn VBWPROCNAME, vbwParameterString
1131       End If
' </VB WATCH>
1132   Dim imclass As Long
1133   Dim Str As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1134   imclass& = FindWindow(imclass&, vbNullString)"
' </VB WATCH>
1134   imclass& = FindWindow(imclass&, vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1135   Str$ = GetCaption(imclass&)"
' </VB WATCH>
1135   Str$ = GetCaption(imclass&)
       'Get's Caption
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1136   GetChatName = Replace(Str$, " & Chr(34) & "-- Chat" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ")"
' </VB WATCH>
1136   GetChatName = Replace(Str$, "-- Chat", "")
       'Get's Caption Filterd and returns caption w/ out Chat
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1137       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1138       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetChatName"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function GetPMName()
' <VB WATCH>
1139       On Error GoTo vbwErrHandler
1140       Const VBWPROCNAME = "Module1.GetPMName"
1141       If vbwTraceProc Then
1142           Dim vbwParameterString As String
1143           If vbwTraceParameters Then
1144               vbwParameterString = "()"
1145           End If
1146           vbwTraceIn VBWPROCNAME, vbwParameterString
1147       End If
' </VB WATCH>
1148   Dim imclass As Long
1149   Dim Str As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1150   imclass& = FindWindow(imclass&, vbNullString)"
' </VB WATCH>
1150   imclass& = FindWindow(imclass&, vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1151   Str$ = GetCaption(imclass&)"
' </VB WATCH>
1151   Str$ = GetCaption(imclass&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1152   GetChatName = Replace(Str$, " & Chr(34) & "-- Instant Message" & Chr(34) & ", " & Chr(34) & "" & Chr(34) & ")"
' </VB WATCH>
1152   GetChatName = Replace(Str$, "-- Instant Message", "")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1153       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1154       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetPMName"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function lagg(TheText As String)
       'ex: call sendtext(lagg(UnSaKreD))
' <VB WATCH>
1155       On Error GoTo vbwErrHandler
1156       Const VBWPROCNAME = "Module1.lagg"
1157       If vbwTraceProc Then
1158           Dim vbwParameterString As String
1159           If vbwTraceParameters Then
1160               vbwParameterString = "(" & vbwReportParameter("TheText", TheText) & ") "
1161           End If
1162           vbwTraceIn VBWPROCNAME, vbwParameterString
1163       End If
' </VB WATCH>
1164   Dim G As String, a As String
1165   Dim W As Long
1166   Dim r$
1167   Dim U$
1168   Dim t$
1169   Dim p$
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1170   G$ = TheText"
' </VB WATCH>
1170   G$ = TheText
1171   Dim s$
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1172   a = Len(G$)"
' </VB WATCH>
1172   a = Len(G$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1173   For W = 1 To a Step 4"
' </VB WATCH>
1173   For W = 1 To a Step 4
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1174       r$ = Mid$(G$, W, 1)"
' </VB WATCH>
1174       r$ = Mid$(G$, W, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1175       U$ = Mid$(G$, W + 1, 1)"
' </VB WATCH>
1175       U$ = Mid$(G$, W + 1, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1176       s$ = Mid$(G$, W + 2, 1)"
' </VB WATCH>
1176       s$ = Mid$(G$, W + 2, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1177       t$ = Mid$(G$, W + 3, 1)"
' </VB WATCH>
1177       t$ = Mid$(G$, W + 3, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1178       p$ = p$ & " & Chr(34) & "<html></<html></html><html></html><html></html><html></html>" & Chr(34) & " & r$ & " & Chr(34) & "<html></<html></html><html></html><html></html><html></html>" & Chr(34) & " & U$ & " & Chr(34) & "<html></<html></html><html></html><html></html><html></html>" & Chr(34) & " & s$ & " & Chr(34) & "<html></<html></html><html></html><html></html><html></html>" & Chr(34) & " & t$"
' </VB WATCH>
1178       p$ = p$ & "<html></<html></html><html></html><html></html><html></html>" & r$ & "<html></<html></html><html></html><html></html><html></html>" & U$ & "<html></<html></html><html></html><html></html><html></html>" & s$ & "<html></<html></html><html></html><html></html><html></html>" & t$
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1179   Next W"
' </VB WATCH>
1179   Next W
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1180   lagg = p$"
' </VB WATCH>
1180   lagg = p$
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1181       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1182       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "lagg"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Sub Y_BudList_Caption(Caption$)
       'changes the caption of your buddylist
       'window
' <VB WATCH>
1183       On Error GoTo vbwErrHandler
1184       Const VBWPROCNAME = "Module1.Y_BudList_Caption"
1185       If vbwTraceProc Then
1186           Dim vbwParameterString As String
1187           If vbwTraceParameters Then
1188               vbwParameterString = "(" & vbwReportParameter("Caption$", Caption$) & ") "
1189           End If
1190           vbwTraceIn VBWPROCNAME, vbwParameterString
1191       End If
' </VB WATCH>
1192   Dim yahoobudlist As Long
1193   Dim SetCaption As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1194   yahoobudlist = FindWindow(" & Chr(34) & "YahooBuddyMain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1194   yahoobudlist = FindWindow("YahooBuddyMain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1195   setcaption = SendMessageByString(yahoobudlist, WM_SETTEXT, 0, Caption$)"
' </VB WATCH>
1195   SetCaption = SendMessageByString(yahoobudlist, WM_SETTEXT, 0, Caption$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1196       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1197       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Y_BudList_Caption"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Form_ExitDown(Form As Form)
       'Gives your form that cool flying down effect
' <VB WATCH>
1198       On Error GoTo vbwErrHandler
1199       Const VBWPROCNAME = "Module1.Form_ExitDown"
1200       If vbwTraceProc Then
1201           Dim vbwParameterString As String
1202           If vbwTraceParameters Then
1203               vbwParameterString = "(" & vbwReportParameter("Form", Form) & ") "
1204           End If
1205           vbwTraceIn VBWPROCNAME, vbwParameterString
1206       End If
' </VB WATCH>
' <VBW_LINE>Do Until Form.Top >= 13000
1207   Do Until vbwExecuteLine(False, "1207   Do Until Form.Top >= 13000") Or _
        Form.Top >= 13000
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1208   Form.Top = Trim(Str(Int(Form.Top) + 175))"
' </VB WATCH>
1208   Form.Top = Trim(Str(Int(Form.Top) + 175))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1209   Loop"
' </VB WATCH>
1209   Loop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1210   Unload Form"
' </VB WATCH>
1210   Unload Form
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1211       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1212       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_ExitDown"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Form_ExitColapse(Form As Form)
       'Colapses you form to the center if your screen
' <VB WATCH>
1213       On Error GoTo vbwErrHandler
1214       Const VBWPROCNAME = "Module1.Form_ExitColapse"
1215       If vbwTraceProc Then
1216           Dim vbwParameterString As String
1217           If vbwTraceParameters Then
1218               vbwParameterString = "(" & vbwReportParameter("Form", Form) & ") "
1219           End If
1220           vbwTraceIn VBWPROCNAME, vbwParameterString
1221       End If
' </VB WATCH>
1222   Dim Counter As Integer
1223   Dim i As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1224   counter = Form.Height"
' </VB WATCH>
1224   Counter = Form.Height
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1225   Do"
' </VB WATCH>
1225   Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1226   DoEvents"
' </VB WATCH>
1226   DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1227   counter = counter - 10"
' </VB WATCH>
1227   Counter = Counter - 10
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1228   Form.Height = counter"
' </VB WATCH>
1228   Form.Height = Counter
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1229   Form.Top = (Screen.Height - Form.Height) / 2"
' </VB WATCH>
1229   Form.Top = (Screen.Height - Form.Height) / 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1230   Loop Until counter <= 10"
' </VB WATCH>
1230   Loop Until Counter <= 10
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1231   i = 15"
' </VB WATCH>
1231   i = 15
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1232   counter = Form.Width"
' </VB WATCH>
1232   Counter = Form.Width
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1233   Do"
' </VB WATCH>
1233   Do
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1234   DoEvents"
' </VB WATCH>
1234   DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1235   counter = counter + i"
' </VB WATCH>
1235   Counter = Counter + i
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1236   Form.Width = counter"
' </VB WATCH>
1236   Form.Width = Counter
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1237   Form.Left = (Screen.Width - Form.Width) / 2"
' </VB WATCH>
1237   Form.Left = (Screen.Width - Form.Width) / 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1238   i = i + 1"
' </VB WATCH>
1238   i = i + 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1239   Loop Until counter >= Screen.Width"
' </VB WATCH>
1239   Loop Until Counter >= Screen.Width
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1240   Unload Form"
' </VB WATCH>
1240   Unload Form
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1241       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1242       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_ExitColapse"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Form_ExitRight(Form As Form)
       'Makes your form fly right
' <VB WATCH>
1243       On Error GoTo vbwErrHandler
1244       Const VBWPROCNAME = "Module1.Form_ExitRight"
1245       If vbwTraceProc Then
1246           Dim vbwParameterString As String
1247           If vbwTraceParameters Then
1248               vbwParameterString = "(" & vbwReportParameter("Form", Form) & ") "
1249           End If
1250           vbwTraceIn VBWPROCNAME, vbwParameterString
1251       End If
' </VB WATCH>
' <VBW_LINE>Do Until Form.Left >= 13000
1252   Do Until vbwExecuteLine(False, "1252   Do Until Form.Left >= 13000") Or _
        Form.Left >= 13000
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1253   Form.Left = Trim(Str(Int(Form.Left) + 175))"
' </VB WATCH>
1253   Form.Left = Trim(Str(Int(Form.Left) + 175))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1254   Loop"
' </VB WATCH>
1254   Loop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1255   Unload Form"
' </VB WATCH>
1255   Unload Form
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1256       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1257       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_ExitRight"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub NewChatSend(what$)
' <VB WATCH>
1258       On Error GoTo vbwErrHandler
1259       Const VBWPROCNAME = "Module1.NewChatSend"
1260       If vbwTraceProc Then
1261           Dim vbwParameterString As String
1262           If vbwTraceParameters Then
1263               vbwParameterString = "(" & vbwReportParameter("what$", what$) & ") "
1264           End If
1265           vbwTraceIn VBWPROCNAME, vbwParameterString
1266       End If
' </VB WATCH>
1267   Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1268   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1268   imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1269   RichEdit = FindWindowEx(imclass, 0&, " & Chr(34) & "richedit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1269   RICHEDIT = FindWindowEx(imclass, 0&, "richedit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1270   Call SendMessageByString(RichEdit, WM_SETTEXT, 0&, what$)"
' </VB WATCH>
1270   Call SendMessageByString(RICHEDIT, WM_SETTEXT, 0&, what$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1271   ClickSend"
' </VB WATCH>
1271   ClickSend
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1272       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1273       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NewChatSend"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ClickSend()
' <VB WATCH>
1274       On Error GoTo vbwErrHandler
1275       Const VBWPROCNAME = "Module1.ClickSend"
1276       If vbwTraceProc Then
1277           Dim vbwParameterString As String
1278           If vbwTraceParameters Then
1279               vbwParameterString = "()"
1280           End If
1281           vbwTraceIn VBWPROCNAME, vbwParameterString
1282       End If
' </VB WATCH>
1283   Dim imclass As Long, Button As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1284   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1284   imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1285   Button = FindWindowEx(imclass, 0&, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1285   Button = FindWindowEx(imclass, 0&, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1286   Button = FindWindowEx(imclass, Button, " & Chr(34) & "button" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1286   Button = FindWindowEx(imclass, Button, "button", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1287   Call SendMessageLong(Button, WM_KEYDOWN, VK_SPACE, 0&)"
' </VB WATCH>
1287   Call SendMessageLong(Button, WM_KEYDOWN, VK_SPACE, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1288   Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)"
' </VB WATCH>
1288   Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1289       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1290       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ClickSend"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub NewChatBoot(code$)
' <VB WATCH>
1291       On Error GoTo vbwErrHandler
1292       Const VBWPROCNAME = "Module1.NewChatBoot"
1293       If vbwTraceProc Then
1294           Dim vbwParameterString As String
1295           If vbwTraceParameters Then
1296               vbwParameterString = "(" & vbwReportParameter("Code$", code$) & ") "
1297           End If
1298           vbwTraceIn VBWPROCNAME, vbwParameterString
1299       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1300   Anti2"
' </VB WATCH>
1300   Anti2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1301   CloseWindow2"
' </VB WATCH>
1301   CloseWindow2
1302   Dim parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1303   parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1303   parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1304   If InStr(GetCaption(parent&), LCase(" & Chr(34) & "-- instant message" & Chr(34) & ")) Then"
' </VB WATCH>
1304   If InStr(GetCaption(parent&), LCase("-- instant message")) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1305"
' </VB WATCH>
1305
' <VB WATCH>
1306       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1307       Exit Sub"
' </VB WATCH>
1307       Exit Sub
1308   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1308   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1309   Call SetFocusApi(parent&)"
' </VB WATCH>
1309   Call SetFocusApi(parent&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1310   Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1310   Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1311   Call SendMessageByString(Child2&, WM_SETTEXT, 0, Code$)"
' </VB WATCH>
1311   Call SendMessageByString(Child2&, WM_SETTEXT, 0, code$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1312   Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
1312   Call ClickMenu(parent&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1313       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1314       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NewChatBoot"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub CloseWindow2()
' <VB WATCH>
1315       On Error GoTo vbwErrHandler
1316       Const VBWPROCNAME = "Module1.CloseWindow2"
1317       If vbwTraceProc Then
1318           Dim vbwParameterString As String
1319           If vbwTraceParameters Then
1320               vbwParameterString = "()"
1321           End If
1322           vbwTraceIn VBWPROCNAME, vbwParameterString
1323       End If
' </VB WATCH>
1324   Dim imclass As Long, RICHEDIT As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1325   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1325   imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1326   RichEdit = FindWindowEx(imclass, 0&, " & Chr(34) & "richedit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1326   RICHEDIT = FindWindowEx(imclass, 0&, "richedit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1327   RichEdit = FindWindowEx(imclass, RichEdit, " & Chr(34) & "richedit" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1327   RICHEDIT = FindWindowEx(imclass, RICHEDIT, "richedit", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1328   Call SendMessageLong(RichEdit, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
1328   Call SendMessageLong(RICHEDIT, WM_CLOSE, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1329       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1330       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CloseWindow2"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub voiceboot()
' <VB WATCH>
1331       On Error GoTo vbwErrHandler
1332       Const VBWPROCNAME = "Module1.voiceboot"
1333       If vbwTraceProc Then
1334           Dim vbwParameterString As String
1335           If vbwTraceParameters Then
1336               vbwParameterString = "()"
1337           End If
1338           vbwTraceIn VBWPROCNAME, vbwParameterString
1339       End If
' </VB WATCH>
1340   Dim yahoobuddymain As Long, parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1341   yahoobuddymain = FindWindow(" & Chr(34) & "yahoobuddymain" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1341   yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1342   Call ClickMenu(yahoobuddymain&, " & Chr(34) & "Send a &Message" & Chr(34) & ")"
' </VB WATCH>
1342   Call ClickMenu(yahoobuddymain&, "Send a &Message")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1343       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1344       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "voiceboot"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub


Sub NEWBOOT(code$)
' <VB WATCH>
1345       On Error GoTo vbwErrHandler
1346       Const VBWPROCNAME = "Module1.NEWBOOT"
1347       If vbwTraceProc Then
1348           Dim vbwParameterString As String
1349           If vbwTraceParameters Then
1350               vbwParameterString = "(" & vbwReportParameter("Code$", code$) & ") "
1351           End If
1352           vbwTraceIn VBWPROCNAME, vbwParameterString
1353       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1354   Anti2"
' </VB WATCH>
1354   Anti2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1355   NewAnti"
' </VB WATCH>
1355   NewAnti
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1356   ClosedaWindow"
' </VB WATCH>
1356   ClosedaWindow
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1357   Closeewindow"
' </VB WATCH>
1357   Closeewindow
1358   Dim parent As Long, Child1 As Long, Child2 As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1359   parent& = FindWindow(" & Chr(34) & "IMClass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1359   parent& = FindWindow("IMClass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1360   If InStr(GetCaption(parent&), LCase(" & Chr(34) & "-- instant message" & Chr(34) & ")) Then"
' </VB WATCH>
1360   If InStr(GetCaption(parent&), LCase("-- instant message")) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1361"
' </VB WATCH>
1361
' <VB WATCH>
1362       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1363       Exit Sub"
' </VB WATCH>
1363       Exit Sub
1364   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1364   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1365   Call SetFocusApi(parent&)"
' </VB WATCH>
1365   Call SetFocusApi(parent&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1366   Child2& = FindWindowEx(parent&, 0&, " & Chr(34) & "RICHEDIT" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1366   Child2& = FindWindowEx(parent&, 0&, "RICHEDIT", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1367   Call SendMessageByString(Child2&, WM_SETTEXT, 0, Code$)"
' </VB WATCH>
1367   Call SendMessageByString(Child2&, WM_SETTEXT, 0, code$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1368   Call ClickMenu(parent&, " & Chr(34) & "Sen&d" & Chr(34) & ")"
' </VB WATCH>
1368   Call ClickMenu(parent&, "Sen&d")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1369       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1370       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NEWBOOT"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub ClosedaWindow()
' <VB WATCH>
1371       On Error GoTo vbwErrHandler
1372       Const VBWPROCNAME = "Module1.ClosedaWindow"
1373       If vbwTraceProc Then
1374           Dim vbwParameterString As String
1375           If vbwTraceParameters Then
1376               vbwParameterString = "()"
1377           End If
1378           vbwTraceIn VBWPROCNAME, vbwParameterString
1379       End If
' </VB WATCH>
1380   Dim imclass As Long, atleeb As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1381   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1381   imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1382   atleeb = FindWindowEx(imclass, 0&, " & Chr(34) & "atl:004eeb20" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1382   atleeb = FindWindowEx(imclass, 0&, "atl:004eeb20", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1383   Call SendMessageLong(atleeb, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
1383   Call SendMessageLong(atleeb, WM_CLOSE, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1384       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1385       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ClosedaWindow"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub pmblock()
' <VB WATCH>
1386       On Error GoTo vbwErrHandler
1387       Const VBWPROCNAME = "Module1.pmblock"
1388       If vbwTraceProc Then
1389           Dim vbwParameterString As String
1390           If vbwTraceParameters Then
1391               vbwParameterString = "()"
1392           End If
1393           vbwTraceIn VBWPROCNAME, vbwParameterString
1394       End If
' </VB WATCH>
1395   Dim imclass As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1396   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1396   imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1397   Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
1397   Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1398       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1399       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "pmblock"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Function YGetName()
' <VB WATCH>
1400       On Error GoTo vbwErrHandler
1401       Const VBWPROCNAME = "Module1.YGetName"
1402       If vbwTraceProc Then
1403           Dim vbwParameterString As String
1404           If vbwTraceParameters Then
1405               vbwParameterString = "()"
1406           End If
1407           vbwTraceIn VBWPROCNAME, vbwParameterString
1408       End If
' </VB WATCH>
1409   Dim imclass As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1410   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1410   imclass = FindWindow("imclass", vbNullString)
1411   Dim TheText As String, TL As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1412   TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)"
' </VB WATCH>
1412   TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1413   TheText = String(TL + 1, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
1413   TheText = String(TL + 1, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1414   Call SendMessageByString(imclass, WM_gettext, TL + 1, TheText)"
' </VB WATCH>
1414   Call SendMessageByString(imclass, WM_GETTEXT, TL + 1, TheText)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1415   TheText = Left(TheText, TL)"
' </VB WATCH>
1415   TheText = Left(TheText, TL)
1416   Dim trimmed
1417   Dim lenght As Integer
1418   Dim Chat As String, Char As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1419   Chat = TheText"
' </VB WATCH>
1419   Chat = TheText
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1420   Char = InStr(Chat, " & Chr(34) & " -- " & Chr(34) & ")"
' </VB WATCH>
1420   Char = InStr(Chat, " -- ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1421   trimmed = Left(Chat, Char)"
' </VB WATCH>
1421   trimmed = Left(Chat, Char)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1422   trimmed = Trim(trimmed)"
' </VB WATCH>
1422   trimmed = Trim(trimmed)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1423   YGetName = trimmed"
' </VB WATCH>
1423   YGetName = trimmed
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1424       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1425       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "YGetName"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Public Sub Menu_Run(lngwindow As Long, strmenutext As String)
       'Runs Menus
       'Thank you unsakred
' <VB WATCH>
1426       On Error GoTo vbwErrHandler
1427       Const VBWPROCNAME = "Module1.Menu_Run"
1428       If vbwTraceProc Then
1429           Dim vbwParameterString As String
1430           If vbwTraceParameters Then
1431               vbwParameterString = "(" & vbwReportParameter("lngwindow", lngwindow) & ", "
1432               vbwParameterString = vbwParameterString & vbwReportParameter("strmenutext", strmenutext) & ") "
1433           End If
1434           vbwTraceIn VBWPROCNAME, vbwParameterString
1435       End If
' </VB WATCH>
1436   Dim intLoop As Integer, intSubLoop As Integer, intSub2Loop As Integer, intSub3Loop As Integer, intSub4Loop As Integer
1437   Dim lngmenu(1 To 5) As Long
1438   Dim lngcount(1 To 5) As Long
1439   Dim lngSubMenuID(1 To 4) As Long
1440   Dim strcaption(1 To 4) As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1441       lngmenu(1) = GetMenu(lngwindow&)"
' </VB WATCH>
1441       lngmenu(1) = GetMenu(lngwindow&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1442       lngcount(1) = GetMenuItemCount(lngmenu(1))"
' </VB WATCH>
1442       lngcount(1) = GetMenuItemCount(lngmenu(1))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1443           For intLoop% = 0 To lngcount(1) - 1"
' </VB WATCH>
1443           For intLoop% = 0 To lngcount(1) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1444               DoEvents"
' </VB WATCH>
1444               DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1445               lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)"
' </VB WATCH>
1445               lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1446               lngcount(2) = GetMenuItemCount(lngmenu(2))"
' </VB WATCH>
1446               lngcount(2) = GetMenuItemCount(lngmenu(2))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1447                   For intSubLoop% = 0 To lngcount(2) - 1"
' </VB WATCH>
1447                   For intSubLoop% = 0 To lngcount(2) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1448                       DoEvents"
' </VB WATCH>
1448                       DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1449                       lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)"
' </VB WATCH>
1449                       lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1450                       strcaption(1) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
1450                       strcaption(1) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1451                       Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)"
' </VB WATCH>
1451                       Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1452                           If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then"
' </VB WATCH>
1452                           If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1453                               Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)"
' </VB WATCH>
1453                               Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)
' <VB WATCH>
1454       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1455                               Exit Sub"
' </VB WATCH>
1455                               Exit Sub
1456                           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1456                           End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1457                       lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)"
' </VB WATCH>
1457                       lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1458                       lngcount(3) = GetMenuItemCount(lngmenu(3))"
' </VB WATCH>
1458                       lngcount(3) = GetMenuItemCount(lngmenu(3))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1459                           If lngcount(3) > 0 Then"
' </VB WATCH>
1459                           If lngcount(3) > 0 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1460                               For intSub2Loop% = 0 To lngcount(3) - 1"
' </VB WATCH>
1460                               For intSub2Loop% = 0 To lngcount(3) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1461                                   DoEvents"
' </VB WATCH>
1461                                   DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1462                                   lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
1462                                   lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1463                                   strcaption(2) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
1463                                   strcaption(2) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1464                                   Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)"
' </VB WATCH>
1464                                   Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1465                                       If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then"
' </VB WATCH>
1465                                       If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1466                                           Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)"
' </VB WATCH>
1466                                           Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)
' <VB WATCH>
1467       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1468                                           Exit Sub"
' </VB WATCH>
1468                                           Exit Sub
1469                                       End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1469                                       End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1470                                   lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
1470                                   lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1471                                   lngcount(4) = GetMenuItemCount(lngmenu(4))"
' </VB WATCH>
1471                                   lngcount(4) = GetMenuItemCount(lngmenu(4))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1472                                       If lngcount(4) > 0 Then"
' </VB WATCH>
1472                                       If lngcount(4) > 0 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1473                                           For intSub3Loop% = 0 To lngcount(4) - 1"
' </VB WATCH>
1473                                           For intSub3Loop% = 0 To lngcount(4) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1474                                               DoEvents"
' </VB WATCH>
1474                                               DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1475                                               lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
1475                                               lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1476                                               strcaption(3) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
1476                                               strcaption(3) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1477                                               Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)"
' </VB WATCH>
1477                                               Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1478                                                   If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then"
' </VB WATCH>
1478                                                   If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1479                                                       Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)"
' </VB WATCH>
1479                                                       Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)
' <VB WATCH>
1480       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1481                                                       Exit Sub"
' </VB WATCH>
1481                                                       Exit Sub
1482                                                   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1482                                                   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1483                                               lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
1483                                               lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1484                                               lngcount(5) = GetMenuItemCount(lngmenu(5))"
' </VB WATCH>
1484                                               lngcount(5) = GetMenuItemCount(lngmenu(5))
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1485                                                   If lngcount(5) > 0 Then"
' </VB WATCH>
1485                                                   If lngcount(5) > 0 Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1486                                                       For intSub4Loop% = 0 To lngcount(5) - 1"
' </VB WATCH>
1486                                                       For intSub4Loop% = 0 To lngcount(5) - 1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1487                                                           DoEvents"
' </VB WATCH>
1487                                                           DoEvents
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1488                                                           lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)"
' </VB WATCH>
1488                                                           lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1489                                                           strcaption(4) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
1489                                                           strcaption(4) = String(75, " ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1490                                                           Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)"
' </VB WATCH>
1490                                                           Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1491                                                               If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then"
' </VB WATCH>
1491                                                               If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1492                                                                   Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)"
' </VB WATCH>
1492                                                                   Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)
' <VB WATCH>
1493       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1494                                                                   Exit Sub"
' </VB WATCH>
1494                                                                   Exit Sub
1495                                                               End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1495                                                               End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1496                                                       Next intSub4Loop%"
' </VB WATCH>
1496                                                       Next intSub4Loop%
1497                                                   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1497                                                   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1498                                           Next intSub3Loop%"
' </VB WATCH>
1498                                           Next intSub3Loop%
1499                                       End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1499                                       End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1500                               Next intSub2Loop%"
' </VB WATCH>
1500                               Next intSub2Loop%
1501                           End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1501                           End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1502                   Next intSubLoop%"
' </VB WATCH>
1502                   Next intSubLoop%
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1503           Next intLoop%"
' </VB WATCH>
1503           Next intLoop%
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1504       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1505       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Menu_Run"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub Closeewindow()
' <VB WATCH>
1506       On Error GoTo vbwErrHandler
1507       Const VBWPROCNAME = "Module1.Closeewindow"
1508       If vbwTraceProc Then
1509           Dim vbwParameterString As String
1510           If vbwTraceParameters Then
1511               vbwParameterString = "()"
1512           End If
1513           vbwTraceIn VBWPROCNAME, vbwParameterString
1514       End If
' </VB WATCH>
1515   Dim imclass As Long, atlebb As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1516   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1516   imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1517   atlebb = FindWindowEx(imclass, 0&, " & Chr(34) & "atl:004ebb50" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1517   atlebb = FindWindowEx(imclass, 0&, "atl:004ebb50", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1518   Call SendMessageLong(atlebb, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
1518   Call SendMessageLong(atlebb, WM_CLOSE, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1519       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1520       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Closeewindow"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub FIRSTBETA5BOOT()
' <VB WATCH>
1521       On Error GoTo vbwErrHandler
1522       Const VBWPROCNAME = "Module1.FIRSTBETA5BOOT"
1523       If vbwTraceProc Then
1524           Dim vbwParameterString As String
1525           If vbwTraceParameters Then
1526               vbwParameterString = "()"
1527           End If
1528           vbwTraceIn VBWPROCNAME, vbwParameterString
1529       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1530   SendChat " & Chr(34) & "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c" & Chr(34) & ""
' </VB WATCH>
1530   SendChat "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1531   pause 0.1"
' </VB WATCH>
1531   Pause 0.1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1532   SendChat " & Chr(34) & "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c" & Chr(34) & ""
' </VB WATCH>
1532   SendChat "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1533   SendChat " & Chr(34) & "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c" & Chr(34) & ""
' </VB WATCH>
1533   SendChat "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1534   pause 0.1"
' </VB WATCH>
1534   Pause 0.1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1535   SendChat " & Chr(34) & "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c" & Chr(34) & ""
' </VB WATCH>
1535   SendChat "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1536   pause 0.1"
' </VB WATCH>
1536   Pause 0.1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1537   SendChat " & Chr(34) & "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c" & Chr(34) & ""
' </VB WATCH>
1537   SendChat "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1538   pause 0.1"
' </VB WATCH>
1538   Pause 0.1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1539   SendChat " & Chr(34) & "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c" & Chr(34) & ""
' </VB WATCH>
1539   SendChat "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1540   pause 0.1"
' </VB WATCH>
1540   Pause 0.1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1541   SendChat " & Chr(34) & "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c" & Chr(34) & ""
' </VB WATCH>
1541   SendChat "<fade=<snd=aux/aux><snd=aux/aux><snd=aux/aux><snd=nul\nul><snd=nul\nul><snd=con/con><snd=con/con><snd=c"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1542   pause 0.1"
' </VB WATCH>
1542   Pause 0.1
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1543       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1544       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FIRSTBETA5BOOT"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub FadeFormYellow(vForm As Form)
       'Example:
       'Private Sub Form_Paint()
       'FadeFormYellow Me
       'End Sub
' <VB WATCH>
1545       On Error GoTo vbwErrHandler
1546       Const VBWPROCNAME = "Module1.FadeFormYellow"
1547       If vbwTraceProc Then
1548           Dim vbwParameterString As String
1549           If vbwTraceParameters Then
1550               vbwParameterString = "(" & vbwReportParameter("vForm", vForm) & ") "
1551           End If
1552           vbwTraceIn VBWPROCNAME, vbwParameterString
1553       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1554       On Error Resume Next"
' </VB WATCH>
1554       On Error Resume Next
1555       Dim intLoop As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1556       vForm.DrawStyle = vbInsideSolid"
' </VB WATCH>
1556       vForm.DrawStyle = vbInsideSolid
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1557       vForm.DrawMode = vbCopyPen"
' </VB WATCH>
1557       vForm.DrawMode = vbCopyPen
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1558       vForm.ScaleMode = vbPixels"
' </VB WATCH>
1558       vForm.ScaleMode = vbPixels
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1559       vForm.DrawWidth = 2"
' </VB WATCH>
1559       vForm.DrawWidth = 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1560       vForm.ScaleHeight = 256"
' </VB WATCH>
1560       vForm.ScaleHeight = 256
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1561       For intLoop = 0 To 255"
' </VB WATCH>
1561       For intLoop = 0 To 255
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1562           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B"
' </VB WATCH>
1562           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1563       Next intLoop"
' </VB WATCH>
1563       Next intLoop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1564       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1565       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FadeFormYellow"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub FadeFormBlue(vForm As Form)
       'Example:
       'Private Sub Form_Paint()
       'FadeFormBlue Me
       'End Sub
' <VB WATCH>
1566       On Error GoTo vbwErrHandler
1567       Const VBWPROCNAME = "Module1.FadeFormBlue"
1568       If vbwTraceProc Then
1569           Dim vbwParameterString As String
1570           If vbwTraceParameters Then
1571               vbwParameterString = "(" & vbwReportParameter("vForm", vForm) & ") "
1572           End If
1573           vbwTraceIn VBWPROCNAME, vbwParameterString
1574       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1575       On Error Resume Next"
' </VB WATCH>
1575       On Error Resume Next
1576       Dim intLoop As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1577       vForm.DrawStyle = vbInsideSolid"
' </VB WATCH>
1577       vForm.DrawStyle = vbInsideSolid
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1578       vForm.DrawMode = vbCopyPen"
' </VB WATCH>
1578       vForm.DrawMode = vbCopyPen
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1579       vForm.ScaleMode = vbPixels"
' </VB WATCH>
1579       vForm.ScaleMode = vbPixels
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1580       vForm.DrawWidth = 2"
' </VB WATCH>
1580       vForm.DrawWidth = 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1581       vForm.ScaleHeight = 256"
' </VB WATCH>
1581       vForm.ScaleHeight = 256
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1582       For intLoop = 0 To 255"
' </VB WATCH>
1582       For intLoop = 0 To 255
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1583           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B"
' </VB WATCH>
1583           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1584       Next intLoop"
' </VB WATCH>
1584       Next intLoop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1585       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1586       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FadeFormBlue"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub FadeFormGrey(vForm As Form)
       'Example:
       'Private Sub Form_Paint()
       'FadeFormGrey Me
       'End Sub
' <VB WATCH>
1587       On Error GoTo vbwErrHandler
1588       Const VBWPROCNAME = "Module1.FadeFormGrey"
1589       If vbwTraceProc Then
1590           Dim vbwParameterString As String
1591           If vbwTraceParameters Then
1592               vbwParameterString = "(" & vbwReportParameter("vForm", vForm) & ") "
1593           End If
1594           vbwTraceIn VBWPROCNAME, vbwParameterString
1595       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1596       On Error Resume Next"
' </VB WATCH>
1596       On Error Resume Next
1597       Dim intLoop As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1598       vForm.DrawStyle = vbInsideSolid"
' </VB WATCH>
1598       vForm.DrawStyle = vbInsideSolid
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1599       vForm.DrawMode = vbCopyPen"
' </VB WATCH>
1599       vForm.DrawMode = vbCopyPen
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1600       vForm.ScaleMode = vbPixels"
' </VB WATCH>
1600       vForm.ScaleMode = vbPixels
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1601       vForm.DrawWidth = 2"
' </VB WATCH>
1601       vForm.DrawWidth = 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1602       vForm.ScaleHeight = 256"
' </VB WATCH>
1602       vForm.ScaleHeight = 256
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1603       For intLoop = 0 To 255"
' </VB WATCH>
1603       For intLoop = 0 To 255
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1604           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B"
' </VB WATCH>
1604           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1605       Next intLoop"
' </VB WATCH>
1605       Next intLoop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1606       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1607       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FadeFormGrey"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub FadeFormGreen(vForm As Form)
       'Example:
       'Private Sub Form_Paint()
       'FadeFormGreen Me
       'End Sub
' <VB WATCH>
1608       On Error GoTo vbwErrHandler
1609       Const VBWPROCNAME = "Module1.FadeFormGreen"
1610       If vbwTraceProc Then
1611           Dim vbwParameterString As String
1612           If vbwTraceParameters Then
1613               vbwParameterString = "(" & vbwReportParameter("vForm", vForm) & ") "
1614           End If
1615           vbwTraceIn VBWPROCNAME, vbwParameterString
1616       End If
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1617   On Error Resume Next"
' </VB WATCH>
1617   On Error Resume Next
1618       Dim intLoop As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1619       vForm.DrawStyle = vbInsideSolid"
' </VB WATCH>
1619       vForm.DrawStyle = vbInsideSolid
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1620       vForm.DrawMode = vbCopyPen"
' </VB WATCH>
1620       vForm.DrawMode = vbCopyPen
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1621       vForm.ScaleMode = vbPixels"
' </VB WATCH>
1621       vForm.ScaleMode = vbPixels
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1622       vForm.DrawWidth = 2"
' </VB WATCH>
1622       vForm.DrawWidth = 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1623       vForm.ScaleHeight = 256"
' </VB WATCH>
1623       vForm.ScaleHeight = 256
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1624       For intLoop = 0 To 255"
' </VB WATCH>
1624       For intLoop = 0 To 255
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1625           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B"
' </VB WATCH>
1625           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1626       Next intLoop"
' </VB WATCH>
1626       Next intLoop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1627       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1628       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FadeFormGreen"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub FadeFormRed(vForm As Form)
       'Example:
       'Private Sub Form_Paint()
       'FadeFormRed Me
       'End Sub
' <VB WATCH>
1629       On Error GoTo vbwErrHandler
1630       Const VBWPROCNAME = "Module1.FadeFormRed"
1631       If vbwTraceProc Then
1632           Dim vbwParameterString As String
1633           If vbwTraceParameters Then
1634               vbwParameterString = "(" & vbwReportParameter("vForm", vForm) & ") "
1635           End If
1636           vbwTraceIn VBWPROCNAME, vbwParameterString
1637       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1638       On Error Resume Next"
' </VB WATCH>
1638       On Error Resume Next
1639       Dim intLoop As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1640       vForm.DrawStyle = vbInsideSolid"
' </VB WATCH>
1640       vForm.DrawStyle = vbInsideSolid
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1641       vForm.DrawMode = vbCopyPen"
' </VB WATCH>
1641       vForm.DrawMode = vbCopyPen
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1642       vForm.ScaleMode = vbPixels"
' </VB WATCH>
1642       vForm.ScaleMode = vbPixels
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1643       vForm.DrawWidth = 2"
' </VB WATCH>
1643       vForm.DrawWidth = 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1644       vForm.ScaleHeight = 256"
' </VB WATCH>
1644       vForm.ScaleHeight = 256
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1645       For intLoop = 0 To 255"
' </VB WATCH>
1645       For intLoop = 0 To 255
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1646           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B"
' </VB WATCH>
1646           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1647       Next intLoop"
' </VB WATCH>
1647       Next intLoop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1648       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1649       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FadeFormRed"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub FadeFormPurple(vForm As Form)
       'Example:
       'Private Sub Form_Paint()
       'FadeFormPurple Me
       'End Sub
' <VB WATCH>
1650       On Error GoTo vbwErrHandler
1651       Const VBWPROCNAME = "Module1.FadeFormPurple"
1652       If vbwTraceProc Then
1653           Dim vbwParameterString As String
1654           If vbwTraceParameters Then
1655               vbwParameterString = "(" & vbwReportParameter("vForm", vForm) & ") "
1656           End If
1657           vbwTraceIn VBWPROCNAME, vbwParameterString
1658       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1659       On Error Resume Next"
' </VB WATCH>
1659       On Error Resume Next
1660       Dim intLoop As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1661       vForm.DrawStyle = vbInsideSolid"
' </VB WATCH>
1661       vForm.DrawStyle = vbInsideSolid
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1662       vForm.DrawMode = vbCopyPen"
' </VB WATCH>
1662       vForm.DrawMode = vbCopyPen
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1663       vForm.ScaleMode = vbPixels"
' </VB WATCH>
1663       vForm.ScaleMode = vbPixels
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1664       vForm.DrawWidth = 2"
' </VB WATCH>
1664       vForm.DrawWidth = 2
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1665       vForm.ScaleHeight = 256"
' </VB WATCH>
1665       vForm.ScaleHeight = 256
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1666       For intLoop = 0 To 255"
' </VB WATCH>
1666       For intLoop = 0 To 255
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1667           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B"
' </VB WATCH>
1667           vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1668       Next intLoop"
' </VB WATCH>
1668       Next intLoop
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1669       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1670       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FadeFormPurple"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Sub ClipboardCopy(Text As String)
       'Copies text to the clipboard
       'Call Clipboardcopy("NewText")
       'or possibly
       'Call Clipboardcopy(text1.text)
' <VB WATCH>
1671       On Error GoTo vbwErrHandler
1672       Const VBWPROCNAME = "Module1.ClipboardCopy"
1673       If vbwTraceProc Then
1674           Dim vbwParameterString As String
1675           If vbwTraceParameters Then
1676               vbwParameterString = "(" & vbwReportParameter("Text", Text) & ") "
1677           End If
1678           vbwTraceIn VBWPROCNAME, vbwParameterString
1679       End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1680   On Error GoTo Error"
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1681   Clipboard.Clear"
' </VB WATCH>
1681   Clipboard.Clear
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1682   Clipboard.SetText Text$"
' </VB WATCH>
1682   Clipboard.SetText Text$
' <VB WATCH>
1683       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1684   Exit Sub"
' </VB WATCH>
1684   Exit Sub
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1685   Error"
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1686   MsgBox Err.Description, vbExclamation, " & Chr(34) & "Error" & Chr(34) & ""
' </VB WATCH>
1686   MsgBox Err.Description, vbExclamation, "Error"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1687       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1688       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ClipboardCopy"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub


Public Function getcrypt(passwd As String)
' <VB WATCH>
1689       On Error GoTo vbwErrHandler
1690       Const VBWPROCNAME = "Module1.getcrypt"
1691       If vbwTraceProc Then
1692           Dim vbwParameterString As String
1693           If vbwTraceParameters Then
1694               vbwParameterString = "(" & vbwReportParameter("passwd", passwd) & ") "
1695           End If
1696           vbwTraceIn VBWPROCNAME, vbwParameterString
1697       End If
' </VB WATCH>

1698   Dim ts As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1699   ts = Space$(50)"
' </VB WATCH>
1699   ts = Space$(50)
1700   Dim X As Long
1701   Dim saltc As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1702   saltc = " & Chr(34) & "_2S43d5f" & Chr(34) & ""
' </VB WATCH>
1702   saltc = "_2S43d5f"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1703   X = venkymd5crypt(passwd, saltc, ts)"
' </VB WATCH>
1703   X = venkymd5crypt(passwd, saltc, ts)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1704   getcrypt = ts"
' </VB WATCH>
1704   getcrypt = ts
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1705       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1706       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "getcrypt"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Function Bot_Offender(Nam As String)
       'Makes fun of someone in a chat
       'Example:
       'Call Bot_Offender(Text1.Text)
' <VB WATCH>
1707       On Error GoTo vbwErrHandler
1708       Const VBWPROCNAME = "Module1.Bot_Offender"
1709       If vbwTraceProc Then
1710           Dim vbwParameterString As String
1711           If vbwTraceParameters Then
1712               vbwParameterString = "(" & vbwReportParameter("Nam", Nam) & ") "
1713           End If
1714           vbwTraceIn VBWPROCNAME, vbwParameterString
1715       End If
' </VB WATCH>
1716   Dim X As Integer, lcse As String, letr As String, dis As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1717   SendChat " & Chr(34) & "<b>Offender Bot: Todays Offender Is... " & Chr(34) & " + Nam$"
' </VB WATCH>
1717   SendChat "<b>Offender Bot: Todays Offender is... " + Nam$
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1718   pause (0.4)"
' </VB WATCH>
1718   Pause (0.4)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1719   For X = 1 To Len(Nam)"
' </VB WATCH>
1719   For X = 1 To Len(Nam)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1720   lcse$ = LCase(Nam)"
' </VB WATCH>
1720   lcse$ = LCase(Nam)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1721   letr$ = Mid(lcse$, X, 1)"
' </VB WATCH>
1721   letr$ = Mid(lcse$, X, 1)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1722   If letr$ = " & Chr(34) & "a" & Chr(34) & " Then"
' </VB WATCH>
1722   If letr$ = "a" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1723        Let dis$ = " & Chr(34) & "a-is for the animals your momma fucks" & Chr(34) & ""
' </VB WATCH>
1723        Let dis$ = "A - is for the animals your mum fucks"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1724        GoTo Dissem"
' </VB WATCH>
1724        GoTo Dissem
1725   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1725   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1726   If letr$ = " & Chr(34) & "b" & Chr(34) & " Then"
' </VB WATCH>
1726   If letr$ = "b" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1727        Let dis$ = " & Chr(34) & "b-is for all the boys you love" & Chr(34) & ""
' </VB WATCH>
1727        Let dis$ = "B - is for all the boys you love"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1728        GoTo Dissem"
' </VB WATCH>
1728        GoTo Dissem
1729   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1729   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1730   If letr$ = " & Chr(34) & "c" & Chr(34) & " Then"
' </VB WATCH>
1730   If letr$ = "c" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1731        Let dis$ = " & Chr(34) & "c-is for the cunt you are" & Chr(34) & ""
' </VB WATCH>
1731        Let dis$ = "C - is for the cunt you are"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1732        GoTo Dissem"
' </VB WATCH>
1732        GoTo Dissem
1733   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1733   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1734   If letr$ = " & Chr(34) & "d" & Chr(34) & " Then"
' </VB WATCH>
1734   If letr$ = "d" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1735        Let dis$ = " & Chr(34) & "d-is for all the times your dissed" & Chr(34) & ""
' </VB WATCH>
1735        Let dis$ = "D - is for all the times your dissed"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1736        GoTo Dissem"
' </VB WATCH>
1736        GoTo Dissem
1737   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1737   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1738   If letr$ = " & Chr(34) & "e" & Chr(34) & " Then"
' </VB WATCH>
1738   If letr$ = "e" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1739        Let dis$ = " & Chr(34) & "e-is for that egghead of yours" & Chr(34) & ""
' </VB WATCH>
1739        Let dis$ = "E - is for that egghead of yours"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1740        GoTo Dissem"
' </VB WATCH>
1740        GoTo Dissem
1741   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1741   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1742   If letr$ = " & Chr(34) & "f" & Chr(34) & " Then"
' </VB WATCH>
1742   If letr$ = "f" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1743        Let dis$ = " & Chr(34) & "f-is for the friday nights you stay home" & Chr(34) & ""
' </VB WATCH>
1743        Let dis$ = "F - is for the friday nights you stay home"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1744        GoTo Dissem"
' </VB WATCH>
1744        GoTo Dissem
1745   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1745   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1746   If letr$ = " & Chr(34) & "g" & Chr(34) & " Then"
' </VB WATCH>
1746   If letr$ = "g" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1747        Let dis$ = " & Chr(34) & "g-is for the girls who hate you" & Chr(34) & ""
' </VB WATCH>
1747        Let dis$ = "G - is for the girls who hate you"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1748        GoTo Dissem"
' </VB WATCH>
1748        GoTo Dissem
1749   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1749   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1750   If letr$ = " & Chr(34) & "h" & Chr(34) & " Then"
' </VB WATCH>
1750   If letr$ = "h" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1751        Let dis$ = " & Chr(34) & "h-is for the ho your momma is" & Chr(34) & ""
' </VB WATCH>
1751        Let dis$ = "H - is for the ho your mum is"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1752        GoTo Dissem"
' </VB WATCH>
1752        GoTo Dissem
1753   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1753   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1754   If letr$ = " & Chr(34) & "i" & Chr(34) & " Then"
' </VB WATCH>
1754   If letr$ = "i" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1755        Let dis$ = " & Chr(34) & "i-is for the idiotic dumbass you are" & Chr(34) & ""
' </VB WATCH>
1755        Let dis$ = "I - is for the idiotic piece of shit you are"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1756        GoTo Dissem"
' </VB WATCH>
1756        GoTo Dissem
1757   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1757   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1758   If letr$ = " & Chr(34) & "j" & Chr(34) & " Then"
' </VB WATCH>
1758   If letr$ = "j" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1759        Let dis$ = " & Chr(34) & "j-is for all the times you jerkoff to your dog" & Chr(34) & ""
' </VB WATCH>
1759        Let dis$ = "J - is for all the times you whack off while thinking about your dog"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1760        GoTo Dissem"
' </VB WATCH>
1760        GoTo Dissem
1761   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1761   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1762   If letr$ = " & Chr(34) & "k" & Chr(34) & " Then"
' </VB WATCH>
1762   If letr$ = "k" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1763        Let dis$ = " & Chr(34) & "k-is for you self esteem that the cool kids killed" & Chr(34) & ""
' </VB WATCH>
1763        Let dis$ = "K - is for your self esteem that the cool kids killed"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1764        GoTo Dissem"
' </VB WATCH>
1764        GoTo Dissem
1765   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1765   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1766   If letr$ = " & Chr(34) & "l" & Chr(34) & " Then"
' </VB WATCH>
1766   If letr$ = "l" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1767        Let dis$ = " & Chr(34) & "l-is for the lame ass you are" & Chr(34) & ""
' </VB WATCH>
1767        Let dis$ = "L - is for the lama's ass you fucked"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1768        GoTo Dissem"
' </VB WATCH>
1768        GoTo Dissem
1769   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1769   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1770   If letr$ = " & Chr(34) & "m" & Chr(34) & " Then"
' </VB WATCH>
1770   If letr$ = "m" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1771        Let dis$ = " & Chr(34) & "m-is for the many men you sucked" & Chr(34) & ""
' </VB WATCH>
1771        Let dis$ = "M - is for the many men you sucked off!"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1772        GoTo Dissem"
' </VB WATCH>
1772        GoTo Dissem
1773   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1773   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1774   If letr$ = " & Chr(34) & "n" & Chr(34) & " Then"
' </VB WATCH>
1774   If letr$ = "n" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1775        Let dis$ = " & Chr(34) & "n-is for the nights you spent alone" & Chr(34) & ""
' </VB WATCH>
1775        Let dis$ = "N - is for the nights you spent alone"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1776        GoTo Dissem"
' </VB WATCH>
1776        GoTo Dissem
1777   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1777   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1778   If letr$ = " & Chr(34) & "o" & Chr(34) & " Then"
' </VB WATCH>
1778   If letr$ = "o" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1779        Let dis$ = " & Chr(34) & "o-is for the sex operation you had" & Chr(34) & ""
' </VB WATCH>
1779        Let dis$ = "O - is for the sex change operation you had"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1780        GoTo Dissem"
' </VB WATCH>
1780        GoTo Dissem
1781   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1781   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1782   If letr$ = " & Chr(34) & "p" & Chr(34) & " Then"
' </VB WATCH>
1782   If letr$ = "p" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1783        Let dis$ = " & Chr(34) & "p-is for the times people p on you" & Chr(34) & ""
' </VB WATCH>
1783        Let dis$ = "P - is for the times you shafted yourself with a pole!"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1784        GoTo Dissem"
' </VB WATCH>
1784        GoTo Dissem
1785   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1785   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1786   If letr$ = " & Chr(34) & "q" & Chr(34) & " Then"
' </VB WATCH>
1786   If letr$ = "q" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1787        Let dis$ = " & Chr(34) & "q-is for the queer you are" & Chr(34) & ""
' </VB WATCH>
1787        Let dis$ = "Q - is for the queer you are"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1788        GoTo Dissem"
' </VB WATCH>
1788        GoTo Dissem
1789   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1789   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1790   If letr$ = " & Chr(34) & "r" & Chr(34) & " Then"
' </VB WATCH>
1790   If letr$ = "r" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1791        Let dis$ = " & Chr(34) & "r-is for all the times i raped your sister" & Chr(34) & ""
' </VB WATCH>
1791        Let dis$ = "R - is for your riggid teeth thats why you can't get a girlfriend"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1792        GoTo Dissem"
' </VB WATCH>
1792        GoTo Dissem
1793   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1793   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1794   If letr$ = " & Chr(34) & "s" & Chr(34) & " Then"
' </VB WATCH>
1794   If letr$ = "s" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1795        Let dis$ = " & Chr(34) & "s-is for the sex u get from ur dad" & Chr(34) & ""
' </VB WATCH>
1795        Let dis$ = "S - is for the sex you never get!"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1796        GoTo Dissem"
' </VB WATCH>
1796        GoTo Dissem
1797   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1797   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1798   If letr$ = " & Chr(34) & "t" & Chr(34) & " Then"
' </VB WATCH>
1798   If letr$ = "t" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1799        Let dis$ = " & Chr(34) & "t-is for the tits youll never see" & Chr(34) & ""
' </VB WATCH>
1799        Let dis$ = "T - is for the tits that you will never feel!"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1800        GoTo Dissem"
' </VB WATCH>
1800        GoTo Dissem
1801   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1801   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1802   If letr$ = " & Chr(34) & "u" & Chr(34) & " Then"
' </VB WATCH>
1802   If letr$ = "u" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1803        Let dis$ = " & Chr(34) & "u-is for your underwear hangin on the flagpole" & Chr(34) & ""
' </VB WATCH>
1803        Let dis$ = "U - is for your underwear hangin on the flagpole"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1804        GoTo Dissem"
' </VB WATCH>
1804        GoTo Dissem
1805   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1805   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1806   If letr$ = " & Chr(34) & "v" & Chr(34) & " Then"
' </VB WATCH>
1806   If letr$ = "v" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1807        Let dis$ = " & Chr(34) & "v-is for the victories you'll never have" & Chr(34) & ""
' </VB WATCH>
1807        Let dis$ = "V - is for the victories you'll never have"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1808        GoTo Dissem"
' </VB WATCH>
1808        GoTo Dissem
1809   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1809   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1810   If letr$ = " & Chr(34) & "w" & Chr(34) & " Then"
' </VB WATCH>
1810   If letr$ = "w" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1811        Let dis$ = " & Chr(34) & "w-is for the 400 pounds you wiegh" & Chr(34) & ""
' </VB WATCH>
1811        Let dis$ = "W - is for the amount of times you've waxed!"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1812        GoTo Dissem"
' </VB WATCH>
1812        GoTo Dissem
1813   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1813   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1814   If letr$ = " & Chr(34) & "x" & Chr(34) & " Then"
' </VB WATCH>
1814   If letr$ = "x" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1815        Let dis$ = " & Chr(34) & "x-is for all the lamers who" & Chr(34) & " & Chr(34) & " & Chr(34) & "[x]'ed" & Chr(34) & " & Chr(34) & " & Chr(34) & " you online" & Chr(34) & ""
' </VB WATCH>
1815        Let dis$ = "X - is for all the twats who" & Chr(34) & "[x]'ed" & Chr(34) & " you online"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1816        GoTo Dissem"
' </VB WATCH>
1816        GoTo Dissem
1817   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1817   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1818   If letr$ = " & Chr(34) & "y" & Chr(34) & " Then"
' </VB WATCH>
1818   If letr$ = "y" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1819        Let dis$ = " & Chr(34) & "y-is for the question of, y your even alive?" & Chr(34) & ""
' </VB WATCH>
1819        Let dis$ = "Y - is for why do you suck all them donkeys off!"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1820        GoTo Dissem"
' </VB WATCH>
1820        GoTo Dissem
1821   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1821   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1822   If letr$ = " & Chr(34) & "z" & Chr(34) & " Then"
' </VB WATCH>
1822   If letr$ = "z" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1823        Let dis$ = " & Chr(34) & "z-is for zero which is what you are" & Chr(34) & ""
' </VB WATCH>
1823        Let dis$ = "Z - is for zero which is what you are"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1824        GoTo Dissem"
' </VB WATCH>
1824        GoTo Dissem
1825   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1825   End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1826   If letr$ = " & Chr(34) & "1" & Chr(34) & " Then"
' </VB WATCH>
1826   If letr$ = "1" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1827        Let dis$ = " & Chr(34) & "1-is for how many inches your dick is" & Chr(34) & ""
' </VB WATCH>
1827        Let dis$ = "1 - is for how many inches your dick is"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1828        GoTo Dissem"
' </VB WATCH>
1828        GoTo Dissem
1829   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1829   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1830   If letr$ = " & Chr(34) & "2" & Chr(34) & " Then"
' </VB WATCH>
1830   If letr$ = "2" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1831        Let dis$ = " & Chr(34) & "2-is for the 2 dollars you make an hour" & Chr(34) & ""
' </VB WATCH>
1831        Let dis$ = "2 - is for the 2 pennies you make an hour"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1832        GoTo Dissem"
' </VB WATCH>
1832        GoTo Dissem
1833   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1833   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1834   If letr$ = " & Chr(34) & "3" & Chr(34) & " Then"
' </VB WATCH>
1834   If letr$ = "3" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1835        Let dis$ = " & Chr(34) & "3-is for the amount of men your girl takes at once" & Chr(34) & ""
' </VB WATCH>
1835        Let dis$ = "3 - is for the amount of men you take at once"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1836        GoTo Dissem"
' </VB WATCH>
1836        GoTo Dissem
1837   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1837   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1838   If letr$ = " & Chr(34) & "4" & Chr(34) & " Then"
' </VB WATCH>
1838   If letr$ = "4" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1839        Let dis$ = " & Chr(34) & "4-is for your mom bein a whore" & Chr(34) & ""
' </VB WATCH>
1839        Let dis$ = "4 - is for your mom bein a whore"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1840        GoTo Dissem"
' </VB WATCH>
1840        GoTo Dissem
1841   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1841   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1842   If letr$ = " & Chr(34) & "5" & Chr(34) & " Then"
' </VB WATCH>
1842   If letr$ = "5" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1843        Let dis$ = " & Chr(34) & "5-is for 5 times an hour you whack off" & Chr(34) & ""
' </VB WATCH>
1843        Let dis$ = "5 - is for 5 times an hour you whack off"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1844        GoTo Dissem"
' </VB WATCH>
1844        GoTo Dissem
1845   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1845   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1846   If letr$ = " & Chr(34) & "6" & Chr(34) & " Then"
' </VB WATCH>
1846   If letr$ = "6" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1847        Let dis$ = " & Chr(34) & "6-is for the years you been single" & Chr(34) & ""
' </VB WATCH>
1847        Let dis$ = "6 - is for the years you been single"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1848        GoTo Dissem"
' </VB WATCH>
1848        GoTo Dissem
1849   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1849   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1850   If letr$ = " & Chr(34) & "7" & Chr(34) & " Then"
' </VB WATCH>
1850   If letr$ = "7" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1851        Let dis$ = " & Chr(34) & "7-is for the times your girl cheated on you..with me" & Chr(34) & ""
' </VB WATCH>
1851        Let dis$ = "7 - is for the times your only girl cheated on you..with me"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1852        GoTo Dissem"
' </VB WATCH>
1852        GoTo Dissem
1853   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1853   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1854   If letr$ = " & Chr(34) & "8" & Chr(34) & " Then"
' </VB WATCH>
1854   If letr$ = "8" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1855        Let dis$ = " & Chr(34) & "8-is for how many people beat the hell outta you today" & Chr(34) & ""
' </VB WATCH>
1855        Let dis$ = "8 - is for how many people will beat the hell outta you today"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1856        GoTo Dissem"
' </VB WATCH>
1856        GoTo Dissem
1857   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1857   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1858   If letr$ = " & Chr(34) & "9" & Chr(34) & " Then"
' </VB WATCH>
1858   If letr$ = "9" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1859        Let dis$ = " & Chr(34) & "9-is for how many boyfriends your momma has" & Chr(34) & ""
' </VB WATCH>
1859        Let dis$ = "9 - is for how many boyfriends your momma has"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1860        GoTo Dissem"
' </VB WATCH>
1860        GoTo Dissem
1861   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1861   End If" 'B
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1862   If letr$ = " & Chr(34) & "0" & Chr(34) & " Then"
' </VB WATCH>
1862   If letr$ = "0" Then
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1863        Let dis$ = " & Chr(34) & "0-is for the amount of girls you get" & Chr(34) & ""
' </VB WATCH>
1863        Let dis$ = "0 - is for the amount of girls you get"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1864        GoTo Dissem"
' </VB WATCH>
1864        GoTo Dissem
1865   End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1865   End If" 'B
' </VB WATCH>

1866 Dissem:
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1867   Call SendChat(dis$)"
' </VB WATCH>
1867   Call SendChat(dis$)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1868   pause (0.4)"
' </VB WATCH>
1868   Pause (0.4)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1869   Next X"
' </VB WATCH>
1869   Next X
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
1870       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1871       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Bot_Offender"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Function"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Function
Sub Anti2()
' <VB WATCH>
1872       On Error GoTo vbwErrHandler
1873       Const VBWPROCNAME = "Module1.Anti2"
1874       If vbwTraceProc Then
1875           Dim vbwParameterString As String
1876           If vbwTraceParameters Then
1877               vbwParameterString = "()"
1878           End If
1879           vbwTraceIn VBWPROCNAME, vbwParameterString
1880       End If
' </VB WATCH>
1881   Dim imclass As Long, atlefb As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1882   imclass = FindWindow(" & Chr(34) & "imclass" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1882   imclass = FindWindow("imclass", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1883   atlefb = FindWindowEx(imclass, 0&, " & Chr(34) & "atl:004efb68" & Chr(34) & ", vbNullString)"
' </VB WATCH>
1883   atlefb = FindWindowEx(imclass, 0&, "atl:004efb68", vbNullString)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1884   Call SendMessageLong(atlefb, WM_CLOSE, 0&, 0&)"
' </VB WATCH>
1884   Call SendMessageLong(atlefb, WM_CLOSE, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1885       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1886       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Anti2"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub Log_Off_Current_User()
' <VB WATCH>
1887       On Error GoTo vbwErrHandler
1888       Const VBWPROCNAME = "Module1.Log_Off_Current_User"
1889       If vbwTraceProc Then
1890           Dim vbwParameterString As String
1891           If vbwTraceParameters Then
1892               vbwParameterString = "()"
1893           End If
1894           vbwTraceIn VBWPROCNAME, vbwParameterString
1895       End If
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1896   On Error GoTo Log_Off_Current_User_Error"
' </VB WATCH>
1896   On Error GoTo Log_Off_Current_User_Error

1897   Dim lngResult As Long

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1898       lngResult = ExitWindowsEx(EWX_FORCE Or EWX_LOGOFF, 0&)"
' </VB WATCH>
1898       lngResult = ExitWindowsEx(EWX_FORCE Or EWX_LOGOFF, 0&)



1899 Log_Off_Current_User_Exit:
' <VB WATCH>
1900       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1901       Exit Sub"
' </VB WATCH>
1901       Exit Sub

1902 Log_Off_Current_User_Error:
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "1903       Err.Raise Err.Number, " & Chr(34) & "Procedure:Log_Off_Current_User" & Chr(34) & ""
' </VB WATCH>
1903       Err.Raise Err.number, "Procedure:Log_Off_Current_User"

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
1904       If vbwTraceProc Then vbwTraceOut VBWPROCNAME
1905       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Log_Off_Current_User"
    Select Case MsgBox("Error " & Err.number & vbCrLf & _
                      Err.Description & vbCrLf & _
                     "Location: " & VBWPROJECT & "." & VBWMODULE & "." & VBWPROCEDURE & vbCrLf & _
                     "Line " & Erl, _
                     vbAbortRetryIgnore)
      Case vbAbort
          End
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
  End Select

If vbwTraceLine Then vbwExecuteLine False, "End Sub"
    If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
End Sub

 Public Sub PlaySound(strSound As String)

    Dim wFlags%
    
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    sndPlaySound strSound, wFlags%

End Sub


