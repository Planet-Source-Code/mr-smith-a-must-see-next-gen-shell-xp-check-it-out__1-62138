Attribute VB_Name = "modSystray"
Public Const WM_SETHOTKEY = &H32
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pNID As NOTIFYICONDATA) As Boolean




Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4

End Enum
Public Type NOTIFYICONDATA
       cbSize As Long 'Size of this Data Type
       hwnd As Long 'Visual output
       uID As Long
       uFlags As Long ' Various Command Parameters\Flags to sent to Api
       uCallbackMessage As Long
       hIcon As Long 'Where the icon is store.
       szTip As String * 64 'Tool TIp Text
      End Type
Public nidProgramData As NOTIFYICONDATA

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long



Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public parent As Long
Public SysBox As Long
Public Sub BootUpSysTray()
    Dim hwnd As Long, rctemp As RECT
    
    hwnd = FindWindow("Shell_TrayWnd", vbNullString)
    hwnd = FindWindowEx(hwnd, 0, "TrayNotifyWnd", vbNullString)
    SysBox = hwnd
    parent = GetParent(SysBox)
    SetParent SysBox, frmTaskbar.picSysTray.hwnd
    SetWindowPos SysBox, 0, 0, 0, 150, 100, 0
    
End Sub
