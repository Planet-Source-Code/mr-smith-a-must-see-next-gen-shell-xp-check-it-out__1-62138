Attribute VB_Name = "modTIHandler"
Option Explicit

Public Const WS_POPUP = &H80000000
Public Const WS_EX_TOPMOST = &H8&
Public Const MFT_STRING = &H0&
Public Const MFT_RADIOCHECK = &H200&
Public Const WC_SYSTRAY As String = "Shell_TrayWnd"
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const HWND_BROADCAST = &HFFFF&

Public Type WNDCLASSEX
  cbSize As Long
  Style As Long
  lpfnWndProc As Long
  cbClsExtra As Long
  cbWndExtra As Long
  hInstance As Long
  hIcon As Long
  hCursor As Long
  hbrBackground As Long
  lpszMenuName As String
  lpszClassName As String
  hIconSm As Long
End Type

Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
                
Public m_hSysTray As Long
Public m_fLog As Boolean
Public m_colTrayIcons As Collection
Public WM_TASKBARCREATED As Long
Public Con As Integer
Public stiLeft As Long
Public KLO As Boolean
Private Const WM_COPYDATA = &H4A

Private Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
           
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Static cds As COPYDATASTRUCT
  If uMsg = WM_COPYDATA Then
    MoveMemory cds, ByVal lParam, Len(cds)
    If cds.dwData = 1 Then
      WindowProc = modTIHandler.TrayIconHandler(cds.lpData)
      Exit Function
    End If
  End If
  
  WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
  
End Function

Public Function FuncPtr(ByVal pfn As Long) As Long
  FuncPtr = pfn
End Function

Public Sub LoadTrayIconHandler()
 Dim wcx As WNDCLASSEX
  Dim lRet As Long
  
  Con = 1
  KLO = False
  WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated")
  
  With wcx
    .cbSize = Len(wcx)
    .lpfnWndProc = FuncPtr(AddressOf WindowProc)
    .hInstance = App.hInstance
    .lpszClassName = WC_SYSTRAY
  End With
  
  Call RegisterClassEx(wcx)
  
  m_hSysTray = CreateWindowEx(WS_EX_TOPMOST, WC_SYSTRAY, vbNullString, WS_POPUP, _
    0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

  Set m_colTrayIcons = New Collection
  m_fLog = False
    
    For lRet = 1 To m_colTrayIcons.Count
      m_colTrayIcons.Remove 1
    Next
 
    Call SendMessage(HWND_BROADCAST, WM_TASKBARCREATED, 0&, ByVal 0&)
  
End Sub

Public Sub UnLoadTrayIconHandler()

  ' destroy systray window ...
  Call DestroyWindow(m_hSysTray)
  
  ' ... and unregister the window class
  Call UnregisterClass(WC_SYSTRAY, App.hInstance)
  
  ' free icon collection
  Set m_colTrayIcons = Nothing

End Sub
Public Function TrayIconHandler(ByVal lpIconData As Long) As Long
  
  Dim nid As NOTIFYICONDATA
  Dim ti As CTrayIcon
  Dim dwMessage As Long
  Dim sKey As String
  
  ' The NIM_ message starts 4 bytes after lpIconData
  MoveMemory dwMessage, ByVal lpIconData + 4, Len(dwMessage)
  ' The NOTIFYICONDATA struct starts 8 bytes after lpIconData
  MoveMemory nid, ByVal lpIconData + 8, Len(nid)

  sKey = KeyFromIcon(nid.hwnd, nid.uID)
  
  On Error Resume Next
  Dim Ol As Long
  Select Case dwMessage
    Case NIM_ADD
      
      Set ti = New CTrayIcon
      ti.ModifyFromNID lpIconData + 8
      m_colTrayIcons.Add ti, sKey
      
      With ti
        '//--Softworld Code 2000-08-12
            If KLO = False Then stiLeft = frmTaskbar.imgTrayIcon(Con - 1).Left + frmTaskbar.imgTrayIcon(Con - 1).Width + 40
            KLO = False
            Load frmTaskbar.imgTrayIcon(Con)
            
            frmTaskbar.imgTrayIcon(Con).Picture = .VBIcon
            frmTaskbar.imgTrayIcon(Con).Top = 40
            frmTaskbar.imgTrayIcon(Con).Left = stiLeft
            frmTaskbar.imgTrayIcon(Con).Width = frmTaskbar.imgTrayIcon(0).Width
            frmTaskbar.imgTrayIcon(Con).Height = frmTaskbar.imgTrayIcon(0).Height
            frmTaskbar.imgTrayIcon(Con).Visible = True
            frmTaskbar.imgTrayIcon(Con).Tag = sKey
            frmTaskbar.imgTrayIcon(Con).ToolTipText = .ToolTipText
            Con = Con + 1
        '//--
      End With
      
    Case NIM_MODIFY
      
      Set ti = m_colTrayIcons(sKey)
      
      With ti
        .ModifyFromNID lpIconData + 8
      '//--Softworld Code
      
      For Ol = 1 To frmTaskbar.imgTrayIcon.Count - 1
        If frmTaskbar.imgTrayIcon(Ol).Tag = sKey Then
            frmTaskbar.imgTrayIcon(Ol).Picture = .VBIcon
            Exit For
        End If
      Next Ol
    
      '//--
      End With
      
    Case NIM_DELETE
      
      m_colTrayIcons.Remove sKey
      '//--Softworld Code
      
      For Ol = 1 To frmTaskbar.imgTrayIcon.Count - 1
        If frmTaskbar.imgTrayIcon(Ol).Tag = sKey Then
            frmTaskbar.imgTrayIcon(Ol).Tag = "skip"
           
            frmTaskbar.imgTrayIcon(Ol).Visible = False
            Call FixTrayIcons
            
        Exit For
        End If
      Next Ol
    
      '//--
  End Select
  
  Set ti = Nothing

  TrayIconHandler = 1

End Function

Private Sub FixTrayIcons()
'//--Softworld Code
Dim Lo As Long
Dim Asa As Long
For Lo = 1 To frmTaskbar.imgTrayIcon.Count - 1
    If frmTaskbar.imgTrayIcon(Lo).Tag <> "skip" Then
       
        frmTaskbar.imgTrayIcon(Lo).Left = 40 + Asa
        Asa = Asa + frmTaskbar.imgTrayIcon(0).Width + 40
    End If
Next Lo
For Lo = frmTaskbar.imgTrayIcon.Count - 1 To 1 Step -1
    If frmTaskbar.imgTrayIcon(Lo).Tag <> "skip" Then
        stiLeft = frmTaskbar.imgTrayIcon(Lo).Left + frmTaskbar.imgTrayIcon(Lo).Width + 40
    Exit For
    End If
Next Lo
KLO = True
End Sub

Private Function KeyFromIcon(ByVal hOwner As Long, ByVal ID As Long) As String
  KeyFromIcon = "K" & Hex$(hOwner) & "-" & Trim$(Str$(ID))
End Function


