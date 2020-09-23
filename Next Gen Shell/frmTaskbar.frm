VERSION 5.00
Begin VB.Form frmTaskbar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      FillColor       =   &H0000FF00&
      FillStyle       =   5  'Downward Diagonal
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picSysTray 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4800
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   480
      Width           =   735
      Begin VB.PictureBox picTime 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H0000FF00&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Image imgTrayIcon 
         Height          =   255
         Index           =   0
         Left            =   -255
         Stretch         =   -1  'True
         Top             =   -50
         Width           =   255
      End
   End
   Begin VB.Timer tmrSysTrayUpdate 
      Interval        =   500
      Left            =   840
      Top             =   2160
   End
   Begin VB.ListBox lstHwndNames 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   1620
      TabIndex        =   4
      Top             =   1860
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstHwnd 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Top             =   1860
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstNames 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   1620
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstApps 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer tmrTaskUpdate 
      Interval        =   250
      Left            =   360
      Top             =   2160
   End
   Begin VB.Image imgSep 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmTaskbar.frx":0000
      Top             =   600
      Width           =   75
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "[TIME]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      FillColor       =   &H0000FF00&
      Height          =   435
      Left            =   0
      Top             =   0
      Width           =   1035
   End
   Begin VB.Image strt 
      Height          =   480
      Left            =   0
      Picture         =   "frmTaskbar.frx":0242
      Top             =   0
      Width           =   1560
   End
   Begin VB.Image startdown 
      Height          =   480
      Left            =   0
      Picture         =   "frmTaskbar.frx":2984
      Top             =   0
      Width           =   1560
   End
   Begin VB.Image start2 
      Height          =   480
      Left            =   0
      Picture         =   "frmTaskbar.frx":50C6
      Top             =   0
      Width           =   1560
   End
   Begin VB.Label lblTask 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Index           =   0
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image taskbarimg 
      Height          =   480
      Left            =   0
      Picture         =   "frmTaskbar.frx":7808
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FG As String
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Integer, ByVal aBOOL As Integer) As Integer
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Integer) As Integer

Private Sub Form_Load()

taskbarimg.Width = Screen.Width 'taskbar img = screen width
frmtext.Show
Dim Handle As Long
Handle& = FindWindowA("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 0


WindowPos Me, 1

 
With Me

.Width = Screen.Width
.Height = 15 * 32
.Left = 0
.Top = Screen.Height - (32 * 15)

End With


With shpborder

.Width = Me.Width
.Height = Me.Height
.Top = 0: .Left = 0

End With


With lblTime

.Left = Me.Width - 90 - .Width
.Top = Me.Height / 2 - .Height / 2

.Caption = Left(Time, 5)

End With

Load imgSep(1)

With imgSep(1)
.Visible = True
.Left = Label1.Width + 45
.Top = Me.Height / 2 - .Height / 2
End With


Load imgSep(2)

With imgSep(2)
.Visible = True
.Left = lblTime.Left - 90 - .Width
.Top = Me.Height / 2 - .Height / 2
End With



lblTask(0).Left = lblTask(0).Left - lblTask(0).Width - picIcon(0).Width
picIcon(0).Left = picIcon(0).Left - lblTask(0).Width - picIcon(0).Width


ListApps


Dim r As String


Call tmrSysTrayUpdate_Timer




End Sub

Private Sub Form_Unload(Cancel As Integer)
SetDesktopArea RF_FROMFULL
Call UnLoadTrayIconHandler
End Sub

Private Sub imgStart_Click()

With frmStartRoot
.Show
.Left = 0
.Top = Me.Top - .Height
End With

End Sub

Private Sub timerSystray_Timer()
SetWindowPos SysBox, 0, 0, 0, Me.picSysTray.ScaleWidth, Me.picSysTray.ScaleHeight, 0
End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
strt.Picture = startdown.Picture
With frmStartRoot
.Show
.Left = 0
.Top = Me.Top - .Height
End With
End Sub



Private Sub taskbarimg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
strt.Picture = start2.Picture
End Sub

Private Sub tmrSysTrayUpdate_Timer()
Dim ET As Long
Dim dtrLeft As Long

For ET = 1 To imgTrayIcon.Count - 1
    If imgTrayIcon(ET).Tag <> "skip" Then dtrLeft = dtrLeft + 300
Next ET

picSysTray.Width = dtrLeft + 100
picSysTray.Left = lblTime.Left - (picSysTray.Width + 180)
End Sub

Private Sub tmrTaskUpdate_Timer()
ListApps
lblTime.Caption = Left(Time, 5)
End Sub

Public Function ListApps()

On Error Resume Next

Dim i As Long, C As Long
Dim d As Long
Dim e As Boolean



Me.lstApps.Clear
Me.lstNames.Clear

fEnumWindows Me.lstApps

DoEvents

i = lstApps.ListCount - 1
C = lstApps.ListCount

Do Until i < 0

d = 0
e = False

'check if window allready has an entry
Do Until d = lstHwnd.ListCount
If lstHwnd.List(d) = lstApps.List(i) Then e = True: Exit Do
d = d + 1
Loop

'Add it if its not there

If e = False Then
Load lblTask(lblTask.UBound + 1)
lblTask(lblTask.UBound).Caption = lstNames.List(i)
lblTask(lblTask.UBound).Left = lblTask(lblTask.UBound - 1).Left + lblTask(picIcon.UBound).Width + picIcon(picIcon.UBound).Width + 30
lblTask(lblTask.UBound).ZOrder 0
lblTask(lblTask.UBound).Tag = lstApps.List(i)
lblTask(lblTask.UBound).Visible = True
Call DrawIcon(picIcon(picIcon.UBound).hdc, lstApps.List(i), 0, 0)
lstHwnd.AddItem lstApps.List(i)
lstHwndNames.AddItem lstNames.List(i)
End If

'Change the buttons text if the one on the form has changed

If e = True Then
C = 0
Do Until lblTask(C).Caption = lstHwndNames.List(d)
C = C + 1
Loop

lstHwndNames.List(d) = lstNames.List(i)
lblTask(C).Caption = lstHwndNames.List(d)

End If

i = i - 1

Loop


i = 0
d = lstApps.ListCount

'Now check top see if windows that we have on the list still exists

Do Until i >= lstHwnd.ListCount

C = 0
e = False

Do Until C = lstApps.ListCount

If lstHwnd.List(i) = lstApps.List(C) Then e = True: Exit Do
C = C + 1

Loop

If e = False And C <> 0 Then
C = 0

Do Until lblTask(C).Caption = lstHwndNames.List(i)
C = C + 1
If C > lblTask.UBound Then GoTo kill
Loop

RemTask C
DoEvents

lstHwnd.RemoveItem i
lstHwndNames.RemoveItem i

End If

kill:

i = i + 1

Loop

End Function

Public Function RemTask(i As Long)
Dim C As Long
C = i
Do Until C = lblTask.UBound
lblTask(C).Caption = lblTask(C + 1).Caption
lblTask(C).Tag = lblTask(C + 1).Tag
picIcon(C).Picture = Nothing
Call DrawIcon(picIcon(C).hdc, lblTask(C + 1).Tag, 0, 0)
C = C + 1
Loop
Unload lblTask(lblTask.UBound)
Unload picIcon(picIcon.UBound)
End Function

Public Sub DrawIcon(hdc As Long, hwnd As Long, X As Integer, Y As Integer)
Ico = GetIcon(hwnd)
DrawIconEx hdc, X, Y, Ico, 16, 16, 0, 0, DI_NORMAL
End Sub

Public Function GetIcon(hwnd As Long) As Long
Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICONSM)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICON)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
strt.Picture = start2.Picture
Dim i As Long
Do Until i = lblTask.UBound + 1
lblTask(i).FontUnderline = False
i = i + 1
Loop
End Sub

Private Sub lblTask_Click(Index As Integer)

If FG = lblTask(Index).Tag Then
SetFGWindow lblTask(Index).Tag, False
FG = 0
Else
SetFGWindow lblTask(Index).Tag, True
FG = lblTask(Index).Tag
End If
End Sub


Private Sub lblTask_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Do Until i = lblTask.UBound + 1
lblTask(i).FontUnderline = False
i = i + 1
Loop
lblTask(Index).FontUnderline = True
End Sub

Private Sub imgTrayicon_DblClick(Index As Integer)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
    msg = WM_LBUTTONDBLCLK
        On Error Resume Next
        Set ti = m_colTrayIcons(frmTaskbar.imgTrayIcon(Index).Tag)
        If Err.number = 0 Then
            ti.PostCallbackMessage msg
        Else
            Err.Clear
        End If
        Set ti = Nothing
End Sub

Private Sub imgTrayicon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
        msg = WM_MOUSEMOVE
        On Error Resume Next
        Set ti = m_colTrayIcons(frmTaskbar.imgTrayIcon(Index).Tag)
        If Err.number = 0 Then
            ti.PostCallbackMessage msg
        Else
            Err.Clear
        End If
        Set ti = Nothing
End Sub

Private Sub imgTrayicon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
 
     If Button = 1 Then
        msg = WM_LBUTTONDOWN
     ElseIf Button = 2 Then
        msg = WM_RBUTTONDOWN
     End If
       
    On Error Resume Next
    Set ti = m_colTrayIcons(frmTaskbar.imgTrayIcon(Index).Tag)
    If Err.number = 0 Then
        ti.PostCallbackMessage msg
    Else
        Err.Clear
    End If
    Set ti = Nothing
End Sub

Private Sub imgTrayicon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
     If Button = 1 Then
        msg = WM_LBUTTONUP
     ElseIf Button = 2 Then
        msg = WM_RBUTTONUP
     End If
       
    On Error Resume Next
    Set ti = m_colTrayIcons(frmTaskbar.imgTrayIcon(Index).Tag)
    If Err.number = 0 Then
        ti.PostCallbackMessage msg
    Else
        Err.Clear
    End If
    Set ti = Nothing
End Sub

Public Sub TaskBar(ByVal Enabled As Boolean)
 Dim EWindow As Integer
 Static TaskbarHWnd As Long
 Static IsTaskbarEnabled As Integer
 If Not Enabled Then
  TaskbarHWnd = FindWindow("Shell_traywnd", "")
  If TaskbarHWnd <> 0 Then
   EWindow = IsWindowEnabled(TaskbarHWnd)
   If EWindow = 1 Then IsTaskbarEnabled = EnableWindow(TaskbarHWnd, 0)
  End If
 Else
  If IsTaskbarEnabled = 0 Then IsTaskbarEnabled = EnableWindow(TaskbarHWnd, 1)
 End If
End Sub


