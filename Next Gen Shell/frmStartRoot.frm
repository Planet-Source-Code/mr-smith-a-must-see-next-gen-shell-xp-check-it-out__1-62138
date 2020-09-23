VERSION 5.00
Begin VB.Form frmStartRoot 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   FillColor       =   &H0000FF00&
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   5160
   End
   Begin VB.PictureBox picQIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   240
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.FileListBox QFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   240
      Pattern         =   "*.lnk"
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   5280
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Image power1 
      Height          =   390
      Left            =   3720
      Picture         =   "frmStartRoot.frx":0000
      Top             =   5800
      Width           =   390
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00404040&
      X1              =   2520
      X2              =   2520
      Y1              =   5520
      Y2              =   240
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Send To"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Recent Documents"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Links"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "- Log Off Next Gen -"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label lblCompName 
      BackStyle       =   0  'Transparent
      Caption         =   "[Computer Name]"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   3075
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4920
      TabIndex        =   11
      Top             =   240
      Width           =   195
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Programs"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmStartRoot.frx":0862
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shutdown"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblMore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "More.."
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmStartRoot.frx":0B6C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblQItem 
      BackStyle       =   0  'Transparent
      Caption         =   "[Caption]"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   840
      MouseIcon       =   "frmStartRoot.frx":0E76
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   10
      Left            =   3240
      MouseIcon       =   "frmStartRoot.frx":1180
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   9
      Left            =   3240
      MouseIcon       =   "frmStartRoot.frx":148A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Control Pannel"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   3240
      MouseIcon       =   "frmStartRoot.frx":1794
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   2880
      X2              =   5340
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   2880
      X2              =   5340
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "[User]"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image powerdown 
      Height          =   390
      Left            =   3720
      Picture         =   "frmStartRoot.frx":1A9E
      Top             =   5805
      Width           =   390
   End
   Begin VB.Image power2 
      Height          =   390
      Left            =   3720
      Picture         =   "frmStartRoot.frx":2300
      Top             =   5805
      Width           =   390
   End
End
Attribute VB_Name = "frmStartRoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private li As Long
Private lq As Long
Public childFrm As Form
Public GC As Boolean
Public cm As Boolean
Private mm As Boolean
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function SHRestartSystem Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal hIcon As Long, ByVal sDir As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
   
Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Integer, ByVal aBOOL As Integer) As Integer
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Integer) As Integer

Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

Private Sub Command1_Click()
Dim Handle As Long
Handle& = FindWindowA("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 1

End
End Sub

Private Sub Form_Load()


Me.Visible = False
DoEvents


Dim U As String
Dim z As String
RRRegion Me, 10
U = CurUserName$
    z = UCase(Left(U, 1))
    U = LCase(Right(U, Len(U) - 1))
    U = z & U
lblUser = U

U = GetWinComputerName
    z = UCase(Left(U, 1))
    U = LCase(Right(U, Len(U) - 1))
    U = z & U
    U = "on " & U
lblCompName = U

QFile.path = sDir(CSIDL_STARTMENU)

Dim i As Long
Dim X As Long

i = QFile.ListCount - 1

Do Until i < 0
Load picQIcon(X + 1)
Load lblQItem(X + 1)

picQIcon(X + 1).Visible = False
lblQItem(X + 1).Visible = False
picQIcon(X + 1).Top = picQIcon(X).Top + picQIcon(0).Height + 60
lblQItem(X + 1).Top = lblQItem(X).Top + picQIcon(0).Height + 60
picQIcon(X + 1).Left = picQIcon(0).Left
lblQItem(X + 1).Left = lblQItem(0).Left

picQIcon(X).Visible = True
lblQItem(X).Visible = True

lblQItem(X) = Left(QFile.List(i), Len(QFile.List(i)) - 4)
DrawStartIcon QFile.path & "\" & QFile.List(i), picQIcon(X)

X = X + 1
i = i - 1

If X = 8 Then
lblMore.Visible = True
Exit Do
End If

Loop

DoEvents
Me.Visible = True
DoEvents

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontUnderline = False
Label1.FontUnderline = False
power1.Picture = power2.Picture
If lq > -1 Then lblQItem(lq).FontUnderline = False
li = -1
lq = -1
If cm = True Then cm = False: childFrm.killChi: GC = False
mm = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
frmTaskbar.strt.Picture = frmTaskbar.start2.Picture
End Sub

Private Sub Label1_Click()
frmShutdown.Show

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontUnderline = True
power1.Picture = powerdown.Picture
End Sub

Private Sub Label2_Click()
Call Log_Off_Current_User


End Sub

Private Sub Label3_Click()
PopupMenu Form1.websites
End Sub

Private Sub Label4_Click()
If GC = True Then Unload childFrm: GC = False

Dim child As New frmStartSub

child.Left = Me.Width - 15
child.pth = (sDir(CSIDL_STARTMENU) & "\programs\")
child.LoadFiles

Set child.Par = Me

GC = True

Set childFrm = child
End Sub

Private Sub Label6_Click()


Dim Handle As Long
Handle& = FindWindowA("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 1
End
End Sub



Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontUnderline = True
End Sub

Private Sub Label7_Click()
frm2.Show

End Sub

Private Sub Label8_Click()
frm3.Show
End Sub

Private Sub lblItem_Click(Index As Integer)
Select Case Index
Case 1
Call Shell("C:\WINDOWS\EXPLORER.EXE /n,C:\My Documents", vbMaximizedFocus)

Case 5
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("control.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)
Case 4
 
    iTask = Shell("explorer.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

Case 9
ShowFindDialog

Case 10
Call ShowRunDialog
End Select
Me.Hide
End Sub

Private Sub lblItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = li Then Exit Sub
If li > -1 Then lblItem(li).FontUnderline = False
DoEvents
lblItem(Index).FontUnderline = True
li = Index
End Sub



Private Sub lblqitem_Click(Index As Integer)
ShellFile QFile.path & "\" & lblQItem(Index) & ".lnk"
Me.Hide
End Sub

Private Sub lblQItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = lq Then Exit Sub
If lq > -1 Then lblQItem(lq).FontUnderline = False
DoEvents
lblQItem(Index).FontUnderline = True
lq = Index
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLink.FontUnderline = True
End Sub

Public Sub killPar()
Unload Me
End Sub



Private Sub Text1_Click()


Dim Handle As Long
Handle& = FindWindowA("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 1
End
End Sub





Private Sub power1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
power1.Picture = powerdown.Picture
End Sub

Private Sub power1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
power1.Picture = power2.Picture
End Sub

Private Sub Timer1_Timer()
If Me.Visible = False Then mm = False
If GC = True Or mm = False Then Exit Sub
Dim X As Long, Y As Long
Dim e As Boolean
X = GetX * 15: Y = GetY * 15
If X < Me.Left Then e = True
If X > Me.Left + Me.Width Then e = True
If Y < Me.Top Then e = True
If Y > Me.Top + Me.Height Then e = True

If e = True Then Unload Me: mm = False
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

