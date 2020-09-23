VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2400
      Top             =   4440
   End
   Begin VB.PictureBox picQIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   480
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.FileListBox QFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   360
      Pattern         =   "*.lnk"
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   6135
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3255
   End
   Begin VB.Line Line6 
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   6120
   End
   Begin VB.Line Line5 
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   6240
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
      Left            =   600
      MouseIcon       =   "FRM2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5400
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
      Left            =   1440
      MouseIcon       =   "FRM2.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frm2"
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

QFile.path = sDir(CSIDL_RECENT)

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

If X = 12 Then
lblMore.Visible = True
Exit Do
End If

Loop

DoEvents
Me.Visible = True
DoEvents

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lq > -1 Then lblQItem(lq).FontUnderline = False
li = -1
lq = -1
If cm = True Then cm = False: childFrm.killChi: GC = False
mm = True
End Sub

Private Sub lblItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = li Then Exit Sub

DoEvents

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

