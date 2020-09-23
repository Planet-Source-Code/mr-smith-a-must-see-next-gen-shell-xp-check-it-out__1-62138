VERSION 5.00
Begin VB.Form frmtext 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Left            =   3480
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Left            =   3480
      Top             =   3000
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "FORMTEXT.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   10
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "FORMTEXT.frx":0CCA
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "FORMTEXT.frx":1994
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "FORMTEXT.frx":265E
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   2520
   End
   Begin VB.PictureBox picQIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.FileListBox QFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   60
      Pattern         =   "*.lnk"
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TEXT"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Explorer"
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
      Left            =   720
      TabIndex        =   11
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Recycle Bin"
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
      Left            =   720
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "My Documents"
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
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "My Computer"
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
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
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
      Left            =   1080
      MouseIcon       =   "FORMTEXT.frx":3328
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblQItem 
      BackColor       =   &H00404040&
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
      MouseIcon       =   "FORMTEXT.frx":3632
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Image imgWall 
      Height          =   6015
      Left            =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmtext"
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
Dim cnt As Integer

Dim Tex As String
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000



Private Sub Form_Click()
'If you want more options then remove the quotation mark
'below. Then click the form and a menu will show


imgWall.Width = Me.Width
imgWall.Height = Me.Height
End Sub

Private Sub Form_DblClick()
Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3", 1)
End Sub


Private Sub Form_Load()
Tex = vbLf & "Welcome To Next Gen Shell By - Shifty "


Me.Visible = False
DoEvents
QFile.path = sDir(CSIDL_desktop)


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




Private Sub Image1_Click()

End Sub

Private Sub imgWall_Click()
PopupMenu Form1.mnuWin
End Sub

Private Sub Label2_Click()
Call Shell("C:\WINDOWS\EXPLORER.EXE /n,C:\My Documents", vbMaximizedFocus)

End Sub

Private Sub Label3_Click()
Call ShowRecycleBin

End Sub

Private Sub Label4_Click()
ShellExecute 0&, vbNullString, "www.yahoo.com", vbNullString, _
      vbNullString, SW_SHOWNORMAL

End Sub

Private Sub Label5_Click()
 iTask = Shell("explorer.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub



Private Sub Label6_Click()
matrix.Show
End Sub


Private Sub lblQItem_DblClick(Index As Integer)
ShellFile QFile.path & "\" & lblQItem(Index) & ".lnk"
End Sub

Private Sub lblQItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = lq Then Exit Sub
If lq > -1 Then lblQItem(lq).FontUnderline = False
DoEvents
lblQItem(Index).FontUnderline = True
lq = Index
End Sub

Public Sub killPar()
Unload Me
End Sub

Private Sub Picture1_DblClick()
 iTask = Shell("explorer.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragObj Picture1


Label5.Move Picture1.Left + Picture1.Width + 30, Picture1.Top + 100

End Sub



Private Sub Picture2_DblClick()
Call Shell("C:\WINDOWS\EXPLORER.EXE /n,C:\My Documents", vbMaximizedFocus)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragObj Picture2
Label2.Move Picture2.Left + Picture2.Width + 30, Picture2.Top + 100
End Sub



Private Sub Picture3_DblClick()
Call ShowRecycleBin

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragObj Picture3
Label3.Move Picture3.Left + Picture3.Width + 30, Picture3.Top + 100
End Sub

Private Sub Picture4_DblClick()
ShellExecute 0&, vbNullString, "www.yahoo.com", vbNullString, _
      vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragObj Picture4
Label4.Move Picture4.Left + Picture4.Width + 30, Picture4.Top + 100
End Sub

Private Sub Timer1_Timer()
Label1.Caption = TextScroll(Tex)
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

Private Function TextScroll(ByVal MyText As String) As String
Static Tx
Tx = Space(Len(MyText)) & MyText
cnt = cnt + 1
MyText = Mid$(Tx, cnt, Len(MyText))
If Len(MyText) = 0 Then cnt = 0
TextScroll = MyText
End Function

Private Sub Timer2_Timer()
Label1.Caption = TextScroll(Tex)

End Sub

