VERSION 5.00
Begin VB.Form frmStartSub 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   9840
   ClientTop       =   2730
   ClientWidth     =   3015
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   240
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   2
      Top             =   -120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   540
      Left            =   1440
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   300
   End
   Begin VB.Shape shpborder 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   6015
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblCaption 
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
      Height          =   240
      Index           =   0
      Left            =   600
      MouseIcon       =   "frmStartSub.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2355
   End
End
Attribute VB_Name = "frmStartSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pth As String
Private lstI
Public hN As Long
Public C As Form
Public GC As Boolean
Public Par As Form
Private childFrm As Form
Public cm As Boolean
Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
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



Public Sub LoadFiles()
Dim i As Long
Dim X As Long
Dim z As String
Dim p As Long
Dim cc As Long

cc = 1

imgIcon(0).Top = imgIcon(0).Top - 240
lblCaption(0).Top = lblCaption(0).Top - 240

i = 1

If Right(pth, 1) <> "\" Then pth = pth & "\"

Dir1.path = pth
File1.path = pth

DoEvents

Do Until X = Dir1.ListCount

Load imgIcon(i)
Load lblCaption(i)

imgIcon(i).Visible = True
lblCaption(i).Visible = True

imgIcon(i).Top = imgIcon(i - 1).Top + imgIcon(i).Height
lblCaption(i).Top = lblCaption(i - 1).Top + imgIcon(i).Height
imgIcon(i).Left = imgIcon(i - 1).Left
lblCaption(i).Left = lblCaption(i - 1).Left

If imgIcon(i).Top > Screen.Height - imgIcon(i).Height - (32 * 15) Then
imgIcon(i).Top = imgIcon(1).Top
imgIcon(i).Left = imgIcon(1).Left + (cc * shpborder.Width)
lblCaption(i).Top = imgIcon(1).Top
lblCaption(i).Left = imgIcon(i).Left + 360
cc = cc + 1
Me.Top = 0
End If


z = Dir1.List(X)

p = Len(z)

Do Until Mid(z, p, 1) = "\"

p = p - 1

Loop

z = Right(z, Len(z) - p)

lblCaption(i) = z
lblCaption(i).ToolTipText = z
lblCaption(i).Tag = "F"

X = X + 1
i = i + 1

Loop
  
X = 0
  
Do Until X = File1.ListCount

Load imgIcon(i)
Load lblCaption(i)

imgIcon(i).Visible = True
lblCaption(i).Visible = True

imgIcon(i).Top = imgIcon(i - 1).Top + imgIcon(i).Height
lblCaption(i).Top = lblCaption(i - 1).Top + imgIcon(i).Height
imgIcon(i).Left = imgIcon(i - 1).Left
lblCaption(i).Left = lblCaption(i - 1).Left

If imgIcon(i).Top > Screen.Height - imgIcon(i).Height - (32 * 15) Then
imgIcon(i).Top = imgIcon(1).Top
imgIcon(i).Left = imgIcon(1).Left + (cc * shpborder.Width)
lblCaption(i).Top = imgIcon(1).Top
lblCaption(i).Left = imgIcon(i).Left + 360
cc = cc + 1
End If


DrawStartIcon pth & File1.List(X), picTemp, True

imgIcon(i).Picture = picTemp.Image

z = File1.List(X)

p = Len(z)

Do Until Mid(z, p, 1) = "."

p = p - 1

Loop

z = Left(z, p - 1)

lblCaption(i) = z
lblCaption(i).ToolTipText = z

X = X + 1
i = i + 1

Loop

If cc = 1 Then
Me.Height = 240 + (imgIcon.UBound * imgIcon(0).Height)
Else
Me.Height = 36 * imgIcon(0).Height + 240
End If
shpborder.Height = Me.Height - 20
shpborder.Width = shpborder.Width * cc
Me.Width = shpborder.Width
DoEvents

Me.Show



End Sub


Private Sub Form_Load()
RRRegion Me, 10
GC = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Par.cm = True
 cm = True
 
End Sub

Private Sub lblCaption_Click(Index As Integer)
If lblCaption(Index).Tag = "F" Then
If GC = True Then childFrm.Enabled = True
Dim child As New frmStartSub
child.pth = pth & lblCaption(Index)
child.Left = Me.Left + lblCaption(Index).Left + lblCaption(Index).Width
child.Top = lblCaption(Index).Top + Me.Top
child.LoadFiles

Set childFrm = child
Set childFrm.Par = Me
GC = True
Else
Dim X As String
X = (lblCaption(Index).Caption)

ShellFile pth & X & ".lnk"




DoEvents
Par.killPar
Unload Me
End If
End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCaption(lstI).FontUnderline = False

lblCaption(Index).FontUnderline = True
lstI = Index

Par.cm = True
 cm = True

End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub tmrCheck_Timer()
If C.Visible = False Then Unload Me: GC = False
End Sub

Private Sub tmrRem_Timer()
tmrCheck.Enabled = True
tmrRem.Enabled = False
End Sub


Public Function killPar()
Par.killPar
Unload Me
End Function

Public Sub killChi()
If GC = True Then childFrm.killChi
Unload Me
End Sub
  
