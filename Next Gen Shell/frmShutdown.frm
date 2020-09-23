VERSION 5.00
Begin VB.Form frmShutdown 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label ShutDownNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&No!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   4200
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   2520
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label ShutDownYes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Yes!"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Sure You Want To Shut Down ? "
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
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   5760
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shutdown..."
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   5775
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShutDownYes.FontUnderline = False
ShutDownNo.FontUnderline = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub ShutDownNo_Click()
Unload Me
End Sub

Private Sub ShutDownNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShutDownNo.FontUnderline = True
End Sub

Private Sub ShutDownYes_Click()
SHShutDownDialog mhOwner
End Sub

Private Sub ShutDownYes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShutDownYes.FontUnderline = True

End Sub
