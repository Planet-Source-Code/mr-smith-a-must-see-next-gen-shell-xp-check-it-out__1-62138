VERSION 5.00
Begin VB.Form matrix 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   4200
      Top             =   5520
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   1080
      Top             =   5640
   End
   Begin VB.Timer Timer6 
      Interval        =   50
      Left            =   4080
      Top             =   4800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   14640
      TabIndex        =   43
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer5 
      Left            =   6480
      Top             =   7680
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   7200
      Top             =   7200
   End
   Begin VB.Timer Timer3 
      Interval        =   400
      Left            =   5160
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4200
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3360
      Top             =   3720
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   14520
      TabIndex        =   49
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   14880
      TabIndex        =   47
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   12000
      TabIndex        =   39
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   12360
      TabIndex        =   38
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   11640
      TabIndex        =   37
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label35 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Height          =   4455
      Left            =   15960
      TabIndex        =   34
      Top             =   10200
      Width           =   3255
   End
   Begin VB.Label Label34 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Height          =   3255
      Left            =   15960
      TabIndex        =   33
      Top             =   12120
      Width           =   2655
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   19080
      TabIndex        =   32
      Top             =   15240
      Width           =   375
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   10920
      TabIndex        =   31
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   11280
      TabIndex        =   30
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   8280
      TabIndex        =   29
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   8640
      TabIndex        =   28
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   9360
      TabIndex        =   27
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   9000
      TabIndex        =   26
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   7560
      TabIndex        =   25
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   7920
      TabIndex        =   24
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   7200
      TabIndex        =   23
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   6480
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   10080
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   10560
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   9720
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   5760
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   6840
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   6120
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   3240
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   3960
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   3600
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   14160
      TabIndex        =   36
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   13800
      TabIndex        =   35
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   13440
      TabIndex        =   41
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   13080
      TabIndex        =   40
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   12720
      TabIndex        =   42
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label55 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   18480
      TabIndex        =   55
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   16680
      TabIndex        =   54
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   17040
      TabIndex        =   53
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   17400
      TabIndex        =   52
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   17760
      TabIndex        =   51
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   18120
      TabIndex        =   50
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   15240
      TabIndex        =   48
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   15600
      TabIndex        =   46
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   16320
      TabIndex        =   45
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   15135
      Left            =   15960
      TabIndex        =   44
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
matrix.Hide
End Sub





Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
matrix.Hide
End Sub



Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
matrix.Hide
End Sub



Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
matrix.Hide
End Sub

Private Sub Label33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Visible = False
matrix.Hide
End Sub

Private Sub Label34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Visible = True
matrix.Hide
End Sub
Private Sub Label35_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Visible = True
matrix.Hide
End Sub



Private Sub Timer1_Timer()
    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label1.Caption = Label1.Caption + "1 "
        Case 2
         Label1.Caption = Label1.Caption + "0 "
        Case 3
         Label1.Caption = Label1.Caption + "1 "
        Case 4
         Label1.Caption = Label1.Caption + "0 "

        Case 6
         Label1.Caption = Label1.Caption + "0 "
        Case 7
         Label1.Caption = Label1.Caption + "1 "
        Case 8
         Label1.Caption = Label1.Caption + "0 "
        Case 9
         Label1.Caption = Label1.Caption + "1 "
        Case 10
         Label1.Caption = Label1.Caption + "0 "
        Case 11
         Label1.Caption = Label1.Caption + "1 "
        Case 12
         Label1.Caption = Label1.Caption + "0 "
        Case 13
         Label1.Caption = Label1.Caption + "1 "
        Case 14
         Label1.Caption = Label1.Caption + "0 "
        Case 15
         Label2.Caption = Label2.Caption + "1 "
        Case 16
         Label1.Caption = Label1.Caption + "0 "
        Case 17
         Label1.Caption = Label1.Caption + "1 "
        Case 18
         Label1.Caption = Label1.Caption + "0 "
        Case 19
         Label1.Caption = Label1.Caption + "1 "
        Case 20
         Label1.Caption = Label1.Caption + "0 "
    End Select
    
    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label2.Caption = Label2.Caption + "1 "
        Case 2
         Label2.Caption = Label2.Caption
        Case 3
         Label2.Caption = Label2.Caption + "1 "
        Case 4
         Label2.Caption = Label2.Caption + "0 "

        Case 6
         Label2.Caption = Label2.Caption + "0 "
        Case 7
         Label2.Caption = Label2.Caption + "1 "
        Case 8
         Label2.Caption = Label2.Caption + "0 "
        Case 9
         Label2.Caption = Label2.Caption
        Case 10
         Label2.Caption = Label2.Caption + "0 "
        Case 11
         Label2.Caption = Label2.Caption + "1 "
        Case 12
         Label2.Caption = Label2.Caption + "0 "
        Case 13
         Label2.Caption = Label2.Caption
        Case 14
         Label2.Caption = Label2.Caption + "0 "
        Case 15
         Label2.Caption = Label2.Caption + "1 "
        Case 16
         Label2.Caption = Label2.Caption + "0 "
        Case 17
         Label2.Caption = Label2.Caption + "1 "
        Case 18
         Label2.Caption = Label2.Caption
        Case 19
         Label2.Caption = Label2.Caption + "1 "
        Case 20
         Label2.Caption = Label2.Caption
    End Select
    
    
      Select Case Int((Rnd * 30) + 1)
        Case 1
         Label3.Caption = Label3.Caption + "1 "
        Case 2
         Label3.Caption = Label3.Caption
        Case 3
         Label3.Caption = Label3.Caption + "1 "
        Case 4
         Label3.Caption = Label3.Caption + "0 "

        Case 6
         Label3.Caption = Label3.Caption + "0 "
        Case 7
         Label3.Caption = Label3.Caption + "1 "
        Case 8
         Label3.Caption = Label3.Caption + "0 "
        Case 9
         Label3.Caption = Label3.Caption
        Case 10
         Label3.Caption = Label3.Caption + "0 "
        Case 11
         Label3.Caption = Label3.Caption + "1 "
        Case 12
         Label3.Caption = Label3.Caption + "0 "
        Case 13
         Label3.Caption = Label3.Caption
        Case 14
         Label3.Caption = Label3.Caption + "0 "
        Case 15
         Label3.Caption = Label3.Caption + "1 "
        Case 16
         Label3.Caption = Label3.Caption + "0 "
        Case 17
         Label3.Caption = Label3.Caption + "1 "
        Case 18
         Label3.Caption = Label3.Caption
        Case 19
         Label3.Caption = Label3.Caption + "1 "
        Case 20
         Label3.Caption = Label3.Caption
    End Select
      
    
        Select Case Int((Rnd * 30) + 1)
        Case 1
         Label4.Caption = Label4.Caption + "1 "
        Case 2
         Label4.Caption = Label4.Caption
        Case 3
         Label4.Caption = Label4.Caption + "1 "
        Case 4
         Label4.Caption = Label4.Caption + "0 "

        Case 6
         Label4.Caption = Label4.Caption + "0 "
        Case 7
         Label4.Caption = Label4.Caption + "1 "
        Case 8
         Label4.Caption = Label4.Caption + "0 "
        Case 9
         Label4.Caption = Label4.Caption
        Case 10
         Label4.Caption = Label4.Caption + "0 "
        Case 11
         Label4.Caption = Label4.Caption + "1 "
        Case 12
         Label4.Caption = Label4.Caption + "0 "
        Case 13
         Label4.Caption = Label4.Caption
        Case 14
         Label4.Caption = Label4.Caption + "0 "
        Case 15
         Label4.Caption = Label4.Caption + "1 "
        Case 16
         Label4.Caption = Label4.Caption + "0 "
        Case 17
         Label4.Caption = Label4.Caption + "1 "
        Case 18
         Label4.Caption = Label4.Caption
        Case 19
         Label4.Caption = Label4.Caption + "1 "
        Case 20
         Label4.Caption = Label4.Caption
    End Select
    
        Select Case Int((Rnd * 30) + 1)
        Case 1
         Label5.Caption = Label5.Caption + "1 "
        Case 2
         Label5.Caption = Label5.Caption
        Case 3
         Label5.Caption = Label5.Caption + "1 "
        Case 4
         Label5.Caption = Label5.Caption + "0 "

        Case 6
         Label5.Caption = Label5.Caption + "0 "
        Case 7
         Label5.Caption = Label5.Caption + "1 "
        Case 8
         Label5.Caption = Label5.Caption + "0 "
        Case 9
         Label5.Caption = Label5.Caption
        Case 10
         Label5.Caption = Label5.Caption + "0 "
        Case 11
         Label5.Caption = Label5.Caption + "1 "
        Case 12
         Label5.Caption = Label5.Caption + "0 "
        Case 13
         Label5.Caption = Label5.Caption
        Case 14
         Label5.Caption = Label5.Caption + "0 "
        Case 15
         Label5.Caption = Label5.Caption + "1 "
        Case 16
         Label5.Caption = Label5.Caption + "0 "
        Case 17
         Label5.Caption = Label5.Caption + "1 "
        Case 18
         Label5.Caption = Label5.Caption
        Case 19
         Label5.Caption = Label5.Caption + "1 "
        Case 20
         Label5.Caption = Label5.Caption
    End Select
    
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label6.Caption = Label6.Caption + "1 "
        Case 2
         Label6.Caption = Label6.Caption
        Case 3
         Label6.Caption = Label6.Caption + "1 "
        Case 4
         Label6.Caption = Label6.Caption + "0 "

        Case 6
         Label6.Caption = Label6.Caption + "0 "
        Case 7
         Label6.Caption = Label6.Caption + "1 "
        Case 8
         Label6.Caption = Label6.Caption + "0 "
        Case 9
         Label6.Caption = Label6.Caption
        Case 10
         Label6.Caption = Label6.Caption + "0 "
        Case 11
         Label6.Caption = Label6.Caption + "1 "
        Case 12
         Label6.Caption = Label6.Caption + "0 "
        Case 13
         Label6.Caption = Label6.Caption
        Case 14
         Label6.Caption = Label6.Caption + "0 "
        Case 15
         Label6.Caption = Label6.Caption + "1 "
        Case 16
         Label6.Caption = Label6.Caption + "0 "
        Case 17
         Label6.Caption = Label6.Caption + "1 "
        Case 18
         Label6.Caption = Label6.Caption
        Case 19
         Label6.Caption = Label6.Caption + "1 "
        Case 20
         Label6.Caption = Label6.Caption
    End Select
    
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label7.Caption = Label7.Caption + "1 "
        Case 2
         Label7.Caption = Label7.Caption
        Case 3
         Label7.Caption = Label7.Caption + "1 "
        Case 4
         Label7.Caption = Label7.Caption + "0 "

        Case 6
         Label7.Caption = Label7.Caption + "0 "
        Case 7
         Label7.Caption = Label7.Caption + "1 "
        Case 8
         Label7.Caption = Label7.Caption + "0 "
        Case 9
         Label7.Caption = Label7.Caption
        Case 10
         Label7.Caption = Label7.Caption + "0 "
        Case 11
         Label7.Caption = Label7.Caption + "1 "
        Case 12
         Label7.Caption = Label7.Caption + "0 "
        Case 13
         Label7.Caption = Label7.Caption
        Case 14
         Label7.Caption = Label7.Caption + "0 "
        Case 15
         Label7.Caption = Label7.Caption + "1 "
        Case 16
         Label7.Caption = Label7.Caption + "0 "
        Case 17
         Label7.Caption = Label7.Caption + "1 "
        Case 18
         Label7.Caption = Label7.Caption
        Case 19
         Label7.Caption = Label7.Caption + "1 "
        Case 20
         Label7.Caption = Label7.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label8.Caption = Label8.Caption + "1 "
        Case 2
         Label8.Caption = Label8.Caption
        Case 3
         Label8.Caption = Label8.Caption + "1 "
        Case 4
         Label8.Caption = Label8.Caption + "0 "

        Case 6
         Label8.Caption = Label8.Caption + "0 "
        Case 7
         Label8.Caption = Label8.Caption + "1 "
        Case 8
         Label8.Caption = Label8.Caption + "0 "
        Case 9
         Label8.Caption = Label8.Caption
        Case 10
         Label8.Caption = Label8.Caption + "0 "
        Case 11
         Label8.Caption = Label8.Caption + "1 "
        Case 12
         Label8.Caption = Label8.Caption + "0 "
        Case 13
         Label8.Caption = Label8.Caption
        Case 14
         Label8.Caption = Label8.Caption + "0 "
        Case 15
         Label8.Caption = Label8.Caption + "1 "
        Case 16
         Label8.Caption = Label8.Caption + "0 "
        Case 17
         Label8.Caption = Label8.Caption + "1 "
        Case 18
         Label8.Caption = Label8.Caption
        Case 19
         Label8.Caption = Label8.Caption + "1 "
        Case 20
         Label8.Caption = Label8.Caption
    End Select
    
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label9.Caption = Label9.Caption + "1 "
        Case 2
         Label9.Caption = Label9.Caption
        Case 3
         Label9.Caption = Label9.Caption + "1 "
        Case 4
         Label9.Caption = Label9.Caption + "0 "

        Case 6
         Label9.Caption = Label9.Caption + "0 "
        Case 7
         Label9.Caption = Label9.Caption + "1 "
        Case 8
         Label9.Caption = Label9.Caption + "0 "
        Case 9
         Label9.Caption = Label9.Caption
        Case 10
         Label9.Caption = Label9.Caption + "0 "
        Case 11
         Label9.Caption = Label9.Caption + "1 "
        Case 12
         Label9.Caption = Label9.Caption + "0 "
        Case 13
         Label9.Caption = Label9.Caption
        Case 14
         Label9.Caption = Label9.Caption + "0 "
        Case 15
         Label9.Caption = Label9.Caption + "1 "
        Case 16
         Label9.Caption = Label9.Caption + "0 "
        Case 17
         Label9.Caption = Label9.Caption + "1 "
        Case 18
         Label9.Caption = Label9.Caption
        Case 19
         Label9.Caption = Label9.Caption + "1 "
        Case 20
         Label9.Caption = Label9.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label10.Caption = Label10.Caption + "1 "
        Case 2
         Label10.Caption = Label10.Caption
        Case 3
         Label10.Caption = Label10.Caption + "1 "
        Case 4
         Label10.Caption = Label10.Caption + "0 "

        Case 6
         Label10.Caption = Label10.Caption + "0 "
        Case 7
         Label10.Caption = Label10.Caption + "1 "
        Case 8
         Label10.Caption = Label10.Caption + "0 "
        Case 9
         Label10.Caption = Label10.Caption
        Case 10
         Label10.Caption = Label10.Caption + "0 "
        Case 11
         Label10.Caption = Label10.Caption + "1 "
        Case 12
         Label10.Caption = Label10.Caption + "0 "
        Case 13
         Label10.Caption = Label10.Caption
        Case 14
         Label10.Caption = Label10.Caption + "0 "
        Case 15
         Label10.Caption = Label10.Caption + "1 "
        Case 16
         Label10.Caption = Label10.Caption + "0 "
        Case 17
         Label10.Caption = Label10.Caption + "1 "
        Case 18
         Label10.Caption = Label10.Caption
        Case 19
         Label10.Caption = Label10.Caption + "1 "
        Case 20
         Label10.Caption = Label10.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label11.Caption = Label11.Caption + "1 "
        Case 2
         Label11.Caption = Label11.Caption
        Case 3
         Label11.Caption = Label11.Caption + "1 "
        Case 4
         Label11.Caption = Label11.Caption + "0 "

        Case 6
         Label11.Caption = Label11.Caption + "0 "
        Case 7
         Label11.Caption = Label11.Caption + "1 "
        Case 8
         Label11.Caption = Label11.Caption + "0 "
        Case 9
         Label11.Caption = Label11.Caption
        Case 10
         Label11.Caption = Label11.Caption + "0 "
        Case 11
         Label11.Caption = Label11.Caption + "1 "
        Case 12
         Label11.Caption = Label11.Caption + "0 "
        Case 13
         Label11.Caption = Label11.Caption
        Case 14
         Label11.Caption = Label11.Caption + "0 "
        Case 15
         Label11.Caption = Label11.Caption + "1 "
        Case 16
         Label11.Caption = Label11.Caption + "0 "
        Case 17
         Label11.Caption = Label11.Caption + "1 "
        Case 18
         Label11.Caption = Label11.Caption
        Case 19
         Label11.Caption = Label11.Caption + "1 "
        Case 20
         Label11.Caption = Label11.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label12.Caption = Label12.Caption + "1 "
        Case 2
         Label12.Caption = Label12.Caption
        Case 3
         Label12.Caption = Label12.Caption + "1 "
        Case 4
         Label12.Caption = Label12.Caption + "0 "

        Case 6
         Label12.Caption = Label12.Caption + "0 "
        Case 7
         Label12.Caption = Label12.Caption + "1 "
        Case 8
         Label12.Caption = Label12.Caption + "0 "
        Case 9
         Label12.Caption = Label12.Caption
        Case 10
         Label12.Caption = Label12.Caption + "0 "
        Case 11
         Label12.Caption = Label12.Caption + "1 "
        Case 12
         Label12.Caption = Label12.Caption + "0 "
        Case 13
         Label12.Caption = Label12.Caption
        Case 14
         Label12.Caption = Label12.Caption + "0 "
        Case 15
         Label12.Caption = Label12.Caption + "1 "
        Case 16
         Label12.Caption = Label12.Caption + "0 "
        Case 17
         Label12.Caption = Label12.Caption + "1 "
        Case 18
         Label12.Caption = Label12.Caption
        Case 19
         Label12.Caption = Label12.Caption + "1 "
        Case 20
         Label12.Caption = Label12.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label13.Caption = Label13.Caption + "1 "
        Case 2
         Label13.Caption = Label13.Caption
        Case 3
         Label13.Caption = Label13.Caption + "1 "
        Case 4
         Label13.Caption = Label13.Caption + "0 "

        Case 6
         Label13.Caption = Label13.Caption + "0 "
        Case 7
         Label13.Caption = Label13.Caption + "1 "
        Case 8
         Label13.Caption = Label13.Caption + "0 "
        Case 9
         Label13.Caption = Label13.Caption
        Case 10
         Label13.Caption = Label13.Caption + "0 "
        Case 11
         Label13.Caption = Label13.Caption + "1 "
        Case 12
         Label13.Caption = Label13.Caption + "0 "
        Case 13
         Label13.Caption = Label13.Caption
        Case 14
         Label13.Caption = Label13.Caption + "0 "
        Case 15
         Label13.Caption = Label13.Caption + "1 "
        Case 16
         Label13.Caption = Label13.Caption + "0 "
        Case 17
         Label13.Caption = Label13.Caption + "1 "
        Case 18
         Label13.Caption = Label13.Caption
        Case 19
         Label13.Caption = Label13.Caption + "1 "
        Case 20
         Label13.Caption = Label13.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label14.Caption = Label14.Caption + "1 "
        Case 2
         Label14.Caption = Label14.Caption
        Case 3
         Label14.Caption = Label14.Caption + "1 "
        Case 4
         Label14.Caption = Label14.Caption + "0 "

        Case 6
         Label14.Caption = Label14.Caption + "0 "
        Case 7
         Label14.Caption = Label14.Caption + "1 "
        Case 8
         Label14.Caption = Label14.Caption + "0 "
        Case 9
         Label14.Caption = Label14.Caption
        Case 10
         Label14.Caption = Label14.Caption + "0 "
        Case 11
         Label14.Caption = Label14.Caption + "1 "
        Case 12
         Label14.Caption = Label14.Caption + "0 "
        Case 13
         Label14.Caption = Label14.Caption
        Case 14
         Label14.Caption = Label14.Caption + "0 "
        Case 15
         Label14.Caption = Label14.Caption + "1 "
        Case 16
         Label14.Caption = Label14.Caption + "0 "
        Case 17
         Label14.Caption = Label14.Caption + "1 "
        Case 18
         Label14.Caption = Label14.Caption
        Case 19
         Label14.Caption = Label14.Caption + "1 "
        Case 20
         Label14.Caption = Label14.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label15.Caption = Label15.Caption + "1 "
        Case 2
         Label15.Caption = Label15.Caption
        Case 3
         Label15.Caption = Label15.Caption + "1 "
        Case 4
         Label15.Caption = Label15.Caption + "0 "

        Case 6
         Label15.Caption = Label15.Caption + "0 "
        Case 7
         Label15.Caption = Label15.Caption + "1 "
        Case 8
         Label15.Caption = Label15.Caption + "0 "
        Case 9
         Label15.Caption = Label15.Caption
        Case 10
         Label15.Caption = Label15.Caption + "0 "
        Case 11
         Label15.Caption = Label15.Caption + "1 "
        Case 12
         Label15.Caption = Label15.Caption + "0 "
        Case 13
         Label15.Caption = Label15.Caption
        Case 14
         Label15.Caption = Label15.Caption + "0 "
        Case 15
         Label15.Caption = Label15.Caption + "1 "
        Case 16
         Label15.Caption = Label15.Caption + "0 "
        Case 17
         Label15.Caption = Label15.Caption + "1 "
        Case 18
         Label15.Caption = Label15.Caption
        Case 19
         Label15.Caption = Label15.Caption + "1 "
        Case 20
         Label15.Caption = Label15.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label16.Caption = Label16.Caption + "1 "
        Case 2
         Label16.Caption = Label16.Caption
        Case 3
         Label16.Caption = Label16.Caption + "1 "
        Case 4
         Label16.Caption = Label16.Caption + "0 "

        Case 6
         Label16.Caption = Label16.Caption + "0 "
        Case 7
         Label16.Caption = Label16.Caption + "1 "
        Case 8
         Label16.Caption = Label16.Caption + "0 "
        Case 9
         Label16.Caption = Label16.Caption
        Case 10
         Label16.Caption = Label16.Caption + "0 "
        Case 11
         Label16.Caption = Label16.Caption + "1 "
        Case 12
         Label16.Caption = Label16.Caption + "0 "
        Case 13
         Label16.Caption = Label16.Caption
        Case 14
         Label16.Caption = Label16.Caption + "0 "
        Case 15
         Label16.Caption = Label16.Caption + "1 "
        Case 16
         Label16.Caption = Label16.Caption + "0 "
        Case 17
         Label16.Caption = Label16.Caption + "1 "
        Case 18
         Label16.Caption = Label16.Caption
        Case 19
         Label16.Caption = Label16.Caption + "1 "
        Case 20
         Label16.Caption = Label16.Caption
    End Select

    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label17.Caption = Label17.Caption + "1 "
        Case 2
         Label17.Caption = Label17.Caption
        Case 3
         Label17.Caption = Label17.Caption + "1 "
        Case 4
         Label17.Caption = Label17.Caption + "0 "
   
        Case 6
         Label17.Caption = Label17.Caption + "0 "
        Case 7
         Label17.Caption = Label17.Caption + "1 "
        Case 8
         Label17.Caption = Label17.Caption + "0 "
        Case 9
         Label17.Caption = Label17.Caption
        Case 10
         Label17.Caption = Label17.Caption + "0 "
        Case 11
         Label17.Caption = Label17.Caption + "1 "
        Case 12
         Label17.Caption = Label17.Caption + "0 "
        Case 13
         Label17.Caption = Label17.Caption
        Case 14
         Label17.Caption = Label17.Caption + "0 "
        Case 15
         Label17.Caption = Label17.Caption + "1 "
        Case 16
         Label17.Caption = Label17.Caption + "0 "
        Case 17
         Label17.Caption = Label17.Caption + "1 "
        Case 18
         Label17.Caption = Label17.Caption
        Case 19
         Label17.Caption = Label17.Caption + "1 "
        Case 20
         Label17.Caption = Label17.Caption
    End Select
    
            Select Case Int((Rnd * 30) + 1)
        Case 1
         Label18.Caption = Label18.Caption + "1 "
        Case 2
         Label18.Caption = Label18.Caption
        Case 3
         Label18.Caption = Label18.Caption + "1 "
        Case 4
         Label18.Caption = Label18.Caption + "0 "

        Case 6
         Label18.Caption = Label18.Caption + "0 "
        Case 7
         Label18.Caption = Label18.Caption + "1 "
        Case 8
         Label18.Caption = Label18.Caption + "0 "
        Case 9
         Label18.Caption = Label18.Caption
        Case 10
         Label18.Caption = Label18.Caption + "0 "
        Case 11
         Label18.Caption = Label18.Caption + "1 "
        Case 12
         Label18.Caption = Label18.Caption + "0 "
        Case 13
         Label18.Caption = Label18.Caption
        Case 14
         Label18.Caption = Label18.Caption + "0 "
        Case 15
         Label18.Caption = Label18.Caption + "1 "
        Case 16
         Label18.Caption = Label18.Caption + "0 "
        Case 17
         Label18.Caption = Label18.Caption + "1 "
        Case 18
         Label18.Caption = Label18.Caption
        Case 19
         Label18.Caption = Label18.Caption + "1 "
        Case 20
         Label18.Caption = Label18.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label19.Caption = Label19.Caption + "1 "
        Case 2
         Label19.Caption = Label19.Caption
        Case 3
         Label19.Caption = Label19.Caption + "1 "
        Case 4
         Label19.Caption = Label19.Caption + "0 "

        Case 6
         Label19.Caption = Label19.Caption + "0 "
        Case 7
         Label19.Caption = Label19.Caption + "1 "
        Case 8
         Label19.Caption = Label19.Caption + "0 "
        Case 9
         Label19.Caption = Label19.Caption
        Case 10
         Label19.Caption = Label19.Caption + "0 "
        Case 11
         Label19.Caption = Label19.Caption + "1 "
        Case 12
         Label19.Caption = Label19.Caption + "0 "
        Case 13
         Label19.Caption = Label19.Caption
        Case 14
         Label19.Caption = Label19.Caption + "0 "
        Case 15
         Label19.Caption = Label19.Caption + "1 "
        Case 16
         Label19.Caption = Label19.Caption + "0 "
        Case 17
         Label19.Caption = Label19.Caption + "1 "
        Case 18
         Label19.Caption = Label19.Caption
        Case 19
         Label19.Caption = Label19.Caption + "1 "
        Case 20
         Label19.Caption = Label19.Caption
    End Select
    
    
                    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label20.Caption = Label20.Caption + "1 "
        Case 2
         Label20.Caption = Label20.Caption
        Case 3
         Label20.Caption = Label20.Caption + "1 "
        Case 4
         Label20.Caption = Label20.Caption + "0 "

        Case 6
         Label20.Caption = Label20.Caption + "0 "
        Case 7
         Label20.Caption = Label20.Caption + "1 "
        Case 8
         Label20.Caption = Label20.Caption + "0 "
        Case 9
         Label20.Caption = Label20.Caption
        Case 10
         Label20.Caption = Label20.Caption + "0 "
        Case 11
         Label20.Caption = Label20.Caption + "1 "
        Case 12
         Label20.Caption = Label20.Caption + "0 "
        Case 13
         Label20.Caption = Label20.Caption
        Case 14
         Label20.Caption = Label20.Caption + "0 "
        Case 15
         Label20.Caption = Label20.Caption + "1 "
        Case 16
         Label20.Caption = Label20.Caption + "0 "
        Case 17
         Label20.Caption = Label20.Caption + "1 "
        Case 18
         Label20.Caption = Label20.Caption
        Case 19
         Label20.Caption = Label20.Caption + "1 "
        Case 20
         Label20.Caption = Label20.Caption
    End Select
    
                    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label21.Caption = Label21.Caption + "1 "
        Case 2
         Label21.Caption = Label21.Caption
        Case 3
         Label21.Caption = Label21.Caption + "1 "
        Case 4
         Label21.Caption = Label21.Caption + "0 "

        Case 6
         Label21.Caption = Label21.Caption + "0 "
        Case 7
         Label21.Caption = Label21.Caption + "1 "
        Case 8
         Label21.Caption = Label21.Caption + "0 "
        Case 9
         Label21.Caption = Label21.Caption
        Case 10
         Label21.Caption = Label21.Caption + "0 "
        Case 11
         Label21.Caption = Label21.Caption + "1 "
        Case 12
         Label21.Caption = Label21.Caption + "0 "
        Case 13
         Label21.Caption = Label21.Caption
        Case 14
         Label21.Caption = Label21.Caption + "0 "
        Case 15
         Label21.Caption = Label21.Caption + "1 "
        Case 16
         Label21.Caption = Label21.Caption + "0 "
        Case 17
         Label21.Caption = Label21.Caption + "1 "
        Case 18
         Label21.Caption = Label21.Caption
        Case 19
         Label21.Caption = Label21.Caption + "1 "
        Case 20
         Label21.Caption = Label21.Caption
    End Select
    
                    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label22.Caption = Label22.Caption + "1 "
        Case 2
         Label22.Caption = Label22.Caption
        Case 3
         Label22.Caption = Label22.Caption + "1 "
        Case 4
         Label22.Caption = Label22.Caption + "0 "

        Case 6
         Label22.Caption = Label22.Caption + "0 "
        Case 7
         Label22.Caption = Label22.Caption + "1 "
        Case 8
         Label22.Caption = Label22.Caption + "0 "
        Case 9
         Label22.Caption = Label22.Caption
        Case 10
         Label22.Caption = Label22.Caption + "0 "
        Case 11
         Label22.Caption = Label22.Caption + "1 "
        Case 12
         Label22.Caption = Label22.Caption + "0 "
        Case 13
         Label22.Caption = Label22.Caption
        Case 14
         Label22.Caption = Label22.Caption + "0 "
        Case 15
         Label22.Caption = Label22.Caption + "1 "
        Case 16
         Label22.Caption = Label22.Caption + "0 "
        Case 17
         Label22.Caption = Label22.Caption + "1 "
        Case 18
         Label22.Caption = Label22.Caption
        Case 19
         Label22.Caption = Label22.Caption + "1 "
        Case 20
         Label22.Caption = Label22.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label23.Caption = Label23.Caption + "1 "
        Case 2
         Label23.Caption = Label23.Caption
        Case 3
         Label23.Caption = Label23.Caption + "1 "
        Case 4
         Label23.Caption = Label23.Caption + "0 "

        Case 6
         Label23.Caption = Label23.Caption + "0 "
        Case 7
         Label23.Caption = Label23.Caption + "1 "
        Case 8
         Label23.Caption = Label23.Caption + "0 "
        Case 9
         Label23.Caption = Label23.Caption
        Case 10
         Label23.Caption = Label23.Caption + "0 "
        Case 11
         Label23.Caption = Label23.Caption + "1 "
        Case 12
         Label23.Caption = Label23.Caption + "0 "
        Case 13
         Label23.Caption = Label23.Caption
        Case 14
         Label23.Caption = Label23.Caption + "0 "
        Case 15
         Label23.Caption = Label23.Caption + "1 "
        Case 16
         Label23.Caption = Label23.Caption + "0 "
        Case 17
         Label23.Caption = Label23.Caption + "1 "
        Case 18
         Label23.Caption = Label23.Caption
        Case 19
         Label23.Caption = Label23.Caption + "1 "
        Case 20
         Label23.Caption = Label23.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label24.Caption = Label24.Caption + "1 "
        Case 2
         Label24.Caption = Label24.Caption
        Case 3
         Label24.Caption = Label24.Caption + "1 "
        Case 4
         Label24.Caption = Label24.Caption + "0 "

        Case 6
         Label24.Caption = Label24.Caption + "0 "
        Case 7
         Label24.Caption = Label24.Caption + "1 "
        Case 8
         Label24.Caption = Label24.Caption + "0 "
        Case 9
         Label24.Caption = Label24.Caption
        Case 10
         Label24.Caption = Label24.Caption + "0 "
        Case 11
         Label24.Caption = Label24.Caption + "1 "
        Case 12
         Label24.Caption = Label24.Caption + "0 "
        Case 13
         Label24.Caption = Label24.Caption
        Case 14
         Label24.Caption = Label24.Caption + "0 "
        Case 15
         Label24.Caption = Label24.Caption + "1 "
        Case 16
         Label24.Caption = Label24.Caption + "0 "
        Case 17
         Label24.Caption = Label24.Caption + "1 "
        Case 18
         Label24.Caption = Label24.Caption
        Case 19
         Label24.Caption = Label24.Caption + "1 "
        Case 20
         Label24.Caption = Label24.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label25.Caption = Label25.Caption + "1 "
        Case 2
         Label25.Caption = Label25.Caption
        Case 3
         Label25.Caption = Label25.Caption + "1 "
        Case 4
         Label25.Caption = Label25.Caption + "0 "

        Case 6
         Label25.Caption = Label25.Caption + "0 "
        Case 7
         Label25.Caption = Label25.Caption + "1 "
        Case 8
         Label25.Caption = Label25.Caption + "0 "
        Case 9
         Label25.Caption = Label25.Caption
        Case 10
         Label25.Caption = Label25.Caption + "0 "
        Case 11
         Label25.Caption = Label25.Caption + "1 "
        Case 12
         Label25.Caption = Label25.Caption + "0 "
        Case 13
         Label25.Caption = Label25.Caption
        Case 14
         Label25.Caption = Label25.Caption + "0 "
        Case 15
         Label25.Caption = Label25.Caption + "1 "
        Case 16
         Label25.Caption = Label25.Caption + "0 "
        Case 17
         Label25.Caption = Label25.Caption + "1 "
        Case 18
         Label25.Caption = Label25.Caption
        Case 19
         Label25.Caption = Label25.Caption + "1 "
        Case 20
         Label25.Caption = Label25.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label26.Caption = Label26.Caption + "1 "
        Case 2
         Label26.Caption = Label26.Caption
        Case 3
         Label26.Caption = Label26.Caption + "1 "
        Case 4
         Label26.Caption = Label26.Caption + "0 "

        Case 6
         Label26.Caption = Label26.Caption + "0 "
        Case 7
         Label26.Caption = Label26.Caption + "1 "
        Case 8
         Label26.Caption = Label26.Caption + "0 "
        Case 9
         Label26.Caption = Label26.Caption
        Case 10
         Label26.Caption = Label26.Caption + "0 "
        Case 11
         Label26.Caption = Label26.Caption + "1 "
        Case 12
         Label26.Caption = Label26.Caption + "0 "
        Case 13
         Label26.Caption = Label26.Caption
        Case 14
         Label26.Caption = Label26.Caption + "0 "
        Case 15
         Label26.Caption = Label26.Caption + "1 "
        Case 16
         Label26.Caption = Label26.Caption + "0 "
        Case 17
         Label26.Caption = Label26.Caption + "1 "
        Case 18
         Label26.Caption = Label26.Caption
        Case 19
         Label26.Caption = Label26.Caption + "1 "
        Case 20
         Label26.Caption = Label26.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label27.Caption = Label27.Caption + "1 "
        Case 2
         Label27.Caption = Label27.Caption
        Case 3
         Label27.Caption = Label27.Caption + "1 "
        Case 4
         Label27.Caption = Label27.Caption + "0 "

        Case 6
         Label27.Caption = Label27.Caption + "0 "
        Case 7
         Label27.Caption = Label27.Caption + "1 "
        Case 8
         Label27.Caption = Label27.Caption + "0 "
        Case 9
         Label27.Caption = Label27.Caption
        Case 10
         Label27.Caption = Label27.Caption + "0 "
        Case 11
         Label27.Caption = Label27.Caption + "1 "
        Case 12
         Label27.Caption = Label27.Caption + "0 "
        Case 13
         Label27.Caption = Label27.Caption
        Case 14
         Label27.Caption = Label27.Caption + "0 "
        Case 15
         Label27.Caption = Label27.Caption + "1 "
        Case 16
         Label27.Caption = Label27.Caption + "0 "
        Case 17
         Label27.Caption = Label27.Caption + "1 "
        Case 18
         Label27.Caption = Label27.Caption
        Case 19
         Label27.Caption = Label27.Caption + "1 "
        Case 20
         Label27.Caption = Label27.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label28.Caption = Label28.Caption + "1 "
        Case 2
         Label28.Caption = Label28.Caption
        Case 3
         Label28.Caption = Label28.Caption + "1 "
        Case 4
         Label28.Caption = Label28.Caption + "0 "

        Case 6
         Label28.Caption = Label28.Caption + "0 "
        Case 7
         Label28.Caption = Label28.Caption + "1 "
        Case 8
         Label28.Caption = Label28.Caption + "0 "
        Case 9
         Label28.Caption = Label28.Caption
        Case 10
         Label28.Caption = Label28.Caption + "0 "
        Case 11
         Label28.Caption = Label28.Caption + "1 "
        Case 12
         Label28.Caption = Label28.Caption + "0 "
        Case 13
         Label28.Caption = Label28.Caption
        Case 14
         Label28.Caption = Label28.Caption + "0 "
        Case 15
         Label28.Caption = Label28.Caption + "1 "
        Case 16
         Label28.Caption = Label28.Caption + "0 "
        Case 17
         Label28.Caption = Label28.Caption + "1 "
        Case 18
         Label28.Caption = Label28.Caption
        Case 19
         Label28.Caption = Label28.Caption + "1 "
        Case 20
         Label28.Caption = Label28.Caption
    End Select
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label29.Caption = Label29.Caption + "1 "
        Case 2
         Label29.Caption = Label29.Caption
        Case 3
         Label29.Caption = Label29.Caption + "1 "
        Case 4
         Label29.Caption = Label29.Caption + "0 "

        Case 6
         Label29.Caption = Label29.Caption + "0 "
        Case 7
         Label29.Caption = Label29.Caption + "1 "
        Case 8
         Label29.Caption = Label29.Caption + "0 "
        Case 9
         Label29.Caption = Label29.Caption
        Case 10
         Label29.Caption = Label29.Caption + "0 "
        Case 11
         Label29.Caption = Label29.Caption + "1 "
        Case 12
         Label29.Caption = Label29.Caption + "0 "
        Case 13
         Label29.Caption = Label29.Caption
        Case 14
         Label29.Caption = Label29.Caption + "0 "
        Case 15
         Label29.Caption = Label29.Caption + "1 "
        Case 16
         Label29.Caption = Label29.Caption + "0 "
        Case 17
         Label29.Caption = Label29.Caption + "1 "
        Case 18
         Label29.Caption = Label29.Caption
        Case 19
         Label29.Caption = Label29.Caption + "1 "
        Case 20
         Label29.Caption = Label29.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label30.Caption = Label30.Caption + "1 "
        Case 2
         Label30.Caption = Label30.Caption
        Case 3
         Label30.Caption = Label30.Caption + "1 "
        Case 4
         Label30.Caption = Label30.Caption + "0 "

        Case 6
         Label30.Caption = Label30.Caption + "0 "
        Case 7
         Label30.Caption = Label30.Caption + "1 "
        Case 8
         Label30.Caption = Label30.Caption + "0 "
        Case 9
         Label30.Caption = Label30.Caption
        Case 10
         Label30.Caption = Label30.Caption + "0 "
        Case 11
         Label30.Caption = Label30.Caption + "1 "
        Case 12
         Label30.Caption = Label30.Caption + "0 "
        Case 13
         Label30.Caption = Label30.Caption
        Case 14
         Label30.Caption = Label30.Caption + "0 "
        Case 15
         Label30.Caption = Label30.Caption + "1 "
        Case 16
         Label30.Caption = Label30.Caption + "0 "
        Case 17
         Label30.Caption = Label30.Caption + "1 "
        Case 18
         Label30.Caption = Label30.Caption
        Case 19
         Label30.Caption = Label30.Caption + "1 "
        Case 20
         Label30.Caption = Label30.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label31.Caption = Label31.Caption + "1 "
        Case 2
         Label31.Caption = Label31.Caption
        Case 3
         Label31.Caption = Label31.Caption + "1 "
        Case 4
         Label31.Caption = Label31.Caption + "0 "

        Case 6
         Label31.Caption = Label31.Caption + "0 "
        Case 7
         Label31.Caption = Label31.Caption + "1 "
        Case 8
         Label31.Caption = Label31.Caption + "0 "
        Case 9
         Label31.Caption = Label31.Caption
        Case 10
         Label31.Caption = Label31.Caption + "0 "
        Case 11
         Label31.Caption = Label31.Caption + "1 "
        Case 12
         Label31.Caption = Label31.Caption + "0 "
        Case 13
         Label31.Caption = Label31.Caption
        Case 14
         Label31.Caption = Label31.Caption + "0 "
        Case 15
         Label31.Caption = Label31.Caption + "1 "
        Case 16
         Label31.Caption = Label31.Caption + "0 "
        Case 17
         Label31.Caption = Label31.Caption + "1 "
        Case 18
         Label31.Caption = Label31.Caption
        Case 19
         Label31.Caption = Label31.Caption + "1 "
        Case 20
         Label31.Caption = Label31.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label32.Caption = Label32.Caption + "1 "
        Case 2
         Label32.Caption = Label32.Caption
        Case 3
         Label32.Caption = Label32.Caption + "1 "
        Case 4
         Label32.Caption = Label32.Caption + "0 "

        Case 6
         Label32.Caption = Label32.Caption + "0 "
        Case 7
         Label32.Caption = Label32.Caption + "1 "
        Case 8
         Label32.Caption = Label32.Caption + "0 "
        Case 9
         Label32.Caption = Label32.Caption
        Case 10
         Label32.Caption = Label32.Caption + "0 "
        Case 11
         Label32.Caption = Label32.Caption + "1 "
        Case 12
         Label32.Caption = Label32.Caption + "0 "
        Case 13
         Label32.Caption = Label32.Caption
        Case 14
         Label32.Caption = Label32.Caption + "0 "
        Case 15
         Label32.Caption = Label32.Caption + "1 "
        Case 16
         Label32.Caption = Label32.Caption + "0 "
        Case 17
         Label32.Caption = Label32.Caption + "1 "
        Case 18
         Label32.Caption = Label32.Caption
        Case 19
         Label32.Caption = Label32.Caption + "1 "
        Case 20
         Label32.Caption = Label32.Caption
    End Select
    
                    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label33.Caption = Label33.Caption + "1 "
        Case 2
         Label33.Caption = Label33.Caption
        Case 3
         Label33.Caption = Label33.Caption + "1 "
        Case 4
         Label33.Caption = Label33.Caption + "0 "

        Case 6
         Label33.Caption = Label33.Caption + "0 "
        Case 7
         Label33.Caption = Label33.Caption + "1 "
        Case 8
         Label33.Caption = Label33.Caption + "0 "
        Case 9
         Label33.Caption = Label33.Caption
        Case 10
         Label33.Caption = Label33.Caption + "0 "
        Case 11
         Label33.Caption = Label33.Caption + "1 "
        Case 12
         Label33.Caption = Label33.Caption + "0 "
        Case 13
         Label33.Caption = Label33.Caption
        Case 14
         Label33.Caption = Label33.Caption + "0 "
        Case 15
         Label33.Caption = Label33.Caption + "1 "
        Case 16
         Label33.Caption = Label33.Caption + "0 "
        Case 17
         Label33.Caption = Label33.Caption + "1 "
        Case 18
         Label33.Caption = Label33.Caption
        Case 19
         Label33.Caption = Label33.Caption + "1 "
        Case 20
         Label33.Caption = Label33.Caption
    End Select
    
                    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label34.Caption = Label34.Caption + "1 "
        Case 2
         Label34.Caption = Label34.Caption
        Case 3
         Label34.Caption = Label34.Caption + "1 "
        Case 4
         Label34.Caption = Label34.Caption + "0 "

        Case 6
         Label34.Caption = Label34.Caption + "0 "
        Case 7
         Label34.Caption = Label34.Caption + "1 "
        Case 8
         Label34.Caption = Label34.Caption + "0 "
        Case 9
         Label34.Caption = Label34.Caption
        Case 10
         Label34.Caption = Label34.Caption + "0 "
        Case 11
         Label34.Caption = Label34.Caption + "1 "
        Case 12
         Label34.Caption = Label34.Caption + "0 "
        Case 13
         Label34.Caption = Label34.Caption
        Case 14
         Label34.Caption = Label34.Caption + "0 "
        Case 15
         Label34.Caption = Label34.Caption + "1 "
        Case 16
         Label34.Caption = Label34.Caption + "0 "
        Case 17
         Label34.Caption = Label34.Caption + "1 "
        Case 18
         Label34.Caption = Label34.Caption
        Case 19
         Label34.Caption = Label34.Caption + "1 "
        Case 20
         Label34.Caption = Label34.Caption
    End Select
    
                    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label35.Caption = Label35.Caption + "1 "
        Case 2
         Label35.Caption = Label35.Caption
        Case 3
         Label35.Caption = Label35.Caption + "1 "
        Case 4
         Label35.Caption = Label35.Caption + "0 "

        Case 6
         Label35.Caption = Label35.Caption + "0 "
        Case 7
         Label35.Caption = Label35.Caption + "1 "
        Case 8
         Label35.Caption = Label35.Caption + "0 "
        Case 9
         Label35.Caption = Label35.Caption
        Case 10
         Label35.Caption = Label35.Caption + "0 "
        Case 11
         Label35.Caption = Label35.Caption + "1 "
        Case 12
         Label35.Caption = Label35.Caption + "0 "
        Case 13
         Label35.Caption = Label35.Caption
        Case 14
         Label35.Caption = Label35.Caption + "0 "
        Case 15
         Label35.Caption = Label35.Caption + "1 "
        Case 16
         Label35.Caption = Label35.Caption + "0 "
        Case 17
         Label35.Caption = Label35.Caption + "1 "
        Case 18
         Label35.Caption = Label35.Caption
        Case 19
         Label35.Caption = Label35.Caption + "1 "
        Case 20
         Label35.Caption = Label35.Caption
    End Select
    
                    Select Case Int((Rnd * 30) + 1)
        Case 1
         Label36.Caption = Label36.Caption + "1 "
        Case 2
         Label36.Caption = Label36.Caption
        Case 3
         Label36.Caption = Label36.Caption + "1 "
        Case 4
         Label36.Caption = Label36.Caption + "0 "

        Case 6
         Label36.Caption = Label36.Caption + "0 "
        Case 7
         Label36.Caption = Label36.Caption + "1 "
        Case 8
         Label36.Caption = Label36.Caption + "0 "
        Case 9
         Label36.Caption = Label36.Caption
        Case 10
         Label36.Caption = Label36.Caption + "0 "
        Case 11
         Label36.Caption = Label36.Caption + "1 "
        Case 12
         Label36.Caption = Label36.Caption + "0 "
        Case 13
         Label36.Caption = Label36.Caption
        Case 14
         Label36.Caption = Label36.Caption + "0 "
        Case 15
         Label36.Caption = Label36.Caption + "1 "
        Case 16
         Label36.Caption = Label36.Caption + "0 "
        Case 17
         Label36.Caption = Label36.Caption + "1 "
        Case 18
         Label36.Caption = Label36.Caption
        Case 19
         Label36.Caption = Label36.Caption + "1 "
        Case 20
         Label36.Caption = Label36.Caption
    End Select
    
           
    
    
    
    
    
End Sub

Private Sub Timer2_Timer()
i = Int((Rnd * 31) + 1)
Controls("Label" & i).ForeColor = &HFF00&
a = Int((Rnd * 31) + 1)
Controls("Label" & a).ForeColor = &HC000&
b = Int((Rnd * 31) + 1)
Controls("Label" & b).ForeColor = &H8000&
End Sub

Private Sub Timer3_Timer()
C = Int((Rnd * 31) + 1)
Controls("Label" & C).ForeColor = &H4000&
End Sub


Private Sub Timer6_Timer()
         Select Case Int((Rnd * 30) + 1)
        Case 1
         Label37.Caption = Label37.Caption + "1 "
        Case 2
         Label37.Caption = Label37.Caption
        Case 3
         Label37.Caption = Label37.Caption + "1 "
        Case 4
         Label37.Caption = Label37.Caption + "0 "

        Case 6
         Label37.Caption = Label37.Caption + "0 "
        Case 7
         Label37.Caption = Label37.Caption + "1 "
        Case 8
         Label37.Caption = Label37.Caption + "0 "
        Case 9
         Label37.Caption = Label37.Caption
        Case 10
         Label37.Caption = Label37.Caption + "0 "
        Case 11
         Label37.Caption = Label37.Caption + "1 "
        Case 12
         Label37.Caption = Label37.Caption + "0 "
        Case 13
         Label37.Caption = Label37.Caption
        Case 14
         Label37.Caption = Label37.Caption + "0 "
        Case 15
         Label37.Caption = Label37.Caption + "1 "
        Case 16
         Label37.Caption = Label37.Caption + "0 "
        Case 17
         Label37.Caption = Label37.Caption + "1 "
        Case 18
         Label37.Caption = Label37.Caption
        Case 19
         Label37.Caption = Label37.Caption + "1 "
        Case 20
         Label37.Caption = Label37.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label38.Caption = Label38.Caption + "1 "
        Case 2
         Label38.Caption = Label38.Caption
        Case 3
         Label38.Caption = Label38.Caption + "1 "
        Case 4
         Label38.Caption = Label38.Caption + "0 "

        Case 6
         Label38.Caption = Label38.Caption + "0 "
        Case 7
         Label38.Caption = Label38.Caption + "1 "
        Case 8
         Label38.Caption = Label38.Caption + "0 "
        Case 9
         Label38.Caption = Label38.Caption
        Case 10
         Label38.Caption = Label38.Caption + "0 "
        Case 11
         Label38.Caption = Label38.Caption + "1 "
        Case 12
         Label38.Caption = Label38.Caption + "0 "
        Case 13
         Label38.Caption = Label38.Caption
        Case 14
         Label38.Caption = Label38.Caption + "0 "
        Case 15
         Label38.Caption = Label38.Caption + "1 "
        Case 16
         Label38.Caption = Label38.Caption + "0 "
        Case 17
         Label38.Caption = Label38.Caption + "1 "
        Case 18
         Label38.Caption = Label38.Caption
        Case 19
         Label38.Caption = Label38.Caption + "1 "
        Case 20
         Label38.Caption = Label38.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label39.Caption = Label39.Caption + "1 "
        Case 2
         Label39.Caption = Label39.Caption
        Case 3
         Label39.Caption = Label39.Caption + "1 "
        Case 4
         Label39.Caption = Label39.Caption + "0 "

        Case 6
         Label39.Caption = Label39.Caption + "0 "
        Case 7
         Label39.Caption = Label39.Caption + "1 "
        Case 8
         Label39.Caption = Label39.Caption + "0 "
        Case 9
         Label39.Caption = Label39.Caption
        Case 10
         Label39.Caption = Label39.Caption + "0 "
        Case 11
         Label39.Caption = Label39.Caption + "1 "
        Case 12
         Label39.Caption = Label39.Caption + "0 "
        Case 13
         Label39.Caption = Label39.Caption
        Case 14
         Label39.Caption = Label39.Caption + "0 "
        Case 15
         Label39.Caption = Label39.Caption + "1 "
        Case 16
         Label39.Caption = Label39.Caption + "0 "
        Case 17
         Label39.Caption = Label39.Caption + "1 "
        Case 18
         Label39.Caption = Label39.Caption
        Case 19
         Label39.Caption = Label39.Caption + "1 "
        Case 20
         Label39.Caption = Label39.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label40.Caption = Label40.Caption + "1 "
        Case 2
         Label40.Caption = Label40.Caption
        Case 3
         Label40.Caption = Label40.Caption + "1 "
        Case 4
         Label40.Caption = Label40.Caption + "0 "

        Case 6
         Label40.Caption = Label40.Caption + "0 "
        Case 7
         Label40.Caption = Label40.Caption + "1 "
        Case 8
         Label40.Caption = Label40.Caption + "0 "
        Case 9
         Label40.Caption = Label40.Caption
        Case 10
         Label40.Caption = Label40.Caption + "0 "
        Case 11
         Label40.Caption = Label40.Caption + "1 "
        Case 12
         Label40.Caption = Label40.Caption + "0 "
        Case 13
         Label40.Caption = Label40.Caption
        Case 14
         Label40.Caption = Label40.Caption + "0 "
        Case 15
         Label40.Caption = Label40.Caption + "1 "
        Case 16
         Label40.Caption = Label40.Caption + "0 "
        Case 17
         Label40.Caption = Label40.Caption + "1 "
        Case 18
         Label40.Caption = Label40.Caption
        Case 19
         Label40.Caption = Label40.Caption + "1 "
        Case 20
         Label40.Caption = Label40.Caption
    End Select
    
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label41.Caption = Label41.Caption + "1 "
        Case 2
         Label41.Caption = Label41.Caption
        Case 3
         Label41.Caption = Label41.Caption + "1 "
        Case 4
         Label41.Caption = Label41.Caption + "0 "

        Case 6
         Label41.Caption = Label41.Caption + "0 "
        Case 7
         Label41.Caption = Label41.Caption + "1 "
        Case 8
         Label41.Caption = Label41.Caption + "0 "
        Case 9
         Label41.Caption = Label41.Caption
        Case 10
         Label41.Caption = Label41.Caption + "0 "
        Case 11
         Label41.Caption = Label41.Caption + "1 "
        Case 12
         Label41.Caption = Label41.Caption + "0 "
        Case 13
         Label41.Caption = Label41.Caption
        Case 14
         Label41.Caption = Label41.Caption + "0 "
        Case 15
         Label41.Caption = Label41.Caption + "1 "
        Case 16
         Label41.Caption = Label41.Caption + "0 "
        Case 17
         Label41.Caption = Label41.Caption + "1 "
        Case 18
         Label41.Caption = Label41.Caption
        Case 19
         Label41.Caption = Label41.Caption + "1 "
        Case 20
         Label41.Caption = Label41.Caption
    End Select
    
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label42.Caption = Label42.Caption + "1 "
        Case 2
         Label42.Caption = Label42.Caption
        Case 3
         Label42.Caption = Label42.Caption + "1 "
        Case 4
         Label42.Caption = Label42.Caption + "0 "

        Case 6
         Label42.Caption = Label42.Caption + "0 "
        Case 7
         Label42.Caption = Label42.Caption + "1 "
        Case 8
         Label42.Caption = Label42.Caption + "0 "
        Case 9
         Label42.Caption = Label42.Caption
        Case 10
         Label42.Caption = Label42.Caption + "0 "
        Case 11
         Label42.Caption = Label42.Caption + "1 "
        Case 12
         Label42.Caption = Label42.Caption + "0 "
        Case 13
         Label42.Caption = Label42.Caption
        Case 14
         Label42.Caption = Label42.Caption + "0 "
        Case 15
         Label42.Caption = Label42.Caption + "1 "
        Case 16
         Label42.Caption = Label42.Caption + "0 "
        Case 17
         Label42.Caption = Label42.Caption + "1 "
        Case 18
         Label42.Caption = Label42.Caption
        Case 19
         Label42.Caption = Label42.Caption + "1 "
        Case 20
         Label42.Caption = Label42.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label43.Caption = Label43.Caption + "1 "
        Case 2
         Label43.Caption = Label43.Caption
        Case 3
         Label43.Caption = Label43.Caption + "1 "
        Case 4
         Label43.Caption = Label43.Caption + "0 "

        Case 6
         Label43.Caption = Label43.Caption + "0 "
        Case 7
         Label43.Caption = Label43.Caption + "1 "
        Case 8
         Label43.Caption = Label43.Caption + "0 "
        Case 9
         Label43.Caption = Label43.Caption
        Case 10
         Label43.Caption = Label43.Caption + "0 "
        Case 11
         Label43.Caption = Label43.Caption + "1 "
        Case 12
         Label43.Caption = Label43.Caption + "0 "
        Case 13
         Label43.Caption = Label43.Caption
        Case 14
         Label43.Caption = Label43.Caption + "0 "
        Case 15
         Label43.Caption = Label43.Caption + "1 "
        Case 16
         Label43.Caption = Label43.Caption + "0 "
        Case 17
         Label43.Caption = Label43.Caption + "1 "
        Case 18
         Label43.Caption = Label43.Caption
        Case 19
         Label43.Caption = Label43.Caption + "1 "
        Case 20
         Label43.Caption = Label43.Caption
    End Select
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label44.Caption = Label44.Caption + "1 "
        Case 2
         Label44.Caption = Label44.Caption
        Case 3
         Label44.Caption = Label44.Caption + "1 "
        Case 4
         Label44.Caption = Label44.Caption + "0 "

        Case 6
         Label44.Caption = Label44.Caption + "0 "
        Case 7
         Label44.Caption = Label44.Caption + "1 "
        Case 8
         Label44.Caption = Label44.Caption + "0 "
        Case 9
         Label44.Caption = Label44.Caption
        Case 10
         Label44.Caption = Label44.Caption + "0 "
        Case 11
         Label44.Caption = Label44.Caption + "1 "
        Case 12
         Label44.Caption = Label44.Caption + "0 "
        Case 13
         Label44.Caption = Label44.Caption
        Case 14
         Label44.Caption = Label44.Caption + "0 "
        Case 15
         Label44.Caption = Label44.Caption + "1 "
        Case 16
         Label44.Caption = Label44.Caption + "0 "
        Case 17
         Label44.Caption = Label44.Caption + "1 "
        Case 18
         Label44.Caption = Label44.Caption
        Case 19
         Label44.Caption = Label44.Caption + "1 "
        Case 20
         Label44.Caption = Label44.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label45.Caption = Label45.Caption + "1 "
        Case 2
         Label45.Caption = Label45.Caption
        Case 3
         Label45.Caption = Label45.Caption + "1 "
        Case 4
         Label45.Caption = Label45.Caption + "0 "

        Case 6
         Label45.Caption = Label45.Caption + "0 "
        Case 7
         Label45.Caption = Label45.Caption + "1 "
        Case 8
         Label45.Caption = Label45.Caption + "0 "
        Case 9
         Label45.Caption = Label45.Caption
        Case 10
         Label45.Caption = Label45.Caption + "0 "
        Case 11
         Label45.Caption = Label45.Caption + "1 "
        Case 12
         Label45.Caption = Label45.Caption + "0 "
        Case 13
         Label45.Caption = Label45.Caption
        Case 14
         Label45.Caption = Label45.Caption + "0 "
        Case 15
         Label45.Caption = Label45.Caption + "1 "
        Case 16
         Label45.Caption = Label45.Caption + "0 "
        Case 17
         Label45.Caption = Label45.Caption + "1 "
        Case 18
         Label45.Caption = Label45.Caption
        Case 19
         Label45.Caption = Label45.Caption + "1 "
        Case 20
         Label45.Caption = Label45.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label46.Caption = Label46.Caption + "1 "
        Case 2
         Label46.Caption = Label46.Caption
        Case 3
         Label46.Caption = Label46.Caption + "1 "
        Case 4
         Label46.Caption = Label46.Caption + "0 "

        Case 6
         Label46.Caption = Label46.Caption + "0 "
        Case 7
         Label46.Caption = Label46.Caption + "1 "
        Case 8
         Label46.Caption = Label46.Caption + "0 "
        Case 9
         Label46.Caption = Label46.Caption
        Case 10
         Label46.Caption = Label46.Caption + "0 "
        Case 11
         Label46.Caption = Label46.Caption + "1 "
        Case 12
         Label46.Caption = Label46.Caption + "0 "
        Case 13
         Label46.Caption = Label46.Caption
        Case 14
         Label46.Caption = Label46.Caption + "0 "
        Case 15
         Label46.Caption = Label46.Caption + "1 "
        Case 16
         Label46.Caption = Label46.Caption + "0 "
        Case 17
         Label46.Caption = Label46.Caption + "1 "
        Case 18
         Label46.Caption = Label46.Caption
        Case 19
         Label46.Caption = Label46.Caption + "1 "
        Case 20
         Label46.Caption = Label46.Caption
    End Select
    
    
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label47.Caption = Label47.Caption + "1 "
        Case 2
         Label47.Caption = Label47.Caption
        Case 3
         Label47.Caption = Label47.Caption + "1 "
        Case 4
         Label47.Caption = Label47.Caption + "0 "

        Case 6
         Label47.Caption = Label47.Caption + "0 "
        Case 7
         Label47.Caption = Label47.Caption + "1 "
        Case 8
         Label47.Caption = Label47.Caption + "0 "
        Case 9
         Label47.Caption = Label47.Caption
        Case 10
         Label47.Caption = Label47.Caption + "0 "
        Case 11
         Label47.Caption = Label47.Caption + "1 "
        Case 12
         Label47.Caption = Label47.Caption + "0 "
        Case 13
         Label47.Caption = Label47.Caption
        Case 14
         Label47.Caption = Label47.Caption + "0 "
        Case 15
         Label47.Caption = Label47.Caption + "1 "
        Case 16
         Label47.Caption = Label47.Caption + "0 "
        Case 17
         Label47.Caption = Label47.Caption + "1 "
        Case 18
         Label47.Caption = Label47.Caption
        Case 19
         Label47.Caption = Label47.Caption + "1 "
        Case 20
         Label47.Caption = Label47.Caption
    End Select
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label48.Caption = Label48.Caption + "1 "
        Case 2
         Label48.Caption = Label48.Caption
        Case 3
         Label48.Caption = Label48.Caption + "1 "
        Case 4
         Label48.Caption = Label48.Caption + "0 "

        Case 6
         Label48.Caption = Label48.Caption + "0 "
        Case 7
         Label48.Caption = Label48.Caption + "1 "
        Case 8
         Label48.Caption = Label48.Caption + "0 "
        Case 9
         Label48.Caption = Label48.Caption
        Case 10
         Label48.Caption = Label48.Caption + "0 "
        Case 11
         Label48.Caption = Label48.Caption + "1 "
        Case 12
         Label48.Caption = Label48.Caption + "0 "
        Case 13
         Label48.Caption = Label48.Caption
        Case 14
         Label48.Caption = Label48.Caption + "0 "
        Case 15
         Label48.Caption = Label48.Caption + "1 "
        Case 16
         Label48.Caption = Label48.Caption + "0 "
        Case 17
         Label48.Caption = Label48.Caption + "1 "
        Case 18
         Label48.Caption = Label48.Caption
        Case 19
         Label48.Caption = Label48.Caption + "1 "
        Case 20
         Label48.Caption = Label48.Caption
    End Select
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label49.Caption = Label49.Caption + "1 "
        Case 2
         Label49.Caption = Label49.Caption
        Case 3
         Label49.Caption = Label49.Caption + "1 "
        Case 4
         Label49.Caption = Label49.Caption + "0 "

        Case 6
         Label49.Caption = Label49.Caption + "0 "
        Case 7
         Label49.Caption = Label49.Caption + "1 "
        Case 8
         Label49.Caption = Label49.Caption + "0 "
        Case 9
         Label49.Caption = Label49.Caption
        Case 10
         Label49.Caption = Label49.Caption + "0 "
        Case 11
         Label49.Caption = Label49.Caption + "1 "
        Case 12
         Label49.Caption = Label49.Caption + "0 "
        Case 13
         Label49.Caption = Label49.Caption
        Case 14
         Label49.Caption = Label49.Caption + "0 "
        Case 15
         Label49.Caption = Label49.Caption + "1 "
        Case 16
         Label49.Caption = Label49.Caption + "0 "
        Case 17
         Label49.Caption = Label49.Caption + "1 "
        Case 18
         Label49.Caption = Label49.Caption
        Case 19
         Label49.Caption = Label49.Caption + "1 "
        Case 20
         Label49.Caption = Label49.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label50.Caption = Label50.Caption + "1 "
        Case 2
         Label50.Caption = Label50.Caption
        Case 3
         Label50.Caption = Label50.Caption + "1 "
        Case 4
         Label50.Caption = Label50.Caption + "0 "

        Case 6
         Label50.Caption = Label50.Caption + "0 "
        Case 7
         Label50.Caption = Label50.Caption + "1 "
        Case 8
         Label50.Caption = Label50.Caption + "0 "
        Case 9
         Label50.Caption = Label50.Caption
        Case 10
         Label50.Caption = Label50.Caption + "0 "
        Case 11
         Label50.Caption = Label50.Caption + "1 "
        Case 12
         Label50.Caption = Label50.Caption + "0 "
        Case 13
         Label50.Caption = Label50.Caption
        Case 14
         Label50.Caption = Label50.Caption + "0 "
        Case 15
         Label50.Caption = Label50.Caption + "1 "
        Case 16
         Label50.Caption = Label50.Caption + "0 "
        Case 17
         Label50.Caption = Label50.Caption + "1 "
        Case 18
         Label50.Caption = Label50.Caption
        Case 19
         Label50.Caption = Label50.Caption + "1 "
        Case 20
         Label50.Caption = Label50.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label51.Caption = Label51.Caption + "1 "
        Case 2
         Label51.Caption = Label51.Caption
        Case 3
         Label51.Caption = Label51.Caption + "1 "
        Case 4
         Label51.Caption = Label51.Caption + "0 "

        Case 6
         Label51.Caption = Label51.Caption + "0 "
        Case 7
         Label51.Caption = Label51.Caption + "1 "
        Case 8
         Label51.Caption = Label51.Caption + "0 "
        Case 9
         Label51.Caption = Label51.Caption
        Case 10
         Label51.Caption = Label51.Caption + "0 "
        Case 11
         Label51.Caption = Label51.Caption + "1 "
        Case 12
         Label51.Caption = Label51.Caption + "0 "
        Case 13
         Label51.Caption = Label51.Caption
        Case 14
         Label51.Caption = Label51.Caption + "0 "
        Case 15
         Label51.Caption = Label51.Caption + "1 "
        Case 16
         Label51.Caption = Label51.Caption + "0 "
        Case 17
         Label51.Caption = Label51.Caption + "1 "
        Case 18
         Label51.Caption = Label51.Caption
        Case 19
         Label51.Caption = Label51.Caption + "1 "
        Case 20
         Label51.Caption = Label51.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label52.Caption = Label52.Caption + "1 "
        Case 2
         Label52.Caption = Label52.Caption
        Case 3
         Label52.Caption = Label52.Caption + "1 "
        Case 4
         Label52.Caption = Label52.Caption + "0 "

        Case 6
         Label52.Caption = Label52.Caption + "0 "
        Case 7
         Label52.Caption = Label52.Caption + "1 "
        Case 8
         Label52.Caption = Label52.Caption + "0 "
        Case 9
         Label52.Caption = Label52.Caption
        Case 10
         Label52.Caption = Label52.Caption + "0 "
        Case 11
         Label52.Caption = Label52.Caption + "1 "
        Case 12
         Label52.Caption = Label52.Caption + "0 "
        Case 13
         Label52.Caption = Label52.Caption
        Case 14
         Label52.Caption = Label52.Caption + "0 "
        Case 15
         Label52.Caption = Label52.Caption + "1 "
        Case 16
         Label52.Caption = Label52.Caption + "0 "
        Case 17
         Label52.Caption = Label52.Caption + "1 "
        Case 18
         Label52.Caption = Label52.Caption
        Case 19
         Label52.Caption = Label52.Caption + "1 "
        Case 20
         Label52.Caption = Label52.Caption
    End Select
    
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label53.Caption = Label53.Caption + "1 "
        Case 2
         Label53.Caption = Label53.Caption
        Case 3
         Label53.Caption = Label53.Caption + "1 "
        Case 4
         Label53.Caption = Label53.Caption + "0 "

        Case 6
         Label53.Caption = Label53.Caption + "0 "
        Case 7
         Label53.Caption = Label53.Caption + "1 "
        Case 8
         Label53.Caption = Label53.Caption + "0 "
        Case 9
         Label53.Caption = Label53.Caption
        Case 10
         Label53.Caption = Label53.Caption + "0 "
        Case 11
         Label53.Caption = Label53.Caption + "1 "
        Case 12
         Label53.Caption = Label53.Caption + "0 "
        Case 13
         Label53.Caption = Label53.Caption
        Case 14
         Label53.Caption = Label53.Caption + "0 "
        Case 15
         Label53.Caption = Label53.Caption + "1 "
        Case 16
         Label53.Caption = Label53.Caption + "0 "
        Case 17
         Label53.Caption = Label53.Caption + "1 "
        Case 18
         Label53.Caption = Label53.Caption
        Case 19
         Label53.Caption = Label53.Caption + "1 "
        Case 20
         Label53.Caption = Label53.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label54.Caption = Label54.Caption + "1 "
        Case 2
         Label54.Caption = Label54.Caption
        Case 3
         Label54.Caption = Label54.Caption + "1 "
        Case 4
         Label54.Caption = Label54.Caption + "0 "

        Case 6
         Label54.Caption = Label54.Caption + "0 "
        Case 7
         Label54.Caption = Label54.Caption + "1 "
        Case 8
         Label54.Caption = Label54.Caption + "0 "
        Case 9
         Label54.Caption = Label54.Caption
        Case 10
         Label54.Caption = Label54.Caption + "0 "
        Case 11
         Label54.Caption = Label54.Caption + "1 "
        Case 12
         Label54.Caption = Label54.Caption + "0 "
        Case 13
         Label54.Caption = Label54.Caption
        Case 14
         Label54.Caption = Label54.Caption + "0 "
        Case 15
         Label54.Caption = Label54.Caption + "1 "
        Case 16
         Label54.Caption = Label54.Caption + "0 "
        Case 17
         Label54.Caption = Label54.Caption + "1 "
        Case 18
         Label54.Caption = Label54.Caption
        Case 19
         Label54.Caption = Label54.Caption + "1 "
        Case 20
         Label54.Caption = Label54.Caption
    End Select
    
                Select Case Int((Rnd * 30) + 1)
        Case 1
         Label55.Caption = Label55.Caption + "1 "
        Case 2
         Label55.Caption = Label55.Caption
        Case 3
         Label55.Caption = Label55.Caption + "1 "
        Case 4
         Label55.Caption = Label55.Caption + "0 "

        Case 6
         Label55.Caption = Label55.Caption + "0 "
        Case 7
         Label55.Caption = Label55.Caption + "1 "
        Case 8
         Label55.Caption = Label55.Caption + "0 "
        Case 9
         Label55.Caption = Label55.Caption
        Case 10
         Label55.Caption = Label55.Caption + "0 "
        Case 11
         Label55.Caption = Label55.Caption + "1 "
        Case 12
         Label55.Caption = Label55.Caption + "0 "
        Case 13
         Label55.Caption = Label55.Caption
        Case 14
         Label55.Caption = Label55.Caption + "0 "
        Case 15
         Label55.Caption = Label55.Caption + "1 "
        Case 16
         Label55.Caption = Label55.Caption + "0 "
        Case 17
         Label55.Caption = Label55.Caption + "1 "
        Case 18
         Label55.Caption = Label55.Caption
        Case 19
         Label55.Caption = Label55.Caption + "1 "
        Case 20
         Label55.Caption = Label55.Caption
    End Select
    
    
    
End Sub


Private Sub Timer7_Timer()

i = Int((Rnd * 500) + 1)
Controls("Timer7").interval = i
i = Int((Rnd * 55) + 1)
Controls("Label" & i).Caption = Controls("Label" & i).Caption + "º "
i = Int((Rnd * 55) + 1)
Controls("Label" & i).Caption = Controls("Label" & i).Caption + "¹ "
i = Int((Rnd * 55) + 1)
Controls("Label" & i).Caption = Controls("Label" & i).Caption + "ø "
i = Int((Rnd * 55) + 1)
Controls("Label" & i).Caption = Controls("Label" & i).Caption + "× "


End Sub











Private Sub Timer8_Timer()

If Len(Label1.Caption) > 50 Then
Label1.Caption = ""
End If
If Len(Label2.Caption) > 50 Then
Label2.Caption = ""
End If
If Len(Label3.Caption) > 50 Then
Label3.Caption = ""
End If
If Len(Label4.Caption) > 50 Then
Label4.Caption = ""
End If
If Len(Label5.Caption) > 50 Then
Label5.Caption = ""
End If
If Len(Label6.Caption) > 50 Then
Label6.Caption = ""
End If
If Len(Label7.Caption) > 50 Then
Label7.Caption = ""
End If
If Len(Label8.Caption) > 50 Then
Label8.Caption = ""
End If
If Len(Label9.Caption) > 50 Then
Label9.Caption = ""
End If
If Len(Label10.Caption) > 50 Then
Label10.Caption = ""
End If
If Len(Label11.Caption) > 50 Then
Label11.Caption = ""
End If
If Len(Label12.Caption) > 50 Then
Label12.Caption = ""
End If
If Len(Label13.Caption) > 50 Then
Label13.Caption = ""
End If
If Len(Label14.Caption) > 50 Then
Label14.Caption = ""
End If
If Len(Label15.Caption) > 50 Then
Label15.Caption = ""
End If
If Len(Label16.Caption) > 50 Then
Label16.Caption = ""
End If
If Len(Label17.Caption) > 50 Then
Label17.Caption = ""
End If
If Len(Label18.Caption) > 50 Then
Label18.Caption = ""
End If
If Len(Label19.Caption) > 50 Then
Label19.Caption = ""
End If
If Len(Label20.Caption) > 50 Then
Label20.Caption = ""
End If
If Len(Label21.Caption) > 50 Then
Label21.Caption = ""
End If
If Len(Label22.Caption) > 50 Then
Label22.Caption = ""
End If
If Len(Label23.Caption) > 50 Then
Label23.Caption = ""
End If
If Len(Label24.Caption) > 50 Then
Label24.Caption = ""
End If
If Len(Label25.Caption) > 50 Then
Label25.Caption = ""
End If
If Len(Label26.Caption) > 50 Then
Label26.Caption = ""
End If
If Len(Label27.Caption) > 50 Then
Label27.Caption = ""
End If
If Len(Label28.Caption) > 50 Then
Label28.Caption = ""
End If
If Len(Label29.Caption) > 50 Then
Label29.Caption = ""
End If
If Len(Label30.Caption) > 50 Then
Label30.Caption = ""
End If
If Len(Label31.Caption) > 50 Then
Label31.Caption = ""
End If
If Len(Label32.Caption) > 50 Then
Label32.Caption = ""
End If
If Len(Label33.Caption) > 50 Then
Label33.Caption = ""
End If
If Len(Label34.Caption) > 50 Then
Label34.Caption = ""
End If
If Len(Label35.Caption) > 50 Then
Label35.Caption = ""
End If
If Len(Label36.Caption) > 50 Then
Label36.Caption = ""
End If
If Len(Label37.Caption) > 50 Then
Label37.Caption = ""
End If
If Len(Label38.Caption) > 50 Then
Label38.Caption = ""
End If
If Len(Label39.Caption) > 50 Then
Label39.Caption = ""
End If
If Len(Label40.Caption) > 50 Then
Label40.Caption = ""
End If
If Len(Label41.Caption) > 50 Then
Label41.Caption = ""
End If
If Len(Label42.Caption) > 50 Then
Label42.Caption = ""
End If
If Len(Label43.Caption) > 50 Then
Label43.Caption = ""
End If
If Len(Label44.Caption) > 50 Then
Label44.Caption = ""
End If
If Len(Label45.Caption) > 50 Then
Label45.Caption = ""
End If
If Len(Label46.Caption) > 50 Then
Label46.Caption = ""
End If
If Len(Label47.Caption) > 50 Then
Label47.Caption = ""
End If
If Len(Label48.Caption) > 50 Then
Label48.Caption = ""
End If
If Len(Label49.Caption) > 50 Then
Label49.Caption = ""
End If
If Len(Label50.Caption) > 50 Then
Label50.Caption = ""
End If
If Len(Label51.Caption) > 50 Then
Label51.Caption = ""
End If
If Len(Label52.Caption) > 50 Then
Label52.Caption = ""
End If
If Len(Label53.Caption) > 50 Then
Label53.Caption = ""
End If
If Len(Label54.Caption) > 50 Then
Label54.Caption = ""
End If
If Len(Label55.Caption) > 50 Then
Label55.Caption = ""
End If








End Sub
