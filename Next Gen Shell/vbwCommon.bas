Attribute VB_Name = "vbwCommon"
Option Explicit

' Options '
Public vbwException As Boolean
Public vbwTraceProc As Boolean
Public vbwTraceParameters As Boolean
Public vbwTraceLine As Boolean
Public vbwDebugger As Boolean
Public vbwLog As Boolean
Public vbwDebugPrint As Boolean
Public vbwInstanceCount As Boolean
Public vbwEmailRecipientName As String
Public vbwEmailRecipientAdress As String
Public vbwDumpStringMaxLength As Long

' Call Stack
Public vbwStackCalls() As String ' array containing each call of the stack
Dim nStackCalls As Long ' number of calls

' Trace
Dim nTraceCalls As Long ' number of calls

' Debugger
Private Type COPYDATASTRUCT
    Handle As Long
    Length As Long
    Message As Long
End Type
Public Enum Enum_MessageType
    APP_TITLE
    VBW_ENTER_PROC
    VBW_EXIT_PROC
    VBW_DEBUG
    VBW_DECLARE_INSTANCE
    VBW_NEW_INSTANCE
    VBW_KILL_INSTANCE
    VBW_STOP
    VBW_LINE
    VBW_LINE_ENCRYPTED
End Enum
Private Const WM_COPYDATA = &H4A
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim DebuggerID As Long ' Unique Identifier for the Debugger

' Profiler
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Dim secTime  As Currency
Dim secFreq As Currency
Dim secTraceInOverhead  As Currency
Dim secTraceOutOverhead As Currency

' Log File
Dim fIsLogInitialize As Boolean
Public vbwLogFile As String
Dim fLogFileOpen As Boolean
Public vbwLogFileNum As Long

' Var Dump
Const VBW_STRING = "**************************"
Global Const VBW_LOCAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* LOCAL LEVEL VARIABLES  *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_MODULE_STRING = vbCrLf & VBW_STRING & vbCrLf & "* MODULE LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_GLOBAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* GLOBAL LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_TYPE_STRING = " (User Defined Type Array)"
Global Const VBW_UNKNOWN_STRING = " = {Unknown Type}"
Global Const VBW_LOCAL_NOT_REPORTED = "Local Variables: not reported"
Global Const VBW_MODULE_NOT_REPORTED = "Module Variables: not reported"
Global Const VBW_GLOBAL_NOT_REPORTED = "Global Variables: not reported"
Global Const VBW_NO_LOCAL_VARIABLES = "No Local Variables"
Global vbwDumpFile As String
Global vbwDumpFileNum As Long

' Thread & processes
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

' Exception handling declarations
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long    ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type
Private Type EXCEPTION_DEBUG_INFO
        pExceptionRecord As EXCEPTION_RECORD
        dwFirstChance As Long
End Type
Private Type CONTEXT
    dblVar(66) As Double ' The real structure is more complex
    lngVar(6) As Long    ' but we don't need those details
End Type
Private Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type
Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_PRIV_INSTRUCTION = &HC0000096
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_GUARD_PAGE = &H80000001
Private Const EXCEPTION_INVALID_HANDLE = &HC0000008
Private Const CONTROL_C_EXIT = &HC000013A

Global VBW_EMPTY As Variant ' for use with vbwExecuteLine() in IIf structures

'Const VBW_EXE_EXTENSION = ".exe" ' this line will be rewritten by VB Watch with the right extension

' vbwNoTraceProc vbwNoTraceLine ' don't remove this !

' <VB WATCH>
Const VBWMODULE = "vbwCommon"
Global Const VBWPROJECT = "Project1"
Global Const VBW_EXE_EXTENSION = ".exe"
' </VB WATCH>

Sub vbwInitialize()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

2          Static vbwIsInitialized As Boolean

3          If vbwIsInitialized Then
4              Exit Sub
5          End If

6          vbwSetOptions

7          vbwLogFile = App.path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & ".log"
8          vbwDumpFile = App.path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & ".dmp"

9          If vbwDebugger Then
               ' declare useful infos for the Debugger: threadID, full exe path , command line
10             vbwSendDebugger APP_TITLE, App.ThreadID & "|" & IIf(vbwIsInIDE, "IDE", App.path) & "\" & App.EXEName & VBW_EXE_EXTENSION & "|" & Command$
               ' declare forms, classes and user controls of the project
11             vbwDeclareInstances
12         End If

13         If vbwException Then
14             vbwHandleException
15         End If

16         vbwDumpStringMaxLength = 128 ' change this value to suit your need - make it 0 to remove the size check (to use with caution)

17         vbwIsInitialized = True

' <VB WATCH>
18         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwInitialize"
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

' </VB WATCH>
End Sub

Sub vbwReportVariable(ByVal lName As String, ByVal lValue As Variant, Optional ByVal lTab As Long)
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
19         On Error GoTo vbwErrHandler
' </VB WATCH>
20         Dim i As Long, j As Long, k As Long, L As Long
21         Dim tDim As Long

22         On Error GoTo ErrDump

23         If VBA.InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
24             tDim = GetArrayDimension(lValue)
25             Select Case tDim
                   Case 1
26                     vbwReportToFile String$(lTab, vbTab) & "Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & ") As " & TypeName(lValue)
27                     For i = LBound(lValue) To UBound(lValue)
28                         vbwReportVariable lName & "(" & i & ")", lValue(i), lTab + 1
29                     Next i
30                 Case 2
31                     vbwReportToFile String$(lTab, vbTab) & "Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & ") As " & TypeName(lValue)
32                     For j = LBound(lValue, 2) To UBound(lValue, 2)
33                         For i = LBound(lValue, 1) To UBound(lValue, 1)
34                             vbwReportVariable lName & "(" & i & "," & j & ")", lValue(i, j), lTab + 1
35                         Next i
36                     Next j
37                 Case 3
38                     vbwReportToFile String$(lTab, vbTab) & "Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & ") As " & TypeName(lValue)
39                     For k = LBound(lValue, 3) To UBound(lValue, 3)
40                         For j = LBound(lValue, 2) To UBound(lValue, 2)
41                             For i = LBound(lValue, 1) To UBound(lValue, 1)
42                                 vbwReportVariable lName & "(" & i & "," & j & "," & k & ")", lValue(i, j, k), lTab + 1
43                             Next i
44                         Next j
45                     Next k
46                 Case 4
47                     vbwReportToFile String$(lTab, vbTab) & "Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & "," & LBound(lValue, 4) & " To " & UBound(lValue, 4) & ") As " & TypeName(lValue)
48                     For L = LBound(lValue, 4) To UBound(lValue, 4)
49                         For k = LBound(lValue, 3) To UBound(lValue, 3)
50                             For j = LBound(lValue, 2) To UBound(lValue, 2)
51                                 For i = LBound(lValue, 1) To UBound(lValue, 1)
52                                     vbwReportVariable lName & "(" & i & "," & j & "," & k & "," & L & ")", lValue(i, j, k, L), lTab + 1
53                                 Next i
54                             Next j
55                         Next k
56                     Next L
57                 Case Else
58                     vbwReportToFile String$(lTab, vbTab) & "Array " & lName & "() not processed: " & tDim & " dimensions"
59             End Select
60         Else
               ' non-array '
61             If IsObject(lValue) Then
62                 vbwReportObject lName, lValue, lTab
63             Else
64                 If VarType(lValue) = vbString Then
65                     lValue = FormatString(lValue)
66                 End If
67                 vbwReportToFile String$(lTab, vbTab) & lName & " = " & lValue & " (" & TypeName(lValue) & ")"
68             End If
69         End If
70         Exit Sub

71 ErrDump:
72         Err.Clear
73         vbwReportToFile String$(lTab, vbTab) & lName & " = {Variable Dumping Error}"
' <VB WATCH>
74         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwReportVariable"
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

' </VB WATCH>
End Sub

Sub vbwReportObject(ByVal lName As String, ByVal lObject As Object, Optional ByVal lTab As Long)
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
75         On Error GoTo vbwErrHandler
' </VB WATCH>

76         On Error GoTo ErrDump
77         If TypeOf lObject Is Form Or TypeOf lObject Is MDIForm Then
78            On Error Resume Next
79            vbwReportToFile "Form " & lName
80            Dim C As Control
81            For Each C In lObject.Controls
82                vbwReportObject C.name & vbwGetIndex(C), C, 1
83            Next C
84         Else
85             If IsNumeric(lObject) Then
86                 vbwReportVariable lName, CDbl(lObject), lTab
87             Else
88                 vbwReportVariable lName, CStr(lObject), lTab
89             End If
90         End If
91         Exit Sub

92 ErrDump:
93         Err.Clear
94         vbwReportToFile String$(lTab, vbTab) & lName & ".Value = {No Value Property}"
' <VB WATCH>
95         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwReportObject"
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

' </VB WATCH>
End Sub

' <BEGIN CHANGE 03/03/2001>
Function vbwReportParameter(ByVal lName As String, ByRef lValue As Variant) As String
       ' <END CHANGE 03/03/2001>
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
96         On Error GoTo vbwErrHandler
' </VB WATCH>
97         Dim i As Long, j As Long, k As Long
98         Dim tDim As Long
99         Dim retString As String

100        On Error GoTo ErrDump

101        If VBA.InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
102            tDim = GetArrayDimension(lValue)
103            If tDim Then
104                retString = lName & "("
105                For i = 1 To tDim
106                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
107                Next i
108                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
109            Else
110                retString = lName & "(Undimensioned Array)"
111            End If
112        Else
               ' non-array '
113            If IsObject(lValue) Then
                   ' object
114                On Error Resume Next
115                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
116                If Err.number Then
117                    On Error GoTo ErrDump
118                    retString = TypeName(lValue) & " " & lName & " = " & lValue.name & vbwGetIndex(lValue)
119                End If
120            Else
       ' <BEGIN CHANGE 03/24/2001>
                   ' non-object
121                If VarType(lValue) = vbString Then
122                   retString = lName & " = " & FormatString(lValue)
123                Else
124                   retString = lName & " = " & lValue
125                End If
       ' <END CHANGE 03/24/2001>
126            End If
127        End If

128        vbwReportParameter = retString
129        Exit Function

130 ErrDump:
131        Err.Clear
132        vbwReportParameter = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
' <VB WATCH>
133        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwReportParameter"
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

' </VB WATCH>
End Function

' <BEGIN CHANGE 06/18/2001>
Function vbwReportParameterByVal(ByVal lName As String, ByVal lValue As Variant) As String
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
134        On Error GoTo vbwErrHandler
' </VB WATCH>
135        Dim i As Long, j As Long, k As Long
136        Dim tDim As Long
137        Dim retString As String

138        On Error GoTo ErrDump

139        If VBA.InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
140            tDim = GetArrayDimension(lValue)
141            If tDim Then
142                retString = lName & "("
143                For i = 1 To tDim
144                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
145                Next i
146                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
147            Else
148                retString = lName & "(Undimensioned Array)"
149            End If
150        Else
               ' non-array '
151            If IsObject(lValue) Then
                   ' object
152                On Error Resume Next
153                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
154                If Err.number Then
155                    On Error GoTo ErrDump
156                    retString = TypeName(lValue) & " " & lName & " = " & lValue.name & vbwGetIndex(lValue)
157                End If
158            Else
                   ' non-object
159                If VarType(lValue) = vbString Then
160                   retString = lName & " = " & FormatString(lValue)
161                Else
162                   retString = lName & " = " & lValue
163                End If
164            End If
165        End If

166        vbwReportParameterByVal = retString
167        Exit Function

168 ErrDump:
169        Err.Clear
170        vbwReportParameterByVal = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
' <VB WATCH>
171        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwReportParameterByVal"
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

' </VB WATCH>
End Function
' <END CHANGE 06/18/2001>

Sub vbwReportToFile(ByRef lString As String)
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
172        On Error GoTo vbwErrHandler
' </VB WATCH>
173         Print #vbwDumpFileNum, lString
' <VB WATCH>
174        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwReportToFile"
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

' </VB WATCH>
End Sub

Sub vbwOpenDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
175        On Error GoTo vbwErrHandler
' </VB WATCH>
176       vbwDumpFileNum = FreeFile
177       Open vbwDumpFile For Output As #vbwDumpFileNum
' <VB WATCH>
178        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwOpenDumpFile"
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

' </VB WATCH>
End Sub

Sub vbwCloseDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
179        On Error GoTo vbwErrHandler
' </VB WATCH>
180       Close #vbwDumpFileNum
' <VB WATCH>
181        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwCloseDumpFile"
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

' </VB WATCH>
End Sub

' <BEGIN CHANGE 03/03/2001>
Private Function GetArrayDimension(ByRef arg As Variant) As Long
       ' <END CHANGE 03/03/2001>
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
182        On Error GoTo vbwErrHandler
' </VB WATCH>
183        Dim i As Long, j As Long
184        On Error Resume Next
185        i = 0
186        Do
187            i = i + 1
188            j = LBound(arg, i)
189        Loop Until Err.number
190        GetArrayDimension = i - 1
' <VB WATCH>
191        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetArrayDimension"
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

' </VB WATCH>
End Function

Function vbwGetIndex(tObject As Variant) As String
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
192        On Error GoTo vbwErrHandler
' </VB WATCH>
193        On Error Resume Next
194        vbwGetIndex = "(" & tObject.Index & ")"
' <VB WATCH>
195        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwGetIndex"
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

' </VB WATCH>
End Function

Private Function FormatString(ByVal arg As String) As String
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
196        On Error GoTo vbwErrHandler
' </VB WATCH>

197        If Right$(arg, 1) = "}" Then ' probably a VB Watch built-in message
198             FormatString = arg
199             Exit Function
200        End If

           ' 1. truncate according to the vbwDumpStringMaxLength value
201        If vbwDumpStringMaxLength Then
202            If Len(arg) > vbwDumpStringMaxLength Then
203                arg = Left$(arg, vbwDumpStringMaxLength + 1)   ' +1: avoids to cut inside a vbCrLf '
204                If Right$(arg, 2) = vbCrLf Then
                       ' don't cut inside a vbCrLf
205                Else
206                    arg = Left$(arg, vbwDumpStringMaxLength)
207                End If
208                arg = arg & "{...}" ' truncated
209            End If
210        End If

           ' 2. make sure string isn't multiline
211        arg = VBA.Replace(arg, vbCrLf, "<CrLf>", , , vbBinaryCompare)
212        arg = VBA.Replace(arg, Chr(13), "<Cr>", , , vbBinaryCompare)
213        arg = VBA.Replace(arg, Chr(10), "<Lf>", , , vbBinaryCompare)

           ' 3. add quotes
214        FormatString = Chr(34) & arg & Chr(34)
' <VB WATCH>
215        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FormatString"
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

' </VB WATCH>
End Function

Sub vbwTraceIn(ByRef lProc As String, Optional ByRef lParameters As String)
' <VB WATCH>
216        On Error GoTo vbwErrHandler
' </VB WATCH>

217        nTraceCalls = nTraceCalls + 1

218        nStackCalls = nStackCalls + 1
219        ReDim Preserve vbwStackCalls(1 To nStackCalls)
220        vbwStackCalls(nStackCalls) = lProc

221        Dim lString As String
222        lString = String$(nTraceCalls - 1, vbTab) & lProc

223        If vbwLog Then
224             vbwSendLog lString & lParameters
225        End If

226        If vbwDebugger Then
227             vbwSendDebugger VBW_ENTER_PROC, lProc & lParameters
228        End If

' <VB WATCH>
229        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwTraceIn"
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

' </VB WATCH>
End Sub

Sub vbwTraceOut(ByRef lProc As String)
' <VB WATCH>
230        On Error GoTo vbwErrHandler
' </VB WATCH>

231        If nTraceCalls > 0 Then ' should always be true
232           nTraceCalls = nTraceCalls - 1
233        End If

234        If nStackCalls > 0 Then ' should always be true
235           nStackCalls = nStackCalls - 1
236        End If

237        If vbwDebugger Then
238             vbwSendDebugger VBW_EXIT_PROC, lProc
239        End If
' <VB WATCH>
240        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwTraceOut"
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

' </VB WATCH>
End Sub


Function vbwExecuteLine(ByRef fEncrypted As String, ByRef lLine As String) As Boolean
' <VB WATCH>
241        On Error GoTo vbwErrHandler
' </VB WATCH>

242        If vbwTraceLine Then

243            If vbwLog Then
244                If nTraceCalls > 0 Then
245                    vbwSendLog String$(nTraceCalls - 1, vbTab) & " -> " & lLine
246                Else
247                    vbwSendLog " -> " & lLine
248                End If
249            End If

250            If vbwDebugger Then
251                If fEncrypted Then
252                    vbwSendDebugger VBW_LINE_ENCRYPTED, lLine
253                Else
254                    vbwSendDebugger VBW_LINE, lLine
255                End If
256            End If

257        End If

           ' This function always returns false
' <VB WATCH>
258        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExecuteLine"
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

' </VB WATCH>
End Function

Function vbwGetStack() As String
' <VB WATCH>
259        On Error GoTo vbwErrHandler
' </VB WATCH>

260        If vbwTraceProc = False Then
261            vbwGetStack = "{Unavailable}"
262            Exit Function
263        End If

264        Dim vbwStackString As String
265        Dim i As Long

266        For i = nStackCalls To 1 Step -1
267            vbwStackString = vbwStackString & String$(i - 1, vbTab) & vbwStackCalls(i) & vbCrLf
268        Next i
269        vbwGetStack = IIf(vbwStackString <> "", vbwStackString, "{Empty}")
' <VB WATCH>
270        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwGetStack"
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

' </VB WATCH>
End Function

Sub vbwSendDebugger(ByRef dwData As Enum_MessageType, ByVal lMsg As String)
' <VB WATCH>
271        On Error GoTo vbwErrHandler
' </VB WATCH>

272        Dim cds  As COPYDATASTRUCT
273        Dim new_hW As Long
274        Static hW As Long
275        Dim i As Long
276        Dim Ret As Long
277        Dim buffer(1 To 1024) As Byte ' can send 1024 bytes at a time

278        new_hW = VBA.Val(VBA.GetSetting("VB Watch", "Debugger", "hWnd", "0")) ' retrieves the handle that the Debugger stored for us

279        If new_hW <> hW Then
               ' new Debugger session !
280            hW = new_hW
281            If dwData <> APP_TITLE Then
                   ' declare us
282                DebuggerID = 0 ' request a new ID
283                vbwSendDebugger APP_TITLE, App.ThreadID & "|" & IIf(vbwIsInIDE, "IDE", App.path) & "\" & App.EXEName & VBW_EXE_EXTENSION & "|" & Command$
284            End If
285        End If

286        If hW Then
               ' get unique Identifier
287            If DebuggerID = 0 Then
288                cds.Handle = -1 ' -1 = request new ID to be written into the registry
289                cds.Length = 1
290                cds.Message = VBA.VarPtr(buffer(1))
291                Ret = SendMessage(hW, WM_COPYDATA, 0&, cds)
292                If Ret = -999999 Then
293                    DebuggerID = VBA.GetSetting("VB Watch", "Debugger", "NewID", "0")   ' read ID '
294                    If DebuggerID = 0 Then
                           ' shouldn't fail but who knows... '
295                        Randomize
296                        DebuggerID = CLng(Rnd * (2 ^ 30))
297                    End If
298                Else
                       ' no response from Debugger
299                    DebuggerID = 0
300                    Exit Sub
301                End If
302            End If

303            lMsg = Left$(lMsg, 1024) ' reduces the string if too long, to avoid a crash
304            Call CopyMemory(buffer(1), ByVal lMsg, Len(lMsg))   ' Copy the string into a byte array, converting it to ASCII '
305            cds.Handle = dwData
306            cds.Length = Len(lMsg) + 1
307            cds.Message = VBA.VarPtr(buffer(1))       ' VarPtr retrieves a pointer to buffer '
308            Ret = SendMessage(hW, WM_COPYDATA, DebuggerID, cds) ' send message to Debugger
309        End If
' <VB WATCH>
310        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwSendDebugger"
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

' </VB WATCH>
End Sub

Sub vbwSendLog(ByRef tMsg As String)
' <VB WATCH>
311        On Error GoTo vbwErrHandler
' </VB WATCH>

       ' <BEGIN CHANGE 03/03/2001>
312        If Err.number Then
               ' Save Err object before being cleared by "On Error GoTo Err_vbwSendLog"
313            Dim ErrDescription As String, ErrHelpFile As String, ErrSource As String
314            Dim ErrHelpContext As Long, ErrNumber As Long
315            ErrDescription = Err.Description
316            ErrHelpContext = Err.HelpContext
317            ErrHelpFile = Err.HelpFile
318            ErrNumber = Err.number
319            ErrSource = Err.Source
320        End If
       ' <END CHANGE 03/03/2001>

321        On Error GoTo Err_vbwSendLog
322        If Not fLogFileOpen Then
323            fLogFileOpen = True
324            vbwLogFileNum = FreeFile
325            Open vbwLogFile For Output As #vbwLogFileNum
326        End If

327        If Not fIsLogInitialize Then
               ' init file '
328            fIsLogInitialize = True
329            Print #vbwLogFileNum, "Tracing " & App.Title
330            Print #vbwLogFileNum, "Session started " & Now
331            Print #vbwLogFileNum, ""
332        End If

           ' log to file
333        Print #vbwLogFileNum, tMsg

        ' <BEGIN CHANGE 03/03/2001>
334       If ErrNumber Then
               ' Restore Err object if cleared by "On Error GoTo Err_vbwSendLog"
335            Err.Description = ErrDescription
336            Err.HelpContext = ErrHelpContext
337            Err.HelpFile = ErrHelpFile
338            Err.number = ErrNumber
339            Err.Source = ErrSource
340        End If
       ' <END CHANGE 03/03/2001>

341        Exit Sub

342 Err_vbwSendLog:
343        If Err.number = 52 Then
               ' the file was closed ! (probably by a general close statement)
344    On Error GoTo vbwErrHandler
345            vbwLogFileNum = FreeFile
346            Open vbwLogFile For Append As #vbwLogFileNum
347            Resume
348        Else
               ' do as if we didn't intercept this error
349            Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
350        End If
' <VB WATCH>
351        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwSendLog"
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

' </VB WATCH>
End Sub

' Ends a component's thread. If this was the last active thread, ends the component's process.
Public Sub vbwExitThread()
' <VB WATCH>
352        On Error GoTo vbwErrHandler
' </VB WATCH>
353        If vbwIsInIDE Then
               ' Executing ExitThread within the IDE will terminate VB without ceremony !
354            Stop ' Press the End button now
355        Else
356            Dim lpExitCode As Long
357            If GetExitCodeThread(GetCurrentThread(), lpExitCode) Then
358                ExitThread lpExitCode
359            End If
360        End If
' <VB WATCH>
361        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExitThread"
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

' </VB WATCH>
End Sub

' Ends a component's process. Equivalent to the End statement.
Public Sub vbwExitProcess()
' <VB WATCH>
362        On Error GoTo vbwErrHandler
' </VB WATCH>
363        If vbwIsInIDE Then
               ' Executing ExitProcess within the IDE will terminate VB without ceremony !
364            Stop ' Press the End button now
365        Else
366            Dim lpExitCode As Long
367            If GetExitCodeProcess(GetCurrentProcess(), lpExitCode) Then
368                ExitProcess lpExitCode
369            End If
370        End If
' <VB WATCH>
371        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExitProcess"
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

' </VB WATCH>
End Sub

' determines if the program is running in the IDE or an EXE File
Public Function vbwIsInIDE() As Boolean
' <VB WATCH>
372        On Error GoTo vbwErrHandler
' </VB WATCH>

373        Dim strFileName As String
374        Dim lngcount As Long

375        strFileName = String(255, 0)
376        lngcount = GetModuleFileName(App.hInstance, strFileName, 255)
377        strFileName = Left(strFileName, lngcount)

378        vbwIsInIDE = UCase$(Right$(strFileName, 8)) Like "\VB#.EXE"

' <VB WATCH>
379        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwIsInIDE"
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

' </VB WATCH>
End Function

' Exception handling stuff
Public Sub vbwHandleException()
           ' Exceptions will be caught and redirected to the failing procedure
' <VB WATCH>
380        On Error GoTo vbwErrHandler
' </VB WATCH>
381        SetUnhandledExceptionFilter AddressOf vbwExceptionFilter
' <VB WATCH>
382        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwHandleException"
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

' </VB WATCH>
End Sub

' Exception handling stuff
Public Sub vbwUnHandleException()
           ' Exceptions are no longer caught and will cause Exceptions
           ' Whenever possible, call this procedure before returning to the VB's IDE
' <VB WATCH>
383        On Error GoTo vbwErrHandler
' </VB WATCH>
384        SetUnhandledExceptionFilter 0
' <VB WATCH>
385        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwUnHandleException"
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

' </VB WATCH>
End Sub

' Exception handling stuff
Public Function vbwExceptionFilter(ByRef pExceptionInfo As EXCEPTION_POINTERS) As Long
       'vbwNoErrorHandler ' DO NOT remove this !!!

386        Dim ExceptionRecord As EXCEPTION_RECORD
387        ExceptionRecord = pExceptionInfo.pExceptionRecord

388        Do While ExceptionRecord.pExceptionRecord ' Empties the exceptions stack
389            CopyMemory ExceptionRecord, ByVal ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
390        Loop

391        vbwExceptionFilter = EXCEPTION_CONTINUE_EXECUTION

       'vbwExitProc ' because the next instruction causes to exit the function ' ' DO NOT remove this !!!

           ' Convert the exception to a normal VB error and go back to the failing procedure '
392        Err.Raise 65535, , ExceptionDescription(ExceptionRecord.ExceptionCode)

End Function

' Exception handling stuff
Private Function ExceptionDescription(ByVal ExceptionCode As Long) As String
       ' vbwNoErrorHandler ' don't remove this !
393        Select Case ExceptionCode
               Case EXCEPTION_ACCESS_VIOLATION
394                ExceptionDescription = "Exception: Access Violation"
395            Case EXCEPTION_DATATYPE_MISALIGNMENT
396                ExceptionDescription = "Exception: Datatype Misalignment"
397            Case EXCEPTION_BREAKPOINT
398                ExceptionDescription = "Exception: Breakpoint"
399            Case EXCEPTION_SINGLE_STEP
400                ExceptionDescription = "Exception: Single Step"
401            Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
402                ExceptionDescription = "Exception: Array Bounds Exceeded"
403            Case EXCEPTION_FLT_DENORMAL_OPERAND
404                ExceptionDescription = "Exception: Float Denormal Operand"
405            Case EXCEPTION_FLT_DIVIDE_BY_ZERO
406                ExceptionDescription = "Exception: Float Divide By Zero"
407            Case EXCEPTION_FLT_INEXACT_RESULT
408                ExceptionDescription = "Exception: Float Inexact Result"
409            Case EXCEPTION_FLT_INVALID_OPERATION
410                ExceptionDescription = "Exception: Float Invalid Operation"
411            Case EXCEPTION_FLT_OVERFLOW
412                ExceptionDescription = "Exception: Float Overflow"
413            Case EXCEPTION_FLT_STACK_CHECK
414                ExceptionDescription = "Exception: Float Stack Check"
415            Case EXCEPTION_FLT_UNDERFLOW
416                ExceptionDescription = "Exception: Float Underflow"
417            Case EXCEPTION_INT_DIVIDE_BY_ZERO
418                ExceptionDescription = "Exception: Integer Divide By Zero"
419            Case EXCEPTION_INT_OVERFLOW
420                ExceptionDescription = "Exception: Integer Overflow"
421            Case EXCEPTION_PRIV_INSTRUCTION
422                ExceptionDescription = "Exception: Priv Instruction"
423            Case EXCEPTION_IN_PAGE_ERROR
424                ExceptionDescription = "Exception: In Page Error"
425            Case EXCEPTION_ILLEGAL_INSTRUCTION
426                ExceptionDescription = "Exception: Illegal Instruction"
427            Case EXCEPTION_NONCONTINUABLE_EXCEPTION
428                ExceptionDescription = "Exception: Non Continuable Exception"
429            Case EXCEPTION_STACK_OVERFLOW
430                ExceptionDescription = "Exception: Stack Overflow"
431            Case EXCEPTION_INVALID_DISPOSITION
432                ExceptionDescription = "Exception: Invalid Disposition"
433            Case EXCEPTION_GUARD_PAGE
434                ExceptionDescription = "Exception: Guard Page"
435            Case EXCEPTION_INVALID_HANDLE
436                ExceptionDescription = "Exception: Invalid Handle"
437            Case CONTROL_C_EXIT
438                ExceptionDescription = "Exception: Control C Exit"
439            Case Else
440                ExceptionDescription = "Unknown Exception"
441        End Select

End Function

'Instance counting
Private Sub vbwDeclareInstances()
' <VB WATCH>
442        On Error GoTo vbwErrHandler
' </VB WATCH>
443        If vbwDebugger Then
              ' here goes the code that declares the list of all classes/forms/usercontrols present in the project
              ' don't remove the line below !!
              ' <DeclareInstances>
444            vbwSendDebugger VBW_DECLARE_INSTANCE, "Form unseen"
445            vbwSendDebugger VBW_DECLARE_INSTANCE, "User SButton"
446            vbwSendDebugger VBW_DECLARE_INSTANCE, "Form booters"
447            vbwSendDebugger VBW_DECLARE_INSTANCE, "Form password"
448            vbwSendDebugger VBW_DECLARE_INSTANCE, "Form Form1"
449            vbwSendDebugger VBW_DECLARE_INSTANCE, "Form crash"

450        End If
' <VB WATCH>
451        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwDeclareInstances"
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

' </VB WATCH>
End Sub

Public Sub vbwNewInstance(ByRef InstanceName As String, ByRef InstanceID As Long)
           ' attributes a unique ID to this instance
' <VB WATCH>
452        On Error GoTo vbwErrHandler
' </VB WATCH>
453        Static InstanceCounter
454        InstanceCounter = InstanceCounter + 1
455        InstanceID = InstanceCounter

456        If vbwDebugger Then
457             vbwSendDebugger VBW_NEW_INSTANCE, InstanceID & " " & InstanceName
458        End If
' <VB WATCH>
459        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwNewInstance"
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

' </VB WATCH>
End Sub

Public Sub vbwKillInstance(ByRef InstanceName As String, ByRef InstanceID As Long)
' <VB WATCH>
460        On Error GoTo vbwErrHandler
' </VB WATCH>
461        If vbwDebugger Then
462             vbwSendDebugger VBW_KILL_INSTANCE, InstanceID & " " & InstanceName
463        End If
' <VB WATCH>
464        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwKillInstance"
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

' </VB WATCH>
End Sub


