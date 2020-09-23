Attribute VB_Name = "vbwOptions"
Option Explicit

' ShiFtY
' <VB WATCH>
Const VBWMODULE = "vbwOptions"
' </VB WATCH>

Public Sub vbwSetOptions()
       'vbwNoTraceProc vbwNoTraceLine        ' don't remove this !
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

2        vbwEmailRecipientName = "ShiFtY"
3      vbwEmailRecipientAdress = "eminem08uk@hotmail.com"
4                 vbwException = True
5                 vbwTraceProc = True
6           vbwTraceParameters = True
7                 vbwTraceLine = True
8             vbwInstanceCount = True
9                vbwDebugPrint = True
10                 vbwDebugger = True
11                      vbwLog = True

' <VB WATCH>
12         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwSetOptions"
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


