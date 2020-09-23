Attribute VB_Name = "blah"
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000


Public lngTPPY As Long
Public lngTPPX As Long
Function Trans(tForm As Form)
On Error Resume Next
Dim Ret As Long
Dim TC As Long
TC = &HFF0000 'This is vbBlue (TC = The Colour As vbBlue)
Ret = GetWindowLong(tForm.hwnd, G_E)
Ret = Ret Or W_E
SetWindowLong tForm.hwnd, G_E, Ret
SetLayeredWindowAttributes tForm.hwnd, TC, 0, LW_KEY
End Function

Sub MouseDown()
    Dim POINT As POINTAPI
    
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
End Sub
Sub MouseMove(ctlForm As Form)
    Dim lngX     As Long
    Dim lngY     As Long
    Dim POINT    As POINTAPI
       
    GetCursorPos POINT
    lngX& = (POINT.X - LastPoint.X) * lngTPPX&
    lngY& = (POINT.Y - LastPoint.Y) * lngTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    ctlForm.Move ctlForm.Left + lngX&, ctlForm.Top + lngY&
End Sub
Sub InitTPP()
    lngTPPX& = Screen.TwipsPerPixelX
    lngTPPY& = Screen.TwipsPerPixelY
End Sub
Public Sub CreateKey(Folder As String, Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value

End Sub

Public Sub CreateIntegerKey(Folder As String, Value As Integer)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value, "REG_DWORD"


End Sub

Public Function ReadKey(Value As String) As String

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
r = b.RegRead(Value)
ReadKey = r
End Function


Public Sub DeleteKey(Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("Wscript.Shell")
b.RegDelete Value
End Sub
