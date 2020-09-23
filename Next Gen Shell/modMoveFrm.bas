Attribute VB_Name = "modMoveFrm"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = 161

Public Function DragForm(F As Form)
ReleaseCapture
SendMessage F.hwnd, WM_NCLBUTTONDOWN, 2, 0&
End Function

Public Function DragObj(O As Object)
ReleaseCapture
SendMessage O.hwnd, WM_NCLBUTTONDOWN, 2, 0&
End Function
