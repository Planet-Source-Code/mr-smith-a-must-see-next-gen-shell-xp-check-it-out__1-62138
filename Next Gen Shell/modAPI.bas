Attribute VB_Name = "modAPI"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global Const conHwndTopmost = -1
    Global Const conSwpNoActivate = &H10
    Global Const conSwpShowWindow = &H40
Public Sub cButtons(frm As Form)
    For Each Control In frm.Controls
        If TypeOf Control Is CommandButton Then Call SendMessage(Control.hwnd, &HF4&, &H0&, 0&)
    Next Control
End Sub

Public Function HexToDecimal(szHex As String) As Long
    Dim C As Integer, szHexVal As String, ASCII As Long
    szHexVal = "0123456789ABCDEF": ASCII = 0
    For C = 1 To Len(szHex)
        ASCII = ASCII + ((InStr(1, szHexVal, Mid(szHex, C, 1), vbTextCompare) - 1) * (16 ^ (Len(szHex) - C)))
    Next C
    HexToDecimal = ASCII
End Function
