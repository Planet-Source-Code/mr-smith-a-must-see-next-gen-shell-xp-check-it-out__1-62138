Attribute VB_Name = "modMouse"

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function GetX()
Dim p As POINTAPI
GetCursorPos p
GetX = p.X
End Function

Public Function GetY()
Dim p As POINTAPI
GetCursorPos p
GetY = p.Y
End Function
