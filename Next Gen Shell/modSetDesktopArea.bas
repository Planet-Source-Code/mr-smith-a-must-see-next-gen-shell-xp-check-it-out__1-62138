Attribute VB_Name = "modSetDesktopArea"

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_GETWORKAREA = 48
Public Const SPI_SETWORKAREA = 47
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_SENDCHANGE = SPIF_SENDWININICHANGE

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum E_RESIZEFROM
    RF_FROMCURRENT = &H4
    RF_FROMFULL = &H5
End Enum

Public Function SetDesktopArea(ByVal RF_FROM As E_RESIZEFROM, Optional lTop As Long = 0, Optional lRight As Long = 0, Optional lLeft As Long = 0, Optional lBottom As Long = 0) As Boolean
    
    Dim rctScreen As RECT, intScreenHeight As Integer, intScreenWidth As Integer, lResult As Long
    
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, rctScreen, SPIF_SENDCHANGE)
    
    Select Case RF_FROM
        Case RF_FROMCURRENT
            
            rctScreen.Top = rctScreen.Top + lTop
            rctScreen.Bottom = rctScreen.Bottom - lBottom
            rctScreen.Left = rctScreen.Left + lLeft
            rctScreen.Right = rctScreen.Right - lRight
            
            
            lResult = SystemParametersInfo(SPI_SETWORKAREA, 0, rctScreen, SPIF_SENDCHANGE)
            
            
            If lResult = 0 Then SetDesktopArea = False Else SetDesktopArea = True
            
            
            Exit Function
        Case RF_FROMFULL
            
            intScreenHeight = Screen.Height / Screen.TwipsPerPixelY
            intScreenWidth = Screen.Width / Screen.TwipsPerPixelX
            
            
            rctScreen.Top = 0 + lTop
            rctScreen.Bottom = intScreenHeight - lBottom
            rctScreen.Left = 0 + lLeft
            rctScreen.Right = intScreenWidth - lRight
            
            
            lResult = SystemParametersInfo(SPI_SETWORKAREA, 0, rctScreen, SPIF_SENDCHANGE)
            
            
            If lResult = 0 Then SetDesktopArea = False Else SetDesktopArea = True
            
            
            Exit Function
        Case Else
            
            SetDesktopArea = False
            Exit Function
        End Select
    
End Function
    
