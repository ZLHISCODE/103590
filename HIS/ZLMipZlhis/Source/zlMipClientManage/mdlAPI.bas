Attribute VB_Name = "mdlAPI"
Option Explicit


Public Const GWL_STYLE = (-16)              'Set the window style
Public Const ETO_CLIPPED = 4
Public Const ETO_GRAYED = 1
Public Const ETO_OPAQUE = 2
Public Const CB_GETDROPPEDSTATE = &H157

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Sub InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
