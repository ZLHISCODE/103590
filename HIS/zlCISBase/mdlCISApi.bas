Attribute VB_Name = "mdlCISApi"
Option Explicit

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
     x As Long
     y As Long
End Type

'--Grid
Public Enum flexResize
    flexResizeNone
    flexResizeColumns
    flexResizeRows
    flexResizeBoth
End Enum

Public Enum flexAlign
    flexAlignLeftTop
    flexAlignLeftCenter
    flexAlignLeftBottom
    flexAlignCenterTop
    flexAlignCenterCenter
    flexAlignCenterBottom
    flexAlignRightTop
    flexAlignRightCenter
    flexAlignRightBottom
    flexAlignGeneral
End Enum

Public Enum flexFocus
    flexFocusNone
    flexFocusLight
    flexFocusHeavy
End Enum

Public Enum flexMerge
    flexMergeNever
    flexMergeFree
    flexMergeRestrictRows
    flexMergeRestrictColumns
    flexMergeRestrictAll
End Enum

Public Const ETO_OPAQUE = 2


Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'ÅÐ¶ÏÊÇ·ñÎª±à¼­¼ü
Public Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function





