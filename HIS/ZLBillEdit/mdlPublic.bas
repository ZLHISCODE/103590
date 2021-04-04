Attribute VB_Name = "mdlPublic"
Option Explicit

''Ö§³Ö¹öÂÖÊó±êAPI**********************************************************
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC   As Long = (-4)
Private Const WM_MOUSEWHEEL As Long = &H20A
Private m_OldWindowProc As Long
Public CtlWheel As Object
''***************************************************************************

Public Sub HookWheel(ByVal frmHwnd)

    m_OldWindowProc = SetWindowLong(frmHwnd, GWL_WNDPROC, AddressOf FlexScroll)

End Sub

Public Sub UnHookWheel(ByVal hwnd As Long)

    Dim lngReturnValue As Long

    lngReturnValue = SetWindowLong(hwnd, GWL_WNDPROC, m_OldWindowProc)

End Sub

Private Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    On Error GoTo errH
    Select Case wMsg
        Case WM_MOUSEWHEEL
            If Not CtlWheel Is Nothing Then
                If TypeOf CtlWheel Is MSHFlexGrid Then
                    With CtlWheel
                        Select Case wParam
                        Case Is > 0
                           If CtlWheel.TopRow > 0 Then
                               CtlWheel.TopRow = CtlWheel.TopRow - 1
                           End If
                        Case Else
                           CtlWheel.TopRow = CtlWheel.TopRow + 1
                        End Select
                     End With
                 End If
           End If
    End Select
errH:
    FlexScroll = CallWindowProc(m_OldWindowProc, hwnd, wMsg, wParam, lParam)
End Function
