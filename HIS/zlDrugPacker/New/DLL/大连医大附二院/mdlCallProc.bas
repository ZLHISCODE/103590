Attribute VB_Name = "mdlCallProc"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public glngResult As Long
Public glngOldWinProc As Long

Public Const WM_COMMAND = &H111
Public Const MF_STRING = &H0&
Public Const WM_USER = &H400
Public Const GWL_WNDPROC = -4

Public Function OnMenu(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case wMsg
        Case WM_COMMAND
            Select Case wParam
                Case 1 + WM_USER
                    Call SetInterface
                Case 2 + WM_USER
                    'MsgBox "u select save", vbInformation, "hello, world!"
                Case 3 + WM_USER
                    'MsgBox "u select save as", vbInformation, "hello, world!"
            End Select
    End Select
    OnMenu = CallWindowProc(glngOldWinProc, hwnd, wMsg, wParam, lParam)
End Function

Public Sub MyProc()
    glngOldWinProc = GetWindowLong(gfrmOwner.hwnd, GWL_WNDPROC)
    SetWindowLong gfrmOwner.hwnd, GWL_WNDPROC, AddressOf OnMenu
End Sub

Public Sub UnMyProc()
    If glngOldWinProc <> 0 Then
        SetWindowLong gfrmOwner.hwnd, GWL_WNDPROC, glngOldWinProc
    End If
End Sub

