Attribute VB_Name = "WinProc"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Public Const WM_NCDESTROY = &H82

Private Type HOOKINFO
    hWnd As Long        'Subclassed window
    Ctrl As Subclass    'Control
    OldWndProc As Long  'Old window procedure
End Type

'Note: These variables will be common to all
'control instances within an application
Private HookArray() As HOOKINFO
Private NumHooks As Integer

'Hooks the specified window/control
Public Sub HookWindow(hWnd As Long, Ctrl As Subclass)
    Dim i As Integer
    If hWnd <> 0 Then
        'Note: Since we use the window handle to identify
        'the subclassing control, we cannot allow more than
        'one control to subclass the same window. So before
        'hooking a window, we remove any existing hooks to
        'that same window.
        UnhookWindow hWnd
        'Add new hook for this window
        NumHooks = NumHooks + 1
        ReDim Preserve HookArray(NumHooks)
        HookArray(NumHooks).hWnd = hWnd
        Set HookArray(NumHooks).Ctrl = Ctrl
        HookArray(NumHooks).OldWndProc = GetWindowLong(hWnd, GWL_WNDPROC)
        'Install custom window procedure for this window
        SetWindowLong hWnd, GWL_WNDPROC, AddressOf WndProc
    End If
End Sub

'Unhook the specified window
'Set nStartIndex to index of window (if known)
Public Sub UnhookWindow(hWnd As Long)
    Dim i As Integer, j As Integer
    'Reset window hook for this window
    For i = 1 To NumHooks
        If HookArray(i).hWnd = hWnd Then
            'Sanity check
            Debug.Assert HookArray(i).OldWndProc <> 0
            'Reset previous window procedure
            SetWindowLong hWnd, GWL_WNDPROC, HookArray(i).OldWndProc
            'Remove hook information from array
            NumHooks = NumHooks - 1
            For j = i To NumHooks
                HookArray(j) = HookArray(j + 1)
            Next j
            ReDim Preserve HookArray(NumHooks)
            Exit For
        End If
    Next i
End Sub

'Call the original window procedure
Public Function CallWndProc(hWnd As Long, Msg As Long, wParam As Long, lParam As Long) As Long
    Dim i As Integer
    'Find hook information for this window
    For i = 1 To NumHooks
        If HookArray(i).hWnd = hWnd Then
            CallWndProc = CallWindowProc(HookArray(i).OldWndProc, hWnd, Msg, wParam, lParam)
            Exit For
        End If
    Next i
End Function

'Replacement window procedure--Invokes control handler
Private Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim i As Integer
    'Find hook information for this window
    For i = 1 To NumHooks
        If HookArray(i).hWnd = hWnd Then
            'Sanity check
            Debug.Assert HookArray(i).Ctrl.hWnd = hWnd
            'Does control want this message?
            If HookArray(i).Ctrl.Messages(Msg) Then
                'Suppress unhandled run-time errors
                On Error Resume Next
                'Send message to control
                WndProc = HookArray(i).Ctrl.RaiseWndProc(Msg, wParam, lParam)
            Else
                'Otherwise, just call default window handler
                WndProc = CallWindowProc(HookArray(i).OldWndProc, hWnd, Msg, wParam, lParam)
            End If
            'Unhook this window if it is being destroyed
            If Msg = WM_NCDESTROY Then
                HookArray(i).Ctrl.hWnd = 0
            End If
            Exit For
        End If
    Next i
End Function

