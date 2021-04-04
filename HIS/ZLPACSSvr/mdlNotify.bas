Attribute VB_Name = "mdlNotify"
Option Explicit

Public OldWindowProc As Long
Public TheForm As Form
Public TheMenu As Menu

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Const WM_USER = &H400
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const SW_RESTORE = 9

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private TheData As NOTIFYICONDATA
' *********************************************
' The replacement window proc.
' *********************************************
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If Msg = TRAY_CALLBACK Then
        ' The user clicked on the tray icon.
        ' Look for click events.
        If lParam = WM_LBUTTONUP Then
            ' On left click, show the form.
            If TheForm.WindowState = vbMinimized Then _
                ShowWindow TheForm.hwnd, SW_RESTORE ' TheForm.WindowState = TheForm.LastState
            TheForm.SetFocus
            Exit Function
        End If
        If lParam = WM_RBUTTONUP Then
            ' On right click, show the menu.
'            TheForm.PopupMenu TheMenu
            Exit Function
        End If
    End If
    
    ' Send other messages to the original
    ' window proc.
    NewWindowProc = CallWindowProc( _
        OldWindowProc, hwnd, Msg, _
        wParam, lParam)
End Function
' *********************************************
' Add the form's icon to the tray.
' *********************************************
Public Sub AddToTray(frm As Form)
    ' ShowInTaskbar must be set to False at
    ' design time because it is read-only at
    ' run time.

    ' Save the form and menu for later use.
    Set TheForm = frm
    'Set TheMenu = mnu
    
    ' Install the new WindowProc.
'    OldWindowProc = SetWindowLong(frm.hwnd, _
'        GWL_WNDPROC, AddressOf NewWindowProc)
    
    ' Install the form's icon in the tray.
    With TheData
        .uID = 0
        .hwnd = frm.hwnd
        .cbSize = Len(TheData)
        .hIcon = frm.Icon.Handle
        .uFlags = NIF_ICON
        .uCallbackMessage = TRAY_CALLBACK
        .uFlags = .uFlags Or NIF_MESSAGE
        .cbSize = Len(TheData)
    End With
    Shell_NotifyIcon NIM_ADD, TheData
End Sub
' *********************************************
' Remove the icon from the system tray.
' *********************************************
Public Sub RemoveFromTray()
    ' Remove the icon from the tray.
    With TheData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData
    
    ' Restore the original window proc.
    SetWindowLong TheForm.hwnd, GWL_WNDPROC, _
        OldWindowProc
End Sub
' *********************************************
' Set a new tray tip.
' *********************************************
Public Sub SetTrayTip(tip As String)
    With TheData
        .szTip = tip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' *********************************************
' Set a new tray icon.
' *********************************************
Public Sub SetTrayIcon(pic As Picture)
    ' Do nothing if the picture is not an icon.
    If pic.Type <> vbPicTypeIcon Then Exit Sub

    ' Update the tray icon.
    With TheData
        .hIcon = pic.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub


