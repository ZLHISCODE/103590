VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTabInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event TabKeyDown()
Public Event TabKeyUp()

Public Event sTabKeyDown()
Public Event sTabKeyUp()

Private Sub Class_Initialize()
    If gobjTab Is Nothing And glngKeyHook = 0 Then
        glngKeyHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf CustomHook, App.hInstance, 0)
        Set gobjTab = Me
    End If
End Sub

Private Sub Class_Terminate()
    If glngKeyHook <> 0 Then
        UnhookWindowsHookEx glngKeyHook
        glngKeyHook = 0
        Set gobjTab = Nothing
    End If
End Sub

Friend Sub ACT_TabKeyDown()
    RaiseEvent TabKeyDown
End Sub

Friend Sub ACT_TabKeyUp()
    RaiseEvent TabKeyUp
End Sub

Friend Sub ACT_sTabKeyDown()
    RaiseEvent sTabKeyDown
End Sub

Friend Sub ACT_sTabKeyUp()
    RaiseEvent sTabKeyUp
End Sub

