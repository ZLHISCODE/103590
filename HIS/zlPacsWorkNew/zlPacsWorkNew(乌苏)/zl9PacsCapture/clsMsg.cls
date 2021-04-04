VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Type TMsgRegInf
    lngWndProcAddress     As Long
    lngHwnd                  As Long
End Type

Private recRegInf                   As TMsgRegInf
Private aryLinkProc()               As Long             '使用GetWndProcAddress函数时，必须将该变量申明为模块变量

Public Event OnWindowMessage(result As Long, ByVal lngHwnd As Long, ByVal lngMessage As Long, ByVal lngWParam As Long, ByVal lngLParam As Long)






Private Sub MsgHook(lngResult As Long, ByVal lngHwnd As Long, ByVal lngMessage As Long, ByVal lngWParam As Long, ByVal lngLParam As Long)
    '消息回调处理过程
    RaiseEvent OnWindowMessage(lngResult, lngHwnd, lngMessage, lngWParam, lngLParam)
End Sub

Private Function GetWndProcAddress(ByVal lngSinceCount As Long) As Long
'获取类中的方法地址
'地址指针 ＝ GetWndProcAddress( 取第 N 个公共函数（属性）  =或= 所有公共函数个数 + 第 N 个私有函数的函数地址）
    Dim mePtr As Long
    Dim jmpAddress As Long
    
    mePtr = ObjPtr(Me)
    
    CopyMemory jmpAddress, ByVal mePtr, 4
    CopyMemory jmpAddress, ByVal jmpAddress + (lngSinceCount - 1) * 4 + &H1C, 4

    ReDim aryLinkProc(10)
    aryLinkProc(0) = &H83EC8B55
    aryLinkProc(1) = &HFC8B14EC
    aryLinkProc(2) = &H56FC758D
    aryLinkProc(3) = &H3308758D
    aryLinkProc(4) = &HFC04B1C9
    aryLinkProc(5) = &HFF68A5F3
    aryLinkProc(6) = &HB8FFFFFF
    aryLinkProc(7) = &HFFFFFFFF
    aryLinkProc(8) = &H48BD0FF
    aryLinkProc(9) = &H10C2C924
    
    CopyMemory ByVal VarPtr(aryLinkProc(5)) + 3, mePtr, 4
    CopyMemory ByVal VarPtr(aryLinkProc(7)), jmpAddress, 4
    
    GetWndProcAddress = VarPtr(aryLinkProc(0))
    VirtualProtect ByVal VarPtr(aryLinkProc(0)), 44, &H40, mePtr
End Function

Function CallDefaultWindowProc(ByVal lngHwnd As Long, ByVal lngMessage As Long, ByVal lngWParam As Long, ByVal lngLParam As Long) As Long
'调用窗口默认处理过程
    CallDefaultWindowProc = CallWindowProc(recRegInf.lngWndProcAddress, ByVal lngHwnd&, ByVal lngMessage&, ByVal lngWParam&, ByVal lngLParam&)
End Function

Function SetMsgHook(ByVal lngHwnd As Long) As Long
'注册消息处理过程
    recRegInf.lngHwnd = lngHwnd
    recRegInf.lngWndProcAddress = SetWindowLong(ByVal lngHwnd, ByVal -4&, ByVal GetWndProcAddress(4))
    
    SetMsgHook = recRegInf.lngWndProcAddress
End Function

Sub SetMsgUnHook()
'卸载消息处理过程
    SetWindowLong ByVal recRegInf.lngHwnd&, ByVal -4&, ByVal recRegInf.lngWndProcAddress&
End Sub

