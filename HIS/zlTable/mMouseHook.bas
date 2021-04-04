Attribute VB_Name = "mMouseHook"
Option Explicit

' ======================================================================================
' GDI声明和辅助函数
' ======================================================================================

'点
Private Type POINTAPI
   X As Long
   Y As Long
End Type

'鼠标钩子结构体
Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const WH_MOUSE = 7

Private Const WM_RBUTTONUP As Long = &H205

'鼠标钩子句柄
Private m_hMouseHook As Long

'未引用指针数组。
Private m_lMouseHookPtr() As Long

'窗体句柄数组。
Private m_lMouseHookhWnd() As Long

'网格接收到的鼠标钩子通知的数目
Private m_iMouseHookCount As Long

'#########################################################################################################
'## 功能：  为指定网格设置鼠标钩子。
'## 参数：  ctlGrid:   需要设置鼠标钩子的网格
'#########################################################################################################
Public Sub AttachMouseHook(ctlGrid As Table)
    Dim lpfn As Long
    Dim lPtr As Long
    Dim i As Long
   
    If m_iMouseHookCount = 0 Then
       lpfn = HookAddress(AddressOf MouseFilter)
       m_hMouseHook = SetWindowsHookEx(WH_MOUSE, lpfn, 0&, GetCurrentThreadId())
       Debug.Assert (m_hMouseHook <> 0)
    End If
    lPtr = ObjPtr(ctlGrid)
    For i = 1 To m_iMouseHookCount
       If lPtr = m_lMouseHookPtr(i) Then
          '已经设置了钩子
          Debug.Assert False
          Exit Sub
       End If
    Next i
    ReDim Preserve m_lMouseHookPtr(1 To m_iMouseHookCount + 1) As Long
    ReDim Preserve m_lMouseHookhWnd(1 To m_iMouseHookCount + 1) As Long
    m_iMouseHookCount = m_iMouseHookCount + 1
    m_lMouseHookPtr(m_iMouseHookCount) = lPtr
    m_lMouseHookhWnd(m_iMouseHookCount) = ctlGrid.hWnd
End Sub

'#########################################################################################################
'## 功能：  为指定网格取消鼠标钩子。
'## 参数：  ctlGrid:   需要取消鼠标钩子的网格
'#########################################################################################################
Public Sub DetachMouseHook(ctlGrid As Table)
    Dim i As Long
    Dim lPtr As Long
    Dim iThis As Long
   
    lPtr = ObjPtr(ctlGrid)
    For i = 1 To m_iMouseHookCount
        If m_lMouseHookPtr(i) = lPtr Then
            iThis = i
            Exit For
        End If
    Next i
    If iThis <> 0 Then
        If m_iMouseHookCount > 1 Then
            For i = iThis To m_iMouseHookCount - 1
                m_lMouseHookPtr(i) = m_lMouseHookPtr(i + 1)
            Next i
        End If
        m_iMouseHookCount = m_iMouseHookCount - 1
        If m_iMouseHookCount >= 1 Then
            ReDim Preserve m_lMouseHookPtr(1 To m_iMouseHookCount) As Long
        Else
            Erase m_lMouseHookPtr
        End If
    Else
       '该网格已经没有钩子了
    End If
    
    If m_iMouseHookCount <= 0 Then
        If (m_hMouseHook <> 0) Then
            UnhookWindowsHookEx m_hMouseHook
            m_hMouseHook = 0
        End If
    End If
End Sub

'#########################################################################################################
'## 功能：  用于返回指定变量的AddressOf的地址用于保存至变量（因为AddressOf是一个一元运算符，不能直接使用）
'## 参数：  lPtr: 用于获取AddressOf的变量
'## 返回：  AddressOf返回的指针
'#########################################################################################################
Private Function HookAddress(ByVal lPtr As Long) As Long
   HookAddress = lPtr
End Function

'#########################################################################################################
'## 功能：  鼠标钩子的回调函数
'## 参数：  nCode:  钩子代码值
'##         wParam: 鼠标消息代码
'##         lParam: 一个指向 MOUSEHOOKSTRUCT 结构体的指针，包含了鼠标消息。
'## 返回：  下一个鼠标钩子的值，如果有的话。
'#########################################################################################################
Private Function MouseFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
    Dim tMHS As MOUSEHOOKSTRUCT
    Dim i As Long
    Dim ctlGrid As Table
    
    On Error GoTo ErrorHandler
    
    ' 解码 lParam:
    CopyMemory tMHS, ByVal LParam, Len(tMHS)
    
    ' 查询绑定的网格（只有一个）
    For i = 1 To m_iMouseHookCount
        '调用该网格的鼠标事件。
        If Not (m_lMouseHookPtr(i) = 0) Then
            If Not (IsWindow(m_lMouseHookhWnd(i)) = 0) Then
                Set ctlGrid = ObjectFromPtr(m_lMouseHookPtr(i))
                If Not ctlGrid Is Nothing Then
                    '调用指定网格的鼠标事件
                    If ctlGrid.MouseEvent(wParam, tMHS.hWnd, tMHS.pt.X, tMHS.pt.Y, tMHS.wHitTestCode) Then
                       
                    End If
                End If
            End If
        End If
    Next i
    
    If Not (m_hMouseHook = 0) Then
        MouseFilter = CallNextHookEx(m_hMouseHook, nCode, wParam, LParam)
    End If
    Exit Function
ErrorHandler:
    Exit Function
End Function



