Attribute VB_Name = "mMouseHook"
Option Explicit

' ======================================================================================
' GDI�����͸�������
' ======================================================================================

'��
Private Type POINTAPI
   X As Long
   Y As Long
End Type

'��깳�ӽṹ��
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

'��깳�Ӿ��
Private m_hMouseHook As Long

'δ����ָ�����顣
Private m_lMouseHookPtr() As Long

'���������顣
Private m_lMouseHookhWnd() As Long

'������յ�����깳��֪ͨ����Ŀ
Private m_iMouseHookCount As Long

'#########################################################################################################
'## ���ܣ�  Ϊָ������������깳�ӡ�
'## ������  ctlGrid:   ��Ҫ������깳�ӵ�����
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
          '�Ѿ������˹���
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
'## ���ܣ�  Ϊָ������ȡ����깳�ӡ�
'## ������  ctlGrid:   ��Ҫȡ����깳�ӵ�����
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
       '�������Ѿ�û�й�����
    End If
    
    If m_iMouseHookCount <= 0 Then
        If (m_hMouseHook <> 0) Then
            UnhookWindowsHookEx m_hMouseHook
            m_hMouseHook = 0
        End If
    End If
End Sub

'#########################################################################################################
'## ���ܣ�  ���ڷ���ָ��������AddressOf�ĵ�ַ���ڱ�������������ΪAddressOf��һ��һԪ�����������ֱ��ʹ�ã�
'## ������  lPtr: ���ڻ�ȡAddressOf�ı���
'## ���أ�  AddressOf���ص�ָ��
'#########################################################################################################
Private Function HookAddress(ByVal lPtr As Long) As Long
   HookAddress = lPtr
End Function

'#########################################################################################################
'## ���ܣ�  ��깳�ӵĻص�����
'## ������  nCode:  ���Ӵ���ֵ
'##         wParam: �����Ϣ����
'##         lParam: һ��ָ�� MOUSEHOOKSTRUCT �ṹ���ָ�룬�����������Ϣ��
'## ���أ�  ��һ����깳�ӵ�ֵ������еĻ���
'#########################################################################################################
Private Function MouseFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
    Dim tMHS As MOUSEHOOKSTRUCT
    Dim i As Long
    Dim ctlGrid As Table
    
    On Error GoTo ErrorHandler
    
    ' ���� lParam:
    CopyMemory tMHS, ByVal LParam, Len(tMHS)
    
    ' ��ѯ�󶨵�����ֻ��һ����
    For i = 1 To m_iMouseHookCount
        '���ø����������¼���
        If Not (m_lMouseHookPtr(i) = 0) Then
            If Not (IsWindow(m_lMouseHookhWnd(i)) = 0) Then
                Set ctlGrid = ObjectFromPtr(m_lMouseHookPtr(i))
                If Not ctlGrid Is Nothing Then
                    '����ָ�����������¼�
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



