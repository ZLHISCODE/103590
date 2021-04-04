Attribute VB_Name = "mdlMsgBox"
Option Explicit


Private hHook As Long
Private hFormhWnd As Long
 
 
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'���VB�е�Msgbox����
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgboxEx(hWnd As Long, sPrompt As String, Optional dwStyle As Long, Optional sTitle As String) As Long

    Dim hInstance As Long
    Dim hThreadId As Long

    hInstance = App.hInstance
    hThreadId = App.ThreadID

    If dwStyle = 0 Then dwStyle = vbOKOnly
    If Len(sTitle) = 0 Then sTitle = App.EXEName

    '����ǰ���ڵľ����������
    hFormhWnd = hWnd

    '���ù���
    hHook = SetWindowsHookEx(WH_CBT, AddressOf CBTProc, hInstance, hThreadId)
    
    '����MessageBox API
    MsgboxEx = MessageBox(hWnd, sPrompt, sTitle, dwStyle)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'HOOK����
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CBTProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    '��������
    Dim rc As RECT
    Dim rcFrm As RECT

    Dim newLeft As Long
    Dim newTop As Long
    Dim dlgWidth As Long
    Dim dlgHeight As Long
    Dim scrWidth As Long
    Dim scrHeight As Long
    Dim frmLeft As Long
    Dim frmTop As Long
    Dim frmWidth As Long
    Dim frmHeight As Long
    Dim hwndMsgBox As Long

    '��MessageBox����ʱ����Msgbox�Ի�����������ڵĴ���
    If nCode = HCBT_ACTIVATE Then
    
        '��ϢΪHCBT_ACTIVATEʱ������wParam��������MessageBox�ľ��
        hwndMsgBox = wParam
        
        '�õ�MessageBox�Ի����Rect
        Call GetWindowRect(hwndMsgBox, rc)
        Call GetWindowRect(hFormhWnd, rcFrm)
        
        'ʹMessageBox����
        frmLeft = rcFrm.Left
        frmTop = rcFrm.Top
        frmWidth = rcFrm.Right - rcFrm.Left
        frmHeight = rcFrm.Bottom - rcFrm.Top
        dlgWidth = rc.Right - rc.Left
        dlgHeight = rc.Bottom - rc.Top
    
        scrWidth = Screen.Width \ Screen.TwipsPerPixelX
        scrHeight = Screen.Height \ Screen.TwipsPerPixelY
    
        newLeft = frmLeft + ((frmWidth - dlgWidth) \ 2)
        newTop = frmTop + ((frmHeight - dlgHeight) \ 2)
        
'        '�޸�ȷ����ť������
'        Call SetDlgItemText(hwndMsgBox, IDOK, "����ȷ����ť")
        SetWindowPos hwndMsgBox, -1, rcFrm.Left, rcFrm.Top, dlgWidth, dlgHeight, 3 '�������ö�
        SetForegroundWindow hwndMsgBox
        
        'Msgbox����
        Call MoveWindow(hwndMsgBox, newLeft, newTop, dlgWidth, dlgHeight, True)
        
        'ж�ع���
        UnhookWindowsHookEx hHook
    End If
    
    CBTProc = False
End Function


