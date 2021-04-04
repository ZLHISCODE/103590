Attribute VB_Name = "mdlMsgBox"
Option Explicit


Private hHook As Long
Private hFormhWnd As Long
 
 
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'替代VB中的Msgbox函数
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgboxEx(hWnd As Long, sPrompt As String, Optional dwStyle As Long, Optional sTitle As String) As Long

    Dim hInstance As Long
    Dim hThreadId As Long

    hInstance = App.hInstance
    hThreadId = App.ThreadID

    If dwStyle = 0 Then dwStyle = vbOKOnly
    If Len(sTitle) = 0 Then sTitle = App.EXEName

    '将当前窗口的句柄付给变量
    hFormhWnd = hWnd

    '设置钩子
    hHook = SetWindowsHookEx(WH_CBT, AddressOf CBTProc, hInstance, hThreadId)
    
    '调用MessageBox API
    MsgboxEx = MessageBox(hWnd, sPrompt, sTitle, dwStyle)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'HOOK处理
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CBTProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    '变量声明
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

    '当MessageBox出现时，将Msgbox对话框居中与所在的窗口
    If nCode = HCBT_ACTIVATE Then
    
        '消息为HCBT_ACTIVATE时，参数wParam包含的是MessageBox的句柄
        hwndMsgBox = wParam
        
        '得到MessageBox对话框的Rect
        Call GetWindowRect(hwndMsgBox, rc)
        Call GetWindowRect(hFormhWnd, rcFrm)
        
        '使MessageBox居中
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
        
'        '修改确定按钮的文字
'        Call SetDlgItemText(hwndMsgBox, IDOK, "这是确定按钮")
        SetWindowPos hwndMsgBox, -1, rcFrm.Left, rcFrm.Top, dlgWidth, dlgHeight, 3 '将窗口置顶
        SetForegroundWindow hwndMsgBox
        
        'Msgbox居中
        Call MoveWindow(hwndMsgBox, newLeft, newTop, dlgWidth, dlgHeight, True)
        
        '卸载钩子
        UnhookWindowsHookEx hHook
    End If
    
    CBTProc = False
End Function


