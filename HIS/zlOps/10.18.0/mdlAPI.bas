Attribute VB_Name = "mdlAPI"
Option Explicit

'######################################################################################################################

Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Const VK_CAPITAL = &H14

Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_NONOTIFY = &H80&
Public Const TPM_RETURNCMD = &H100&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_CHECKED = &H8&
Public Const MF_BYCOMMAND = &H0&
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息

Public Const KLF_REORDER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)

'自定义类型
'----------------------------------------------------------------------------------------------------------------------
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Public Type POINT
    X As Long
    Y As Long
End Type

'过程和函数申明
'----------------------------------------------------------------------------------------------------------------------
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal sCaption As String) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, nIgnored As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub GetKeyboardStateByString Lib "user32" Alias "GetKeyboardState" (ByVal pbKeyState As String)
Public Declare Sub SetKeyboardStateByString Lib "user32" Alias "SetKeyboardState" (ByVal lppbKeyState As String)

'######################################################################################################################

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    '******************************************************************************************************************
    '功能：如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    '参数：
    '返回：
    '******************************************************************************************************************
    
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "", Optional ByVal Align As Byte, Optional ByVal Off As Single = 0)
    '******************************************************************************************************************
    '功能：将PictureBox模拟成3D平面按钮
    '参数：intStyle:0=平面,-1=凹下,1=凸起
    '返回：
    '******************************************************************************************************************
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            If Align = 1 Then
                .CurrentX = 45
            Else
                .CurrentX = Off + (.ScaleWidth - .TextWidth(strName)) / 2
            End If
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal X As Single, ByVal Y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = X / 15
    lngY = Y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub



