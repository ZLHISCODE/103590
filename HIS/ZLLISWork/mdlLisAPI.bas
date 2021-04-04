Attribute VB_Name = "mdlLisAPI"
Option Explicit

Private Const MAX_PATH = 260

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
'--Grid
Public Type CellAttribute
    Disabled As Boolean
    EditMode As EditType
    ItemIndex As Long    '单元列表指针
    ListIndex As Long '单元列表当前项目索引
End Type
Public Type CellItem '单元列表
    List() As String
End Type
Public Enum EditType
    editTextBox = 0
    editComboBox = 1
    editDate = 2
End Enum

Public Enum flexMode
    flexNone
    flexEdit
    flexAdd
End Enum

Public Enum flexResize
    flexResizeNone
    flexResizeColumns
    flexResizeRows
    flexResizeBoth
End Enum

Public Enum flexAlign
    flexAlignLeftTop
    flexAlignLeftCenter
    flexAlignLeftBottom
    flexAlignCenterTop
    flexAlignCenterCenter
    flexAlignCenterBottom
    flexAlignRightTop
    flexAlignRightCenter
    flexAlignRightBottom
    flexAlignGeneral
End Enum

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Enum flexText
    flexTextFlat
    flexTextRaised
    flexTextInset
    flexTextRaisedLight
    flexTextInsetLight
End Enum

Public Enum flexFocus
    flexFocusNone
    flexFocusLight
    flexFocusHeavy
End Enum

Public Enum flexGridLine
    flexGridNone
    flexGridFlat
    flexGridInset
    flexGridRaised
    flexGridDashes
    flexGridDots
End Enum

Public Enum flexHighLight
    flexHighlightNever
    flexHighlightAlways
    flexHighlightWithFocus
End Enum

Public Enum flexMerge
    flexMergeNever
    flexMergeFree
    flexMergeRestrictRows
    flexMergeRestrictColumns
    flexMergeRestrictAll
End Enum

Public Enum flexRowSize
    flexRowSizeIndividual
    flexRowSizeAll
End Enum

Public Const ETO_OPAQUE = 2
'--
Public Const PS_DOT = 2
Public Const PS_DASH = 1
Public Const PS_SOLID = 0
Public Const PS_INSIDEFRAME = 6
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const R2_XORPEN = 7

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Public Const BF_SOFT = &H1000

Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
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

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Public Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Public Declare Function GetROP2 Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal sCaption As String) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, nIgnored As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Const INVALID_HANDLE_VALUE = -1
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Const Process_Query_Information = &H400
Public Const Still_Active = &H103
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long






'判断是否为编辑键
Public Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function


Public Sub PicShowFlat(objPic As Object, Optional intStyle As Integer = -1, Optional strName As String = "", Optional intAlign As mTextAlign)
'功能：将PictureBox模拟成凹下或凸起的面板
'参数：intStyle:0=平面,-1=凹下,1=凸起
'      intAlign=如果要显示文本,则指定对齐方式
    
    Dim vRect As RECT, lngTmp As Long
    
    With objPic
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            vRect.Left = .ScaleLeft
            vRect.Top = .ScaleTop
            vRect.Right = .ScaleWidth
            vRect.Bottom = .ScaleHeight
            DrawEdge .hDC, vRect, CLng(IIf(intStyle = 1, EDGE_RAISED Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If intAlign = taCenterAlign Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2 '中间对齐
            ElseIf intAlign = taRightAlign Then
                .CurrentX = .ScaleWidth - .TextWidth(strName) - 45 '右边对齐
            Else
                .CurrentX = 45 '左边对齐
            End If
            objPic.Print strName
        End If
    End With
End Sub

Public Function GetTrayHeight() As Long
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取任务栏的高度
    '------------------------------------------------------------------------------------------------------------------
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

Public Sub WriteLog(ByVal strModule As String, ByVal strFunction As String, ByVal strErr As String)
    '调用公共方法记录日志
    Call zl9Comlib.LogWrite("LIS老版通讯程序错误调试日志", strModule, strFunction, strErr)
End Sub

