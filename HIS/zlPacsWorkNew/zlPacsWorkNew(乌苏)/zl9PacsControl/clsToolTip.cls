VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Download by http://www.codefans.net
' Create by Vincent Lavoie
' 2005
' You can you this code freely where you want for what you want
'
' This code set a balloon tool tip on every type of control with or without window handle
' (I cannot garanty that it will work with all the existing control, but it works with
' all I tested)
'
' Please vote if you like it
' Adapt the this code for your needs

Option Explicit


Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef iccInit As ICCEX) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long


Private Type PointAPI
   X As Long
   Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Type ToolInfo
    cbSize As Long
    dwFlags As Long
    hWnd As Long
    dwID As Long
    rtRect As RECT
    hInst As Long
    lpszText As Long
    lParam  As Long
End Type

Private Type ICCEX
    dwSize As Long
    dwICC As Long
End Type

Public Enum EditTipIcon
    etiNone = 0
    etiInfo = 1
    etiWarning = 2
    etiError = 3
End Enum

Private Type EDITBALLOONTIP
    cbStruct As Long
    pszTitle As Long
    pszText As Long
    ttiIcon As Long
End Type

Public Enum TOOLSTYLE
    szClassic = 1
    szBalloon = 64
End Enum


' Set Window Pos Flags
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const HS_DIAGCROSS = 5

' Init Common Controls
Private Const ICC_WIN95_CLASSES As Long = &HFF

' Misc
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETWINDOWTHEME As Long = (CCM_FIRST + &HB)
Private Const WM_USER As Long = &H400
Private Const CW_USEDEFAULT As Long = &H80000000
Private Const ECM_FIRST As Long = &H1500

' Edit Box Tip
Private Const EM_SHOWBALLOONTIP = ECM_FIRST + 3

' Window Styles
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOPMOST As Long = &H8&


' ToolTips Class
Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"

' ToolTips Flags
Private Const TTF_ABSOLUTE As Long = &H80
Private Const TTF_CENTERTIP As Long = &H2
Private Const TTF_DI_SETITEM As Long = &H8000
Private Const TTF_IDISHWND As Long = &H1
Private Const TTF_RTLREADING As Long = &H4
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_TRACK As Long = &H20
Private Const TTF_TRANSPARENT As Long = &H100

' ToolTips Icon
Private Const TTI_ERROR As Long = 3
Private Const TTI_INFO As Long = 1
Private Const TTI_NONE As Long = 0
Private Const TTI_WARNING As Long = 2

' ToolTips Message
Private Const TTM_ACTIVATE As Long = (WM_USER + 1)
Private Const TTM_ADDTOOL As Long = (WM_USER + 4)
Private Const TTM_ADJUSTRECT As Long = (WM_USER + 31)
Private Const TTM_DELTOOL As Long = (WM_USER + 5)
Private Const TTM_ENUMTOOLS As Long = (WM_USER + 14)
Private Const TTM_GETBUBBLESIZE As Long = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOL As Long = (WM_USER + 15)
Private Const TTM_GETDELAYTIME As Long = (WM_USER + 21)
Private Const TTM_GETMARGIN As Long = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH As Long = (WM_USER + 25)
Private Const TTM_GETTEXT As Long = (WM_USER + 11)
Private Const TTM_GETTIPBKCOLOR As Long = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR As Long = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
Private Const TTM_GETTOOLINFO As Long = (WM_USER + 8)
Private Const TTM_HITTEST As Long = (WM_USER + 10)
Private Const TTM_NEWTOOLRECT As Long = (WM_USER + 6)
Private Const TTM_POP As Long = (WM_USER + 28)
Private Const TTM_POPUP As Long = (WM_USER + 34)
Private Const TTM_RELAYEVENT As Long = (WM_USER + 7)
Private Const TTM_SETDELAYTIME As Long = (WM_USER + 3)
Private Const TTM_SETMARGIN As Long = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Private Const TTM_SETTITLE As Long = (WM_USER + 32)
Private Const TTM_SETTOOLINFO As Long = (WM_USER + 9)
Private Const TTM_SETWINDOWTHEME As Long = CCM_SETWINDOWTHEME
Private Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Private Const TTM_TRACKPOSITION As Long = (WM_USER + 18)
Private Const TTM_UPDATE As Long = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXT As Long = (WM_USER + 12)
Private Const TTM_WINDOWFROMPOINT As Long = (WM_USER + 16)

' ToolTips Notification
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFO As Long = (TTN_FIRST - 0)
Private Const TTN_LAST As Long = (-549)
Private Const TTN_LINKCLICK As Long = (TTN_FIRST - 3)
Private Const TTN_NEEDTEXT As Long = TTN_GETDISPINFO
Private Const TTN_POP As Long = (TTN_FIRST - 2)
Private Const TTN_SHOW As Long = (TTN_FIRST - 1)

' ToolTips Creation Flags
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_BALLOON As Long = &H40
Private Const TTS_NOANIMATE As Long = &H10
Private Const TTS_NOFADE As Long = &H20
Private Const TTS_NOPREFIX As Long = &H2

Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

Private m_hwndTip As Long
Private m_hwndObject As Long
Private m_blnCentered As Boolean
Private m_strText As String
Private m_strTitle As String
Private m_strForeColor As String
Private m_strBackColor As String
Private m_varEditTipIcon As EditTipIcon
Private m_intDelay As Integer
Private m_intKillAfter As Integer


'Public Sub CreateBalloon(Object As Object, hwndObject As Long, szText As String, _
'    Style As TOOLSTYLE, szCentered As Boolean, Optional szTitle As String, _
'    Optional mvarIcon As EditTipIcon, Optional BackColor As String, _
'    Optional ForeColor As String)

Public Sub CreateBalloon(hwndObject As Long, szText As String, _
    Style As TOOLSTYLE, szCentered As Boolean, Optional szTitle As String, _
    Optional mvarIcon As EditTipIcon, Optional BackColor As String, _
    Optional ForeColor As String)
    
    m_blnCentered = szCentered
    m_strText = szText
    m_strTitle = szTitle
    m_strBackColor = BackColor
    m_strForeColor = ForeColor
    m_varEditTipIcon = mvarIcon
    m_hwndObject = hwndObject
    

    
    CreateWndTips hwndObject, Style
    
    Dim tiInfo As ToolInfo
    
    If hwndObject <> 0 Then
''        SetToolTip tiInfo, Object
        SetToolTip tiInfo, Nothing
    End If
    
End Sub

Public Sub SetHandle(Object As Object)
    If m_hwndObject <> 0 Then Exit Sub
    Debug.Print "SetHandle Handle:" & m_hwndTip

    
    Dim myCurrCurPos As PointAPI
    GetCursorPos myCurrCurPos
    m_hwndObject = WindowFromPoint(myCurrCurPos.X, myCurrCurPos.Y)
    
    Dim tiInfo As ToolInfo
    
    SetToolTip tiInfo, Object

End Sub

Private Sub CreateWndTips(hWndParent As Long, Style As Long)
    
    Dim dwFlags As Long
    Dim ICEx As ICCEX
    
    dwFlags = TTS_NOPREFIX Or TTS_ALWAYSTIP Or Style
    
    With ICEx
        .dwSize = Len(ICEx)
        .dwICC = ICC_WIN95_CLASSES
    End With
    
    InitCommonControlsEx ICEx
    
    m_hwndTip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASSA, "", WS_POPUP Or dwFlags, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, hWndParent, 0, App.hInstance, ByVal 0&)
    
    If m_hwndTip = 0 Then Exit Sub
    
    SetWindowPos m_hwndTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
End Sub



Private Sub SetToolTip(ByRef tiInfo As ToolInfo, Object As Object)
   
    With tiInfo
        If m_blnCentered = True Then
            .dwFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_TRANSPARENT
        Else
            .dwFlags = TTF_SUBCLASS Or TTF_TRANSPARENT
        End If
        
        .hWnd = m_hwndObject
        .lpszText = StrPtr(StrConv(m_strText, vbFromUnicode))
        .hInst = App.hInstance
        
'        If m_hwndObject = Object.Parent.hWnd Then
'            .rtRect.Left = Object.Left / Screen.TwipsPerPixelY
'            .rtRect.Top = Object.Top / Screen.TwipsPerPixelX
'            .rtRect.Bottom = .rtRect.Top + (Object.Height / Screen.TwipsPerPixelY)
'            .rtRect.Right = .rtRect.Left + (Object.Width / Screen.TwipsPerPixelX)
'        Else
            GetClientRect m_hwndObject, .rtRect
'        End If
        
        .cbSize = Len(tiInfo)
    End With
    SendMessage m_hwndTip, TTM_ADDTOOL, 0&, tiInfo

    If m_strTitle <> "" Then
        SendMessage m_hwndTip, TTM_SETTITLE, CLng(m_varEditTipIcon), ByVal m_strTitle
    End If
    
    If m_strBackColor <> "" Then
        SendMessage m_hwndTip, TTM_SETTIPBKCOLOR, m_strBackColor, 0&
    End If
    
    If m_strForeColor <> "" Then
        SendMessage m_hwndTip, TTM_SETTIPTEXTCOLOR, m_strForeColor, 0&
    End If
    
    If m_intKillAfter <> -1 Then
        SendMessageLong m_hwndTip, TTM_SETDELAYTIME, TTDT_AUTOPOP, m_intKillAfter
    End If
    If m_intDelay <> -1 Then
       SendMessageLong m_hwndTip, TTM_SETDELAYTIME, TTDT_INITIAL, m_intDelay
    End If

End Sub

Private Sub Class_Initialize()
    m_intDelay = -1
    m_intKillAfter = -1
End Sub

Private Sub Class_Terminate()
' if you remove this comments, you will have to declare all you class globally
' anyway, windows attach the tooltip window to the control, when the control is deleted,
' the tooltip window is deleted too
'    If m_hwndTip <> 0 Then
'        DestroyWindow m_hwndTip
'    End If
End Sub

Public Property Get VisibleTime() As Integer
   VisibleTime = m_intKillAfter
End Property

Public Property Let VisibleTime(ByVal lData As Integer)
   m_intKillAfter = lData
End Property

Public Property Get DelayTime() As Integer
   DelayTime = m_intDelay
End Property

Public Property Let DelayTime(ByVal lData As Integer)
   m_intDelay = lData
End Property


