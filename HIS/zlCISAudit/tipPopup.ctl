VERSION 5.00
Begin VB.UserControl tipPopup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "tipPopup.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   1575
      Top             =   540
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   900
      Top             =   540
   End
End
Attribute VB_Name = "tipPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAGIC_END_EDIT_IGNORE_WINDOW_PROP As String = "VBAL:SGRID:EDITOR"

Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" ( _
        ByVal hIml As Long, ByVal i As Long, _
        ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal fStyle As Long _
    ) As Long
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_OVERLAYMASK = 3840
Private Const ILC_COLOR = &H0
Private Const ILC_COLOR32 = &H20
Private Const ILC_MASK = &H1&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadIconString Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Const IMAGE_ICON = 1
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000&
Private Const LR_CREATEDIBSECTION = &H2000&
Private Const LR_COPYFROMRESOURCE = &H4000&
Private Const LR_SHARED = &H8000&

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
   Private Const GWL_STYLE = (-16)
   Private Const WS_BORDER = &H800000
   Private Const WS_CHILD = &H40000000
   Private Const WS_DISABLED = &H8000000
   Private Const WS_VISIBLE = &H10000000
   Private Const WS_TABSTOP = &H100000
   Private Const WS_HSCROLL = &H100000
   Private Const GWL_EXSTYLE = (-20)
   Private Const WS_EX_TOPMOST = &H8&
   Private Const WS_EX_CLIENTEDGE = &H200&
   Private Const WS_EX_STATICEDGE = &H20000
   Private Const WS_EX_WINDOWEDGE = &H100&
   Private Const WS_EX_APPWINDOW = &H40000
   Private Const WS_EX_TOOLWINDOW = &H80&
   Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
   Private Const SW_HIDE = 0

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
   Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
   Private Const SWP_NOACTIVATE = &H10
   Private Const SWP_NOMOVE = &H2
   Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
   Private Const SWP_NOREDRAW = &H8
   Private Const SWP_NOSIZE = &H1
   Private Const SWP_NOZORDER = &H4
   Private Const SWP_SHOWWINDOW = &H40
   Private Const HWND_DESKTOP = 0
   Private Const HWND_NOTOPMOST = -2
   Private Const HWND_TOP = 0
   Private Const HWND_TOPMOST = -1
   Private Const HWND_BOTTOM = 1

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Private Const DT_LEFT = &H0&
    Private Const DT_TOP = &H0&
    Private Const DT_CENTER = &H1&
    Private Const DT_RIGHT = &H2&
    Private Const DT_VCENTER = &H4&
    Private Const DT_BOTTOM = &H8&
    Private Const DT_WORDBREAK = &H10&
    Private Const DT_SINGLELINE = &H20&
    Private Const DT_EXPANDTABS = &H40&
    Private Const DT_TABSTOP = &H80&
    Private Const DT_NOCLIP = &H100&
    Private Const DT_EXTERNALLEADING = &H200&
    Private Const DT_CALCRECT = &H400&
    Private Const DT_NOPREFIX = &H800
    Private Const DT_INTERNAL = &H1000&
    Private Const DT_WORD_ELLIPSIS = &H40000

Private Type OSVERSIONINFO
   dwVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion(0 To 127) As Byte
End Type
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_MAX = RGN_COPY
Private Const RGN_MIN = RGN_AND
Private Const WINDING = 2
Private Const ALTERNATE = 1

Public Enum etpStandardIcons
   etpNone
   IDI_ERROR = 32513 ' Stop Error icon
   IDI_QUESTION = 32514 'Question mark icon.
   IDI_WARNING = 32515 'Exclamation point icon.
   IDI_INFORMATION = 32516 'Asterisk icon.
End Enum

Public Enum etpShowDirection
   etpShowBelow
   etpShowAbove
End Enum

Public Event Click()
Public Event TimeOut()

Private m_hWnd As Long
Private m_bDesignTime As Boolean
Private m_bIsNt As Boolean
Private m_bIsXp As Boolean
Private m_bShown As Boolean
Private m_eShowDirection As etpShowDirection
Private m_hRgn As Long
Private m_hRgnFrame As Long

Private m_lWidth As Long
Private m_lHeight As Long
Private m_lMinWidth As Long
Private m_lBubbleArrowSize As Long
Private m_rcText As RECT
Private m_rcTitle As RECT

Private m_sTitle As String
Private m_sText As String
Private m_hIml As Long
Private m_lIcon As Long
Private m_eStandardIcon As etpStandardIcons
Private m_hIcon As Long
Private m_bShowCloseButton As Boolean
Private m_lTimeOut As Long
Private m_tP As POINTAPI, m_hWndRelativeTo As Long


Public Property Get Showing() As Boolean
'On Error Resume Next
   Showing = m_bShown
End Property
Public Sub Show(ByVal hWndRelativeTo As Long, ByVal X As Long, ByVal Y As Long)
'On Error Resume Next
    Timer1.Enabled = False
    m_tP.X = X
    m_tP.Y = Y
    m_hWndRelativeTo = hWndRelativeTo
    Timer1.Enabled = True
End Sub
Public Sub Hide()
'On Error Resume Next
   pHidePopup
End Sub

Public Property Get ShowCloseButton() As Boolean
'On Error Resume Next
   ShowCloseButton = m_bShowCloseButton
End Property
Public Property Let ShowCloseButton(ByVal bState As Boolean)
'On Error Resume Next
   m_bShowCloseButton = bState
   pEvalSize
   pPaint
   PropertyChanged "ShowCloseButton"
End Property

Public Property Get hIml() As Long
'On Error Resume Next
   hIml = m_hIml
End Property
Public Property Let hIml(ByVal lhIml As Long)
'On Error Resume Next
   m_hIml = lhIml
   pEvalSize
   pPaint
End Property
Public Property Get Title() As String
'On Error Resume Next
   Title = m_sTitle
End Property
Public Property Let Title(ByVal sTitle As String)
'On Error Resume Next
   m_sTitle = sTitle
   pEvalSize
   pPaint
   PropertyChanged "Title"
End Property
Public Property Get StandardIcon() As etpStandardIcons
'On Error Resume Next
   StandardIcon = m_eStandardIcon
End Property
Public Property Let StandardIcon(ByVal eIcon As etpStandardIcons)
'On Error Resume Next
   m_eStandardIcon = eIcon
   If Not (m_hIcon = 0) Then
      DestroyIcon m_hIcon
   End If
   If (eIcon = IDI_ERROR) Or (eIcon = IDI_INFORMATION) Or (eIcon = IDI_QUESTION) Or (eIcon = IDI_WARNING) Then
      m_hIcon = LoadIconString(0, "#" & eIcon)
   End If
   pEvalSize
   PropertyChanged "StandardIcon"
End Property
Public Property Get IconIndex() As Long
'On Error Resume Next
   IconIndex = m_lIcon
End Property
Public Property Let IconIndex(ByVal lIndex As Long)
'On Error Resume Next
   m_lIcon = lIndex
   pEvalSize
End Property
Public Property Get Text() As String
'On Error Resume Next
   Text = m_sText
End Property
Public Property Let Text(ByVal sText As String)
'On Error Resume Next
   m_sText = sText
   pEvalSize
   pPaint
   PropertyChanged "Text"
End Property
Public Property Get TimeOut() As Long
'On Error Resume Next
   TimeOut = m_lTimeOut
End Property
Public Property Let TimeOut(ByVal lTimeOut As Long)
'On Error Resume Next
   m_lTimeOut = lTimeOut
   PropertyChanged "TimeOut"
End Property

Public Property Get BackColor() As OLE_COLOR
'On Error Resume Next
   BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
'On Error Resume Next
   UserControl.BackColor = oColor
   pPaint
   PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
'On Error Resume Next
   ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
'On Error Resume Next
   UserControl.ForeColor = oColor
   pPaint
   PropertyChanged "ForeColor"
End Property
Public Property Get Font() As IFont
'On Error Resume Next
   Set Font = UserControl.Font
End Property
Public Property Let Font(ByVal fnt As IFont)
'On Error Resume Next
   pSetFont fnt
   PropertyChanged "Font"
End Property
Private Sub pSetFont(fnt As IFont)
'On Error Resume Next
   Set UserControl.Font = fnt
   pEvalSize
   pPaint
End Sub

Private Sub pPaint()
'On Error Resume Next
Dim lHDC As Long
Dim lTOp As Long
Dim rcWork As RECT
Dim hBr As Long
Dim lR As Long

   UserControl.Cls
   
   lHDC = UserControl.hDC
   
   hBr = CreateSolidBrush(TranslateColor(vbButtonShadow))
   lR = FrameRgn(lHDC, m_hRgnFrame, hBr, 1, 1)
'   Debug.Print lR
   DeleteObject hBr
   
   
   If (m_eShowDirection = etpShowBelow) Then
      lTOp = m_lBubbleArrowSize
   End If
      
   ' Draw the icon if any
   If Not (m_hIcon = 0) Then
      DrawIconEx lHDC, 8, 4 + lTOp, m_hIcon, 16, 16, 0, 0, DI_NORMAL
   ElseIf Not (m_hIml = 0) And (m_lIcon > -1) Then
   End If
   
   SetTextColor lHDC, TranslateColor(ForeColor)
   
   ' Draw the caption
   Dim iFntNow As IFont
   Set iFntNow = UserControl.Font
   Set UserControl.Font = BoldFont
   LSet rcWork = m_rcTitle
   OffsetRect rcWork, 0, lTOp
   DrawText lHDC, m_sTitle, -1, rcWork, DT_SINGLELINE Or DT_VCENTER
   Set UserControl.Font = iFntNow
   
   ' Draw the close button if required
   
   
   ' Draw the text
   LSet rcWork = m_rcText
   OffsetRect rcWork, 0, lTOp
   DrawText lHDC, m_sText, -1, rcWork, DT_WORDBREAK
   
   UserControl.Refresh
   
End Sub

Private Property Get IFontOf(iFnt As IFont)
'On Error Resume Next
   Set IFontOf = iFnt
End Property

Private Property Get BoldFont() As IFont
'On Error Resume Next
Dim sFntBold As New StdFont
Dim iFnt As IFont
   Set iFnt = UserControl.Font
   iFnt.Clone sFntBold
   sFntBold.Bold = True
   Set BoldFont = sFntBold
End Property

Private Sub pEvalSize()
'On Error Resume Next
Dim lMaxWidth As Long
Dim lWidth As Long
Dim lTitleHeight As Long
Dim lTextWidth As Long
Dim lHeight As Long
Dim rc As RECT
Dim lHDC As Long

   lHDC = UserControl.hDC

   '
   ' Determine the size of the title
   '
   If Len(m_sTitle) > 0 Then
      Dim iFntNow As IFont
      Set iFntNow = UserControl.Font
      Set UserControl.Font = BoldFont
      DrawText lHDC, m_sTitle, -1, rc, DT_CALCRECT Or DT_SINGLELINE
      Set UserControl.Font = iFntNow
   End If
   m_rcTitle.top = 4
   m_rcTitle.left = 4
   m_rcTitle.right = 4 + rc.right - rc.left
   m_rcTitle.bottom = 4 + rc.bottom - rc.top
   lTextWidth = rc.right - rc.left
   lTitleHeight = rc.bottom - rc.top
   ' Add spaces:
   lWidth = lTextWidth + 16
      
   ' Add Size of the close button
   If (m_bShowCloseButton) Then
      lWidth = lWidth + 16 + 8
      If (lTitleHeight < 20) Then
         lTitleHeight = 20
      End If
      OffsetRect m_rcTitle, 20, 0
   End If
   If (m_eStandardIcon = IDI_ERROR Or m_eStandardIcon = IDI_INFORMATION Or m_eStandardIcon = IDI_QUESTION Or m_eStandardIcon = IDI_WARNING) _
      Or (Not (m_hIml = 0) And (m_lIcon > -1)) Then
      lWidth = lWidth + 16 + 8
      OffsetRect m_rcTitle, 24, 0
      If (lTitleHeight < 20) Then
         lTitleHeight = 20
         OffsetRect m_rcTitle, 0, (20 - m_rcTitle.bottom - m_rcTitle.top) \ 2
      End If
   End If
   
   If (lWidth < m_lMinWidth) Then
      lWidth = m_lMinWidth
      lTextWidth = lWidth - 16
   End If
   
   '
   ' Evaluate the size of the text
   '
   m_rcText.left = 8
   m_rcText.right = m_rcText.left + lWidth - 16
   m_rcText.top = lTitleHeight + 4
   m_rcText.bottom = 512
   DrawText lHDC, m_sText, -1, m_rcText, DT_WORDBREAK Or DT_CALCRECT
   '
   
   m_lWidth = lWidth
   m_lHeight = m_rcText.top + m_rcText.bottom - m_rcText.top + 8 + m_lBubbleArrowSize
   
   pSetRegion
   '
End Sub
Private Sub pSetRegion()
'On Error Resume Next
   '
   
   Dim hRgnMain As Long
   hRgnMain = CreateRoundRectRgn(0, IIf(m_eShowDirection = etpShowAbove, 0, m_lBubbleArrowSize), m_lWidth, m_lHeight, 16, 16)
   Dim hRgnBubble As Long
   ReDim tP(0 To 2) As POINTAPI
   If (m_eShowDirection = etpShowAbove) Then
      tP(0).X = 32
      tP(0).Y = m_lHeight - m_lBubbleArrowSize
      tP(1).X = 32 + m_lBubbleArrowSize
      tP(1).Y = m_lHeight - m_lBubbleArrowSize
      tP(2).X = 32 + m_lBubbleArrowSize
      tP(2).Y = m_lHeight
   Else
      tP(0).X = 32
      tP(0).Y = 0
      tP(1).X = 32
      tP(1).Y = m_lBubbleArrowSize
      tP(2).X = 32 + m_lBubbleArrowSize
      tP(2).Y = m_lBubbleArrowSize
   End If
   hRgnBubble = CreatePolygonRgn(tP(0), 3, WINDING)
   Dim hRgn As Long
   Dim lR As Long
   hRgn = CreateRectRgn(0, 0, 0, 0)
   lR = CombineRgn(hRgn, hRgnMain, hRgnBubble, RGN_OR)
   If Not (m_hRgnFrame = 0) Then
      DeleteObject m_hRgnFrame
   End If
   m_hRgnFrame = CreateRectRgn(0, 0, 0, 0)
   lR = CombineRgn(m_hRgnFrame, hRgnMain, hRgnBubble, RGN_OR)
   DeleteObject hRgnMain
   DeleteObject hRgnBubble
   
   If Not (m_hWnd = 0) Then
      SetWindowRgn m_hWnd, hRgn, 0
   End If
   If (m_hRgn) Then
      DeleteObject m_hRgn
   End If
   m_hRgn = hRgn
   
   '
End Sub

Private Sub pShowPopup(ByVal X As Long, ByVal Y As Long)
'On Error Resume Next
Dim rc As RECT
   
   pEvalSize
   ' Set the style of the object so it works as a popup:
   Dim lStyle As Long
   lStyle = GetWindowLong(m_hWnd, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_TOOLWINDOW
   lStyle = lStyle And Not (WS_EX_APPWINDOW)
   SetWindowLong m_hWnd, GWL_EXSTYLE, lStyle
   SetParent m_hWnd, HWND_DESKTOP
   SetProp m_hWnd, MAGIC_END_EDIT_IGNORE_WINDOW_PROP, 1
   SetWindowPos m_hWnd, HWND_TOPMOST, X, Y, m_lWidth, m_lHeight, SWP_SHOWWINDOW
   pPaint
   m_bShown = True
   If (m_lTimeOut > -1) Then
      tmrTimeOut.Tag = timeGetTime
      tmrTimeOut.Enabled = True
   End If
   
End Sub

Private Sub pHidePopup()
'On Error Resume Next
   If (m_bShown) Then
      ShowWindow m_hWnd, SW_HIDE
      RemoveProp m_hWnd, MAGIC_END_EDIT_IGNORE_WINDOW_PROP
      m_bShown = False
   End If
End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
'On Error Resume Next
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
'On Error Resume Next
Dim objT As Object
   If Not (lPtr = 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory objT, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set ObjectFromPtr = objT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory objT, 0&, 4
   End If
End Property

Private Sub VerInitialise()
'On Error Resume Next
   
   Dim tOSV As OSVERSIONINFO
   tOSV.dwVersionInfoSize = Len(tOSV)
   GetVersionEx tOSV
   
   m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
   If (tOSV.dwMajorVersion > 5) Then
      'm_bHasGradientAndTransparency = True
      m_bIsXp = True
      'm_bIs2000OrAbove = True
   ElseIf (tOSV.dwMajorVersion = 5) Then
      'm_bHasGradientAndTransparency = True
      'm_bIs2000OrAbove = True
      If (tOSV.dwMinorVersion >= 1) Then
         m_bIsXp = True
      End If
   ElseIf (tOSV.dwMajorVersion = 4) Then ' NT4 or 9x/ME/SE
      'If (tOSV.dwMinorVersion >= 10) Then
      '   m_bHasGradientAndTransparency = True
      'End If
   Else ' Too old
   End If
   
End Sub

Private Sub DrawText( _
      ByVal lHDC As Long, _
      ByVal sText As String, _
      ByVal lLength As Long, _
      tR As RECT, _
      ByVal lFlags As Long _
   )
'On Error Resume Next
Dim lPtr As Long
   If (m_bIsNt) Then
      lPtr = StrPtr(sText)
      If Not (lPtr = 0) Then ' NT4 crashes with ptr = 0
         DrawTextW lHDC, lPtr, -1, tR, lFlags
      End If
   Else
      DrawTextA lHDC, sText, -1, tR, lFlags
   End If
End Sub

Private Sub pInitialise()
'On Error Resume Next
   
   m_bDesignTime = Not (UserControl.Ambient.UserMode)
   m_hWnd = UserControl.hWnd
   If (m_bDesignTime) Then
      pEvalSize
      pPaint
   Else
      UserControl.Extender.Visible = False
   End If
   
End Sub

Private Sub Timer1_Timer()
   ClientToScreen m_hWndRelativeTo, m_tP
   pShowPopup m_tP.X, m_tP.Y
   Timer1.Enabled = False
End Sub

Private Sub tmrTimeOut_Timer()
'On Error Resume Next
   '
Dim lT As Long
   If Len(tmrTimeOut.Tag) > 0 Then
      lT = CLng(tmrTimeOut.Tag)
      If (timeGetTime - lT > m_lTimeOut) Then
         RaiseEvent TimeOut
         pHidePopup
      End If
   Else
      tmrTimeOut.Enabled = False
   End If
   '
End Sub

Private Sub UserControl_Initialize()
'On Error Resume Next
   m_lTimeOut = -1
   m_eStandardIcon = etpNone
   VerInitialise
   m_lMinWidth = 220
   m_lBubbleArrowSize = 16
End Sub

Private Sub UserControl_InitProperties()
'On Error Resume Next
   '
   pInitialise
   '
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
   '
   If (m_bShown) Then
      RaiseEvent Click
      pHidePopup
   End If
   '
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'On Error Resume Next
   '
   Title = PropBag.ReadProperty("Title", "")
   Text = PropBag.ReadProperty("Text", "")
   StandardIcon = PropBag.ReadProperty("StandardIcon", etpNone)
   TimeOut = PropBag.ReadProperty("TimeOut", -1)
   BackColor = PropBag.ReadProperty("BackColor", vbInfoBackground)
   ForeColor = PropBag.ReadProperty("ForeColor", vbInfoText)
   Dim sFnt As New StdFont
   sFnt.Name = "Tahoma"
   sFnt.Size = 8.25
   Font = PropBag.ReadProperty("Font", sFnt)
   ShowCloseButton = PropBag.ReadProperty("ShowCloseButton", False)
   
   pInitialise
   '
End Sub

Private Sub UserControl_Resize()
'On Error Resume Next
   '
   pPaint
   '
End Sub

Private Sub UserControl_Show()
'On Error Resume Next
   '
End Sub

Private Sub UserControl_Terminate()
'On Error Resume Next
   pHidePopup
   DeleteObject m_hRgnFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'On Error Resume Next
   '
   PropBag.WriteProperty "Title", Title, ""
   PropBag.WriteProperty "Text", Text, ""
   PropBag.WriteProperty "StandardIcon", StandardIcon, etpNone
   PropBag.WriteProperty "TimeOut", TimeOut, -1
   PropBag.WriteProperty "BackColor", BackColor, vbInfoBackground
   PropBag.WriteProperty "ForeColor", ForeColor, vbInfoText
   Dim sFnt As New StdFont
   sFnt.Name = "Tahoma"
   sFnt.Size = 8.25
   PropBag.WriteProperty "Font", Font, sFnt
   PropBag.WriteProperty "ShowCloseButton", ShowCloseButton, False
   '
End Sub


