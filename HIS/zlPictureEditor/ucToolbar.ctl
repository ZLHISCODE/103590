VERSION 5.00
Begin VB.UserControl ucToolbar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ClipControls    =   0   'False
   ForeColor       =   &H80000014&
   LockControls    =   -1  'True
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ToolboxBitmap   =   "ucToolbar.ctx":0000
   Begin VB.Timer tmrTip 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblTipRect 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   -375
      TabIndex        =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "ucToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BF_RECT         As Long = &HF
Private Const ILC_MASK        As Long = &H1
Private Const ILC_COLORDDB    As Long = &HFE
Private Const ILD_TRANSPARENT As Long = 1
Private Const DST_ICON        As Long = &H3
Private Const DSS_DISABLED    As Long = &H20
Private Const CLR_INVALID     As Long = &HFFFF
Private Const COLOR_BTNFACE   As Long = 15

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal dx As Long, ByVal dy As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2, ByVal m_hBrush As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Const VK_MBUTTON As Long = &H4

'//

'-- Public Enums.:
Public Enum tbOrientationConstants
    [tbHorizontal]
    [tbVertical]
End Enum

'-- Private Enums.:
Private Enum tbButtonStateConstants
    [btDown] = -1
    [btFlat] = 0
    [btOver] = 1
End Enum

Private Enum tbButtonTypeConstants
    [btNormal] = 0
    [btCheck] = 1
    [btOption] = 2
End Enum

Private Enum tbMouseEventConstants
    [btMouseDown] = -1
    [btMouseMove] = 0
    [btMouseUp] = 1
End Enum

'-- Private Types:
Private Type tButton
    Type      As tbButtonTypeConstants
    State     As tbButtonStateConstants
    Enabled   As Boolean
    Checked   As Boolean
    Over      As Boolean
    Separator As RECT2
End Type

'-- Private Constants:
Private Const BTN_STL_NORMAL As String = "N"
Private Const BTN_STL_CHECK  As String = "C"
Private Const BTN_STL_OPTION As String = "O"
Private Const BTN_SEPARATOR  As String = "|"
Private Const SEP_LENGTH     As Long = 8
Private Const IMG_BORDER     As Long = 3
Private Const IMG_OFFSET     As Long = 1
   
'-- Default Property Values:
Private Const m_def_BarOrientation As Integer = [tbHorizontal]
Private Const m_def_BarEdge        As Boolean = False

'-- Property Variables:
Private m_BarOrientation As tbOrientationConstants
Private m_BarEdge        As Boolean

'-- Private Variables:
Private m_hIL            As Long    ' Image list handle
Private m_hBrush         As Long    ' Brush (check effect)
Private m_BarRect        As RECT2   ' Bar rectangle
Private m_ExtRect()      As RECT2   ' Button rects. (edge area)
Private m_ClkRect()      As RECT2   ' Button rects. (click area)
Private m_uButton()      As tButton ' Buttons
Private m_ToolTip()      As String  ' Tool tips
Private m_Count          As Integer ' Button count
Private m_LastOver       As Integer ' Last over
Private m_IconSize       As Integer ' Icon size (W = H)
Private m_ButtonSize     As Integer ' Button size (W = H)

'-- Event Declarations:
Public Event ButtonClick(ByVal Index As Long, ByVal xRight As Long, ByVal yTop As Long)
Public Event ButtonCheck(ByVal Index As Long, ByVal xRight As Long, ByVal yTop As Long)
Public Event MouseDown(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xRight As Long, ByVal yTop As Long)
Public Event MouseUp(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xRight As Long, ByVal yTop As Long)



'==================================================================================================
' UserControl
'==================================================================================================

Private Sub UserControl_Initialize()

  Dim aIdx           As Byte
  Dim nBytes(1 To 8) As Integer
  Dim hBitmap        As Long
    
    '-- Build brush for check effect
    For aIdx = 1 To 8 Step 2
        nBytes(aIdx) = &HAA
        nBytes(aIdx + 1) = &H55
    Next aIdx
    hBitmap = CreateBitmap(8, 8, 1, 1, nBytes(1))
    m_hBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy image list and pattern brush
    Call pvDestroyIL
    Call DeleteObject(m_hBrush)
End Sub

'//

Private Sub UserControl_Show()

    '-- Refresh on start up
    Call pvRefresh
End Sub

Private Sub UserControl_Resize()
    
    '-- Adjust for alignment
    Select Case m_BarOrientation
        Case [tbHorizontal]
            m_BarRect.x2 = ScaleWidth
        Case [tbVertical]
            m_BarRect.y2 = ScaleHeight
    End Select
    '-- Refresh whole control
    Call FillRect(hdc, m_BarRect, GetSysColorBrush(COLOR_BTNFACE))
    Call pvRefresh
End Sub

'//

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Dim nBtn As Integer
    
    '-- Restore last
    If (m_LastOver) Then
        Call pvUpdateButtonState(m_LastOver, 0, 0, [btMouseMove])
    End If
    
    '-- Update tooltip label pos.
    For nBtn = 1 To m_Count
        If (PtInRect(m_ExtRect(nBtn), X, Y) And m_uButton(nBtn).Enabled) Then
            Call pvSetTipArea(nBtn)
            m_LastOver = nBtn
        End If
    Next nBtn
End Sub

Private Sub lblTipRect_DblClick()
     
    If (GetAsyncKeyState(VK_RBUTTON) >= 0 And GetAsyncKeyState(VK_MBUTTON) >= 0) Then '*
        '-- Preserve second click
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    End If
    
'*: Should be previously checked GetSystemMetrics(SM_SWAPBUTTON)
End Sub

Private Sub lblTipRect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  Dim lx As Long
  Dim ly As Long
    
    If (m_LastOver) Then
        If (m_uButton(m_LastOver).Enabled) Then
            '-- Translate to [pixels]
            lx = X \ Screen.TwipsPerPixelX + lblTipRect.Left
            ly = Y \ Screen.TwipsPerPixelY + lblTipRect.Top
            '-- Refresh state [?]
            If (PtInRect(m_ExtRect(m_LastOver), lx, ly) <> 0) Then
                Call pvUpdateButtonState(m_LastOver, True, Button, [btMouseDown])
            End If
        End If
        RaiseEvent MouseDown(m_LastOver, Button, m_ExtRect(m_LastOver).x2, m_ExtRect(m_LastOver).y1)
    End If
End Sub

Private Sub lblTipRect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  Dim lx As Long
  Dim ly As Long
    
    If (m_LastOver) Then
        If (m_uButton(m_LastOver).Enabled) Then
            '-- Translate to [pixels]
            lx = X \ Screen.TwipsPerPixelX + lblTipRect.Left
            ly = Y \ Screen.TwipsPerPixelY + lblTipRect.Top
            '-- Refresh state
            Call pvUpdateButtonState(m_LastOver, PtInRect(m_ExtRect(m_LastOver), lx, ly) <> 0, Button, [btMouseMove])
        End If
        If (Button = vbLeftButton) Then tmrTip.Enabled = True
    End If
End Sub

Private Sub lblTipRect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Dim lx As Long
  Dim ly As Long
    
    If (m_LastOver) Then
        If (m_uButton(m_LastOver).Enabled) Then
            '-- Translate to [pixels]
            lx = X \ Screen.TwipsPerPixelX + lblTipRect.Left
            ly = Y \ Screen.TwipsPerPixelY + lblTipRect.Top
            '-- Refresh state
            Call pvUpdateButtonState(m_LastOver, PtInRect(m_ExtRect(m_LastOver), lx, ly) <> 0, Button, [btMouseUp])
        End If
        RaiseEvent MouseUp(m_LastOver, Button, m_ExtRect(m_LastOver).x2, m_ExtRect(m_LastOver).y1)
    End If
End Sub

'==================================================================================================
' Methods
'==================================================================================================

Public Sub Refresh()

    '-- Refresh shole bar
    Call pvRefresh
End Sub

Public Function BuildToolbar(Image As StdPicture, ByVal MaskColor As OLE_COLOR, ByVal IconSize As Integer, Optional ByVal FormatMask As String) As Boolean
    
  Dim nIdx As Integer
  Dim nBtn As Integer
  Dim sKey As String
  Dim lPos As Long
    
    If (pvExtractImages(Image, MaskColor, IIf(IconSize > 0, IconSize, 1))) Then
        
        '-- Missing 'FormatMask': Normal buttons, no separators
        If (FormatMask = vbNullString) Then
            FormatMask = String$(ImageList_GetImageCount(m_hIL), BTN_STL_NORMAL)
        End If
        
        '-- Button ext. size (image[state] and edge offsets)
        m_ButtonSize = m_IconSize + 2 * IMG_BORDER + IMG_OFFSET
        
        '-- Extract buttons...
        Do While nIdx < Len(FormatMask)
            
            '-- Key count / extract key
            nIdx = nIdx + 1
            sKey = UCase$(Mid$(FormatMask, nIdx, 1))
            
            Select Case sKey
                
                '-- Normal button, check button and option buttons
                Case BTN_STL_NORMAL, BTN_STL_CHECK, BTN_STL_OPTION
                
                    nBtn = nBtn + 1
                    lPos = lPos + m_ButtonSize
                    
                    '-- Redim. button rectangles
                    ReDim Preserve m_ExtRect(1 To nBtn)
                    ReDim Preserve m_ClkRect(1 To nBtn)
                    ReDim Preserve m_uButton(1 To nBtn)
                    '-- Store button type
                    Select Case sKey
                        Case BTN_STL_NORMAL: m_uButton(nBtn).Type = [btNormal]
                        Case BTN_STL_CHECK:  m_uButton(nBtn).Type = [btCheck]
                        Case BTN_STL_OPTION: m_uButton(nBtn).Type = [btOption]
                    End Select
                    '-- Enabled [?]
                    m_uButton(nBtn).Enabled = UserControl.Enabled
                    
                    '-- Button ext. rect.
                    Select Case m_BarOrientation
                        Case [tbHorizontal]
                            Call SetRect(m_ExtRect(nBtn), lPos - m_ButtonSize, 0, lPos, m_ButtonSize - 1)
                        Case [tbVertical]
                            Call SetRect(m_ExtRect(nBtn), 0, lPos - m_ButtonSize, m_ButtonSize - 1, lPos)
                    End Select
                    Call OffsetRect(m_ExtRect(nBtn), 1, 1)
                    '-- Button click rect.
                    m_ClkRect(nBtn) = m_ExtRect(nBtn)
                    Call InflateRect(m_ClkRect(nBtn), -2, -2)
               
                '-- Separator
                Case BTN_SEPARATOR
                
                    lPos = lPos + SEP_LENGTH
                    With m_ClkRect(nBtn)
                        Select Case m_BarOrientation
                            Case [tbHorizontal]
                                Call SetRect(m_uButton(nBtn).Separator, .x2 + (SEP_LENGTH \ 2 + 1), .y1, .x2 + (SEP_LENGTH \ 2 + 3), .y2)
                            Case [tbVertical]
                                Call SetRect(m_uButton(nBtn).Separator, .x1, .y2 + (SEP_LENGTH \ 2 + 1), .x2, .y2 + (SEP_LENGTH \ 2 + 3))
                        End Select
                    End With
            End Select
        Loop
        
        '-- Resize control
        With m_ExtRect(nBtn)
            UserControl.Width = (.x2 + 1) * Screen.TwipsPerPixelX
            UserControl.Height = (.y2 + 1) * Screen.TwipsPerPixelY
        End With
        Call SetRect(m_BarRect, 0, 0, ScaleWidth, ScaleHeight)
        
        '-- Buttons count / success
        m_Count = nBtn
        BuildToolbar = (m_Count > 0)
    End If
End Function

Public Sub SetTooltips(ByVal TooltipsList As String)
    '-- Extract tooltips...
    m_ToolTip() = Split(TooltipsList, BTN_SEPARATOR)
End Sub

Public Sub SetTooltip(ByVal Index As Integer, ByVal Tooltip As String)
    m_ToolTip(Index) = Tooltip
End Sub
Public Function GetTooltip(ByVal Index As Integer) As String
    GetTooltip = m_ToolTip(Index)
End Function
Public Sub EnableButton(ByVal Index As Integer, ByVal Enable As Boolean)
    Call pvEnableButton(Index, Enable)
End Sub

Public Function IsButtonEnabled(ByVal Index As Integer) As Boolean
    IsButtonEnabled = m_uButton(Index).Enabled
End Function

Public Sub CheckButton(ByVal Index As Integer, ByVal Check As Boolean)

    If (m_Count) Then
        If (Index And Index <= m_Count) Then
            If (m_uButton(Index).Type <> [btNormal] And m_uButton(Index).Checked <> Check) Then
                    
                '-- Update button
                With m_uButton(Index)
                    .Checked = Check
                    .State = [btDown] And Check
                End With
                Call pvRefresh(Index)
                Call pvUpdateOptionButtons(Index)
                '-- Update Tooltip label pos.
                Call pvSetTipArea(Index)
                '-- Store <last over> index
                m_LastOver = Index
            
                '-- Raise <Check> event
                With m_ExtRect(Index)
                    RaiseEvent ButtonCheck(Index, .x2, .y1)
                End With
            End If
        End If
    End If
End Sub
Public Function IsButtonChecked(ByVal Index As Integer) As Boolean
    IsButtonChecked = m_uButton(Index).Checked
End Function

'==================================================================================================
' Private
'==================================================================================================

Private Function pvExtractImages(Image As StdPicture, ByVal MaskColor As OLE_COLOR, ByVal IconSize As Integer) As Boolean
    
    '-- Extract images
    If (Not Image Is Nothing) Then
        If (pvCreateIL(IconSize)) Then
            pvExtractImages = (ImageList_AddMasked(m_hIL, Image.handle, pvTranslateColor(MaskColor)) <> -1)
        End If
    End If
End Function

Private Function pvCreateIL(ByVal IconSize As Integer) As Boolean
     
    '-- Destroy previous [?]
    Call pvDestroyIL
    '-- Create the image list object:
    m_hIL = ImageList_Create(IconSize, IconSize, ILC_MASK Or ILC_COLORDDB, 0, 0)
    If (m_hIL <> 0 And m_hIL <> -1) Then
        m_IconSize = IconSize
        pvCreateIL = True
      Else
        m_hIL = 0
    End If
End Function

Private Sub pvDestroyIL()

    '-- Kill the image list if we have one:
    If (m_hIL <> 0) Then
        Call ImageList_Destroy(m_hIL)
        m_hIL = 0
    End If
End Sub

'//

Private Sub pvSetTipArea(ByVal Index As Integer)
    
    '-- Move label
    Select Case m_BarOrientation
        Case [tbHorizontal]
            Call lblTipRect.Move(m_ExtRect(Index).x1, 0, m_ButtonSize, m_ButtonSize)
        Case [tbVertical]
            Call lblTipRect.Move(0, m_ExtRect(Index).y1, m_ButtonSize, m_ButtonSize)
    End Select
    '-- Set tool tip text
    On Error Resume Next
       lblTipRect.ToolTipText = m_ToolTip(Index - 1)
    On Error GoTo 0
End Sub

'//

Private Sub pvEnableBar(ByVal bEnable As Boolean)

  Dim nBtn As Integer
    
    If (m_Count) Then
        '-- Enable/disable
        For nBtn = 1 To m_Count
            m_uButton(nBtn).Enabled = bEnable
        Next nBtn
        '-- Refresh
        Call pvRefresh
    End If
End Sub

Private Sub pvEnableButton(ByVal Index As Integer, ByVal bEnable As Boolean)
    
    If (m_Count) Then
        If (Index And Index <= m_Count And m_uButton(Index).Enabled <> bEnable) Then
            '-- Enable/disable
            With m_uButton(Index)
                If (Not bEnable And .Type = [btNormal]) Then
                    .State = [btFlat]
                End If
                .Enabled = bEnable
            End With
            '-- Refresh
            Call pvRefresh(Index)
        End If
    End If
End Sub

'//

Private Sub pvRefresh(Optional ByVal Index As Integer = 0)

  Dim nBtn As Integer
    
    If (m_Count) Then
        If (Index = 0) Then
            '== All buttons...
            '-- Draw buttons
            For nBtn = 1 To m_Count
                Call pvPaintButton(nBtn)
                Call pvPaintBitmap(nBtn)
                If (IsRectEmpty(m_uButton(nBtn).Separator) = 0) Then
                    Call DrawEdge(hdc, m_uButton(nBtn).Separator, BDR_SUNKENOUTER, BF_RECT)
                End If
            Next nBtn
          Else
            '== Single button
            Call pvPaintButton(Index)
            Call pvPaintBitmap(Index)
        End If
        '-- Flat border [?]
        If (m_BarEdge) Then
            Call DrawEdge(hdc, m_BarRect, BDR_RAISEDINNER, BF_RECT)
        End If
        
        '-- Refresh
        Call UserControl.Refresh
    End If
End Sub

Private Sub pvPaintButton(ByVal Index As Integer)
    
    '-- Background
    If (m_uButton(Index).Checked And m_uButton(Index).State = [btDown] And Not m_uButton(Index).Over) Then
        Call FillRect(hdc, m_ClkRect(Index), m_hBrush)
      Else
        Call FillRect(hdc, m_ExtRect(Index), GetSysColorBrush(COLOR_BTNFACE))
    End If
    '-- Edge
    Select Case m_uButton(Index).State
        Case [btOver]
            Call DrawEdge(hdc, m_ExtRect(Index), BDR_RAISEDINNER, BF_RECT)
        Case [btDown]
            Call DrawEdge(hdc, m_ExtRect(Index), BDR_SUNKENOUTER, BF_RECT)
    End Select
End Sub

Private Sub pvPaintBitmap(ByVal Index As Integer)
  
  Dim lOffset As Long
  
    '-- Image offset
    lOffset = IMG_BORDER + (IMG_OFFSET * -(m_uButton(Index).State = [btDown]))
    '-- Paint masked bitmap
    With m_ExtRect(Index)
        Call pvDrawImage(Index, hdc, .x1 + lOffset, .y1 + lOffset)
    End With
End Sub

Private Sub pvDrawImage(ByVal Index As Integer, ByVal hdc As Long, ByVal X As Integer, ByVal Y As Integer)

  Dim hIcon As Long

    If (m_uButton(Index).Enabled) Then
        '-- Normal
        Call ImageList_Draw(m_hIL, Index - 1, hdc, X, Y, ILD_TRANSPARENT)
      Else
        '-- Disabled
        hIcon = ImageList_GetIcon(m_hIL, Index - 1, 0)
        Call DrawState(hdc, 0, 0, hIcon, 0, X, Y, m_IconSize, m_IconSize, DST_ICON Or DSS_DISABLED)
        Call DestroyIcon(hIcon)
    End If
End Sub

Private Function pvTranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    
    '-- OLE/RGB color to RGB color
    If (OleTranslateColor(clr, hPal, pvTranslateColor)) Then
        pvTranslateColor = CLR_INVALID
    End If
End Function

'//

Private Sub pvUpdateButtonState(ByVal Index As Integer, ByVal InButton As Boolean, ByVal MouseButton As MouseButtonConstants, ByVal MouseEvent As tbMouseEventConstants)
    
  Dim uTmpButton  As tButton
    
    '-- Store current button state
    uTmpButton = m_uButton(Index)
    
    '-- Over button [?]
    m_uButton(Index).Over = InButton
    
    '-- Check new state
    With m_uButton(Index)
        
        Select Case MouseEvent
            
            Case [btMouseDown] '-- Mouse pressed
            
                If (MouseButton = vbLeftButton) Then
                    .State = [btDown]
                End If
                
             Case [btMouseMove] '-- Mouse moving
             
                If (InButton) Then
                    If (MouseButton = vbLeftButton) Then
                        .State = [btDown]
                      Else
                        If (Not .Checked) Then
                            .State = [btOver]
                        End If
                        tmrTip.Enabled = True
                    End If
                  Else
                    If (Not .Checked) Then
                        .State = [btFlat]
                    End If
                End If
                
            Case [btMouseUp]  '-- Mouse released
            
                 If (InButton) Then
                    If (MouseButton = vbLeftButton) Then
                        Select Case .Type
                            Case [btNormal]
                                .State = [btOver]
                            Case [btCheck]
                                .Checked = Not .Checked
                                .State = -.Checked * [btDown]
                            Case [btOption]
                                .Checked = True
                                .State = [btDown]
                                Call pvUpdateOptionButtons(Index)
                        End Select
                      Else
                        If (Not .Checked And MouseButton = vbEmpty) Then
                            .State = [btFlat]
                        End If
                    End If
                End If
        End Select
        
        '-- Refresh [?]
        If (.State <> uTmpButton.State Or .Checked <> uTmpButton.Checked Or .Over <> uTmpButton.Over) Then
            Call pvRefresh(Index)
        End If
        
        '-- Raise [Click]/[Check] event [?]
        If (InButton And MouseEvent = [btMouseUp] And MouseButton = vbLeftButton) Then
            
            Select Case m_uButton(Index).Type
                Case [btNormal]
                    RaiseEvent ButtonClick(Index, m_ExtRect(Index).x2, m_ExtRect(Index).y1)
                Case [btCheck], [btOption]
                    RaiseEvent ButtonClick(Index, m_ExtRect(Index).x2, m_ExtRect(Index).y1)
                    If (.Checked <> uTmpButton.Checked) Then
                        RaiseEvent ButtonCheck(Index, m_ExtRect(Index).x2, m_ExtRect(Index).y1)
                    End If
            End Select
        End If
    End With
End Sub

Private Sub pvUpdateOptionButtons(ByVal CurrentIndex As Integer)

  Dim nIdx As Integer
    
    '-- Right/below buttons
    nIdx = CurrentIndex
    Do While nIdx < m_Count
        If (IsRectEmpty(m_uButton(nIdx).Separator) = 0) Then
            Exit Do
          Else
            nIdx = nIdx + 1
            With m_uButton(nIdx)
                If (.Type = [btOption] And .Checked) Then
                    .Checked = False
                    .State = [btFlat]
                    Call pvRefresh(nIdx)
                End If
            End With
        End If
    Loop
    
    '-- Left/above buttons
    nIdx = CurrentIndex
    Do While nIdx > 1
        nIdx = nIdx - 1
        If (IsRectEmpty(m_uButton(nIdx).Separator) = 0) Then
            Exit Do
          Else
            With m_uButton(nIdx)
                If (.Type = [btOption] And .Checked) Then
                    .Checked = False
                    .State = [btFlat]
                     Call pvRefresh(nIdx)
                End If
            End With
        End If
    Loop
End Sub

'//

Private Sub tmrTip_Timer()
  
  Dim uPt As POINTAPI
    
    '-- Cursor out of toolbar [?]
    Call GetCursorPos(uPt)
    If (WindowFromPoint(uPt.X, uPt.Y) <> UserControl.hwnd) Then
        '-- Disable timer and refresh
        tmrTip.Enabled = False
        Call pvUpdateButtonState(m_LastOver, 0, 0, [btMouseMove])
    End If
End Sub

'==================================================================================================
' Properties
'==================================================================================================

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call pvEnableBar(New_Enabled)
End Property

Public Property Get BarOrientation() As tbOrientationConstants
    BarOrientation = m_BarOrientation
End Property
Public Property Let BarOrientation(ByVal New_BarOrientation As tbOrientationConstants)
    If (Not Ambient.UserMode) Then
        m_BarOrientation = New_BarOrientation
    End If
End Property

Public Property Get BarEdge() As Boolean
    BarEdge = m_BarEdge
End Property
Public Property Let BarEdge(ByVal New_BarEdge As Boolean)
    m_BarEdge = New_BarEdge
    Call pvRefresh
End Property

Public Property Get ButtonsCount() As Integer
    ButtonsCount = m_Count
End Property

'//

Private Sub UserControl_InitProperties()
    m_BarOrientation = m_def_BarOrientation
    m_BarEdge = m_def_BarEdge
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.Enabled = .ReadProperty("Enabled", True)
        m_BarOrientation = .ReadProperty("BarOrientation", m_def_BarOrientation)
        m_BarEdge = .ReadProperty("BarEdge", m_def_BarEdge)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("BarOrientation", m_BarOrientation, m_def_BarOrientation)
        Call .WriteProperty("BarEdge", m_BarEdge, m_def_BarEdge)
    End With
End Sub
