VERSION 5.00
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.UserControl ScrollBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   ScaleHeight     =   510
   ScaleWidth      =   1095
   ToolboxBitmap   =   "ScrollBar.ctx":0000
   Begin zlSubclass.Subclass Subclass1 
      Left            =   720
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin zlRichEditor.FButton cmdButton 
      Height          =   240
      Index           =   0
      Left            =   405
      Top             =   90
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin zlRichEditor.FButton optButton 
      Height          =   240
      Index           =   0
      Left            =   135
      Top             =   90
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      IsOptButton     =   -1  'True
   End
End
Attribute VB_Name = "ScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################################################
'##模 块 名：ScrollBar.ctl
'##创 建 人：吴庆伟
'##日    期：2005年3月1日
'##修 改 人：
'##日    期：
'##描    述：一个 Word 2003 风格的滚动条控件。可以在水平/垂直滚动条上加入自定义按钮。
'##版    本：
'######################################################################################

Option Explicit
Public Enum ESBCScrollTypes
   esbcHorizontal
   esbcVertical
   esbcSizeGripper
End Enum
Public Enum ESBCButtonPositionConstants
   esbcButtonPositionDefault
   esbcButtonPositionLeftTop
   esbcButtonPositionRightBottom
End Enum
Private Type tButtonInfo
   sKey As String
   sHelpText As String
   lIconIndexUp As Long
   lIconIndexDown As Long
   ePosition As ESBCButtonPositionConstants
   bCheck As Boolean
   sCheckGroup As String
   ctlThis As Control
End Type
Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   itemData As Long
End Type
Private m_hWndControl As Long
Private m_hWndParent As Long
Private m_hWnd As Long
Private m_eScrollType As ESBCScrollTypes
Private m_iButtonCount As Long
Private m_tButtons() As tButtonInfo
Private m_iOptCount As Long
Private m_iCmdCount As Long
Private m_lPos1 As Long
Private m_lPos2 As Long
Private m_hIml As Long
'Private m_ptrVb6ImageList As Long
Private m_lIconSizeX As Long
Private m_lIconSizeY As Long
Private m_lSmallChange As Long
Private m_bScrollEnabled As Boolean
Private m_bNoFlatScrollBars As Boolean
Private m_bXPStyleButtons As Boolean

Public Event ButtonClick(ByVal lButton As Long)
Public Event Change()
Public Event Scroll()

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get ButtonKey( _
      ByVal lButton As Long _
   ) As String
   If (ButtonIndex(lButton) > 0) Then
      ButtonKey = m_tButtons(lButton).sKey
   End If
End Property

Public Property Get ButtonToolTipText( _
      ByVal vKey As Variant _
   ) As String
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      ButtonToolTipText = m_tButtons(iBtnIndex).sHelpText
   End If
End Property

Public Property Let ButtonToolTipText( _
      ByVal vKey As Variant, _
      ByVal sText As String _
   )
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      m_tButtons(iBtnIndex).sHelpText = sText
      m_tButtons(iBtnIndex).ctlThis.ToolTipText = sText
   End If
End Property

Public Property Get ButtonVisible( _
      ByVal vKey As Variant _
   ) As Boolean
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      ButtonVisible = m_tButtons(iBtnIndex).ctlThis.Visible
   End If
End Property

Public Property Let ButtonVisible( _
      ByVal vKey As Variant, _
      ByVal bState As Boolean _
   )
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      m_tButtons(iBtnIndex).ctlThis.Visible = bState
      UserControl_Resize
   End If
End Property
Public Property Get ButtonEnabled( _
      ByVal vKey As Variant _
   ) As Boolean
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      ButtonEnabled = m_tButtons(iBtnIndex).ctlThis.Enabled
   End If
End Property

Public Property Let ButtonEnabled( _
      ByVal vKey As Variant, _
      ByVal bState As Boolean _
   )
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      m_tButtons(iBtnIndex).ctlThis.Enabled = bState
   End If
End Property

Public Property Get ButtonValue( _
      ByVal vKey As Variant _
   ) As Boolean
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      If (m_tButtons(iBtnIndex).ctlThis.IsOptButton) Then
         ButtonValue = Abs(m_tButtons(iBtnIndex).ctlThis.Value)
      Else
         ButtonValue = m_tButtons(iBtnIndex).ctlThis.Value
      End If
   End If
End Property


Public Property Let ButtonValue( _
      ByVal vKey As Variant, _
      oValue As Boolean _
   )
Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      If (m_tButtons(iBtnIndex).ctlThis.IsOptButton) Then
         m_tButtons(iBtnIndex).ctlThis.Value = -1 * oValue
      Else
         m_tButtons(iBtnIndex).ctlThis.Value = oValue
      End If
   End If
End Property


Private Function pTranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, pTranslateColor) Then
        pTranslateColor = CLR_INVALID
    End If
End Function

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property

'Private Sub pDrawImage( _
'      ByVal ptrVB6ImageList As Long, _
'      ByVal hIml As Long, _
'      ByVal iIndex As Long, _
'      ByVal hdc As Long, _
'      ByVal xPixels As Integer, _
'      ByVal yPixels As Integer, _
'      ByVal lIconSizeX As Long, ByVal lIconSizeY As Long, _
'      Optional ByVal bSelected = False, _
'      Optional ByVal bCut = False, _
'      Optional ByVal bDisabled = False, _
'      Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
'      Optional ByVal hExternalIml As Long = 0 _
'    )
'Dim hIcon As Long
'Dim lFlags As Long
'Dim lhIml As Long
'Dim lColor As Long
'Dim iImgIndex As Long
'
'   ' Draw the image at 1 based index or key supplied in vKey.
'   ' on the hDC at xPixels,yPixels with the supplied options.
'   ' You can even draw an ImageList from another ImageList control
'   ' if you supply the handle to hExternalIml with this function.
'
'   iImgIndex = iIndex
'   If (iImgIndex > -1) Then
'      If (hExternalIml <> 0) Then
'          lhIml = hExternalIml
'      Else
'          lhIml = hIml
'      End If
'
'      lFlags = ILD_Transparent
'      If (bSelected) Or (bCut) Then
'          lFlags = lFlags Or ILD_SELECTED
'      End If
'
'      If (bCut) Then
'        ' Draw dithered:
'        lColor = pTranslateColor(oCutDitherColour)
'        If (lColor = -1) Then lColor = pTranslateColor(vbWindowBackground)
'        ImageList_DrawEx _
'              lhIml, _
'              iImgIndex, _
'              hdc, _
'              xPixels, yPixels, 0, 0, _
'              CLR_NONE, lColor, _
'              lFlags
'      ElseIf (bDisabled) Then
'         If (ptrVB6ImageList <> 0) Then
'            Dim o As Object
'            On Error Resume Next
'            Set o = ObjectFromPtr(ptrVB6ImageList)
'            If Not (o Is Nothing) Then
'                hIcon = o.ListImages(iImgIndex + 1).ExtractIcon()
'            End If
'            On Error GoTo 0
'         Else
'            ' extract a copy of the icon:
'            hIcon = ImageList_GetIcon(hIml, iImgIndex, 0)
'         End If
'         If (hIcon <> 0) Then
'            ' Draw it disabled at x,y:
'            DrawState hdc, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED
'            ' Clear up the icon:
'            DestroyIcon hIcon
'         End If
'      Else
'         If (ptrVB6ImageList <> 0) Then
'             On Error Resume Next
'             Set o = ObjectFromPtr(ptrVB6ImageList)
'             If Not (o Is Nothing) Then
'                 o.ListImages(iImgIndex + 1).Draw hdc, xPixels * Screen.TwipsPerPixelX, yPixels * Screen.TwipsPerPixelY, lFlags
'             End If
'             On Error GoTo 0
'         Else
'            ' Standard draw:
'            ImageList_Draw _
'                lhIml, _
'                iImgIndex, _
'                hdc, _
'                xPixels, _
'                yPixels, _
'                lFlags
'         End If
'      End If
'   End If
'End Sub


'Public Property Let ImageList(vThis As Variant)
'    m_hIml = 0
'    m_ptrVb6ImageList = 0
'    If (VarType(vThis) = vbLong) Then
'        ' Assume a handle to an image list:
'        m_hIml = vThis
'    ElseIf (VarType(vThis) = vbObject) Then
'        ' Assume a VB image list:
'        On Error Resume Next
'        ' Get the image list initialised..
'        vThis.ListImages(1).Draw 0, 0, 0, 1
'        m_hIml = vThis.hImageList   '赋值其句柄
'        If (Err.Number = 0) Then
'            ' Check for VB6 image list:
'            If (TypeName(vThis) = "ImageList") Then
'                If (vThis.ListImages.Count <> ImageList_GetImageCount(m_hIml)) Then
'                    Dim o As Object
'                    Set o = vThis
'                    m_ptrVb6ImageList = ObjPtr(o)
'                End If
'            End If
'        Else
'            Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
'        End If
'        On Error GoTo 0
'    End If
'    If (m_hIml <> 0) Then
'        If (m_ptrVb6ImageList <> 0) Then
'            m_lIconSizeX = vThis.ImageWidth
'            m_lIconSizeY = vThis.ImageHeight
'        Else
'            Dim rc As RECT
'            ImageList_GetImageRect m_hIml, 0, rc
'            m_lIconSizeX = rc.Right - rc.Left
'            m_lIconSizeY = rc.Bottom - rc.Top
'        End If
'    End If
'End Property

Public Sub AddButton( _
      Optional ByVal sKey As String = "", _
      Optional ByVal sToolTipText As String = "", _
      Optional ByVal IconPicture As StdPicture, _
      Optional ByVal ePosition As ESBCButtonPositionConstants = esbcButtonPositionDefault, _
      Optional ByVal bCheck As Boolean = False, _
      Optional ByVal sCheckGroup As String = "", _
      Optional ByVal bVisible As Boolean = True, _
      Optional ByVal vKeyBefore As Variant _
   )
Dim lBtnIndex As Long
Dim iBtn As Long
Dim lStyle As Long

   If (m_eScrollType = esbcSizeGripper) Then
      ' No buttons on size grippers.
      Exit Sub
   End If

   ' Check if inserting a button:
   If Not (IsMissing(vKeyBefore)) Then
      ' Get button:
      lBtnIndex = ButtonIndex(vKeyBefore)
      If (lBtnIndex > 0) Then
         m_iButtonCount = m_iButtonCount + 1
         ReDim Preserve m_tButtons(1 To m_iButtonCount) As tButtonInfo
         ' Shift the array:
         For iBtn = m_iButtonCount To lBtnIndex + 1 Step -1
            LSet m_tButtons(iBtn) = m_tButtons(iBtn - 1)
         Next iBtn
      Else
         Exit Sub
      End If
   Else
      m_iButtonCount = m_iButtonCount + 1
      lBtnIndex = m_iButtonCount
      ReDim Preserve m_tButtons(1 To m_iButtonCount) As tButtonInfo
   End If
   
   ' Set the values:
   With m_tButtons(lBtnIndex)
      .sKey = sKey
      .sHelpText = sToolTipText
'      .lIconIndexUp = lIconIndexUp
'      .lIconIndexDown = lIconIndexDown
      If (ePosition = esbcButtonPositionDefault) Then
         If (m_eScrollType = esbcHorizontal) Then
            .ePosition = esbcButtonPositionLeftTop
         Else
            .ePosition = esbcButtonPositionRightBottom
         End If
      Else
         .ePosition = ePosition
      End If
      .bCheck = bCheck
      .sCheckGroup = sCheckGroup
      If (bCheck) Then
         m_iOptCount = m_iOptCount + 1
         If (m_iOptCount > 1) Then
            Load optButton(m_iOptCount - 1)
         End If
         Set .ctlThis = optButton(m_iOptCount - 1)
         
      Else
         m_iCmdCount = m_iCmdCount + 1
         If (m_iCmdCount > 1) Then
            Load cmdButton(m_iCmdCount - 1)
         End If
         Set .ctlThis = cmdButton(m_iCmdCount - 1)
      End If
      .ctlThis.Picture = IconPicture
      .ctlThis.Visible = bVisible
      .ctlThis.ToolTipText = sToolTipText
   End With
      
   pResizeButtons
   
End Sub

Public Property Get ButtonCount() As Long
   ButtonCount = m_iButtonCount
End Property

Public Property Get ButtonIndex(ByVal vKey As Variant) As Long
Dim lBtn As Long
Dim lIndex As Long
   If (IsNumeric(vKey)) Then
      lBtn = CLng(vKey)
      If (lBtn > 0) And (lBtn <= m_iButtonCount) Then
         lIndex = lBtn
      End If
   Else
      For lBtn = 1 To m_iButtonCount
         If (m_tButtons(lBtn).sKey = vKey) Then
            lIndex = lBtn
            Exit For
         End If
      Next lBtn
   End If
   If (lIndex > 0) Then
      ButtonIndex = lIndex
   Else
      Err.Raise 9, App.EXEName & ".ScrollButton", "下标越界"
   End If
   
End Property

Public Property Get ScrollType() As ESBCScrollTypes
   ScrollType = m_eScrollType
End Property
Public Property Let ScrollType(ByVal eType As ESBCScrollTypes)
   m_eScrollType = eType
   pCreateScrollControl
   PropertyChanged "ScrollType"
   Resize
End Property
Public Property Get XpStyleButtons() As Boolean
   XpStyleButtons = m_bXPStyleButtons
End Property
Public Property Let XpStyleButtons(ByVal bState As Boolean)
   m_bXPStyleButtons = bState
End Property

Public Property Get Visible() As Boolean
   Visible = UserControl.Extender.Visible
End Property
Public Property Let Visible(ByVal bState As Boolean)
   UserControl.Extender.Visible = bState
   Select Case m_eScrollType
   Case esbcVertical
      If (m_hWndParent <> 0) Then
         SetProp m_hWndParent, "vbalScrollButtons:VERT", Abs(bState)
      End If
   Case esbcHorizontal
      If (m_hWndParent <> 0) Then
         SetProp m_hWndParent, "vbalScrollButtons:HORZ", Abs(bState)
      End If
   End Select
End Property
Public Property Get SmallChange() As Long
   SmallChange = m_lSmallChange
End Property
Property Let SmallChange(ByVal lSmallChange As Long)
   m_lSmallChange = lSmallChange
End Property
Property Get ScrollEnabled() As Boolean
   Enabled = m_bScrollEnabled
End Property
Property Let ScrollEnabled(ByVal bEnabled As Boolean)
Dim lF As Long
        
   If (bEnabled) Then
      lF = ESB_ENABLE_BOTH
   Else
      lF = ESB_DISABLE_BOTH
   End If
   If (m_bNoFlatScrollBars) Then
      EnableScrollBar m_hWnd, SB_CTL, lF
   Else
      FlatSB_EnableScrollBar m_hWnd, SB_CTL, lF
   End If
    
End Property
Private Sub pGetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    
   tSI.fMask = fMask
   tSI.cbSize = LenB(tSI)
   
   If (m_bNoFlatScrollBars) Then
       GetScrollInfo m_hWnd, SB_CTL, tSI
   Else
       FlatSB_GetScrollInfo m_hWnd, SB_CTL, tSI
   End If

End Sub
Private Sub pLetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)
        
   tSI.fMask = fMask
   tSI.cbSize = LenB(tSI)
   If (m_bNoFlatScrollBars) Then
       SetScrollInfo m_hWnd, SB_CTL, tSI, True
   Else
       FlatSB_SetScrollInfo m_hWnd, SB_CTL, tSI, True
   End If
    
End Sub

Property Get Min() As Long
Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_RANGE
    Min = tSI.nMin
End Property

Property Get Max() As Long
Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_RANGE Or SIF_PAGE
    Max = tSI.nMax - tSI.nPage
End Property

Property Get Value() As Long
Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_POS
    Value = tSI.nPos
End Property

Property Get LargeChange() As Long
Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_PAGE
    LargeChange = tSI.nPage
End Property

Property Let Min(ByVal iMin As Long)
Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = Max + LargeChange
    pLetSI tSI, SIF_RANGE
End Property

Property Let Max(ByVal iMax As Long)
Dim tSI As SCROLLINFO
    tSI.nMax = iMax + LargeChange
    tSI.nMin = Min
    pLetSI tSI, SIF_RANGE
'    pRaiseEvent False
End Property

Property Let Value(ByVal iValue As Long)
Dim tSI As SCROLLINFO
Dim lPercent As Long
    If (iValue <> Value) Then
        tSI.nPos = iValue
        pLetSI tSI, SIF_POS
        If Me.ScrollType = esbcVertical Then
            Dim Hi As Long, k As Long
            Hi = (PAGEMARGIN + PubInfo.PaperHeight) * PubInfo.ZoomFactor
            k = Hi / VSTEP
            k = CInt(iValue / k) + 1
            UserControl.Extender.ToolTipText = "页码: " & CInt(k) & "      "           '第N页
        End If
        pRaiseEvent False
    End If
End Property

Property Let LargeChange(ByVal iLargeChange As Long)
Dim tSI As SCROLLINFO
Dim lCurMax As Long
Dim lCurLargeChange As Long
    
   pGetSI tSI, SIF_ALL
   tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
   tSI.nPage = iLargeChange
   pLetSI tSI, SIF_PAGE Or SIF_RANGE
End Property

Private Function pRaiseEvent(ByVal bScroll As Boolean)
Static s_lLastValue As Long
   If (Value <> s_lLastValue) Then
      If (bScroll) Then
         RaiseEvent Scroll
      Else
         RaiseEvent Change
      End If
      s_lLastValue = Value
   End If
End Function
Private Sub pCreateScrollControl()
Dim lStyle As Long
Dim lWidth As Long
Dim lHeight As Long
   
   If (m_hWndParent <> 0) Then
      pDestroyScrollControl
      lStyle = WS_CHILD Or WS_VISIBLE
      If (m_eScrollType = esbcHorizontal) Then
         lStyle = lStyle Or SBS_HORZ And Not SBS_VERT
         lWidth = UserControl.Width \ Screen.TwipsPerPixelX
         lHeight = CW_USEDEFAULT
      ElseIf (m_eScrollType = esbcVertical) Then
         lStyle = lStyle Or SBS_VERT And Not SBS_HORZ
         lHeight = UserControl.Height \ Screen.TwipsPerPixelY
         lWidth = CW_USEDEFAULT
      Else
         lStyle = lStyle Or SBS_SIZEBOX Or SBS_SIZEBOXBOTTOMRIGHTALIGN
      End If
      
      m_hWnd = CreateWindowEX(0, "SCROLLBAR", "", lStyle, 0, 0, lWidth, lHeight, UserControl.hwnd, 0, App.hInstance, ByVal 0&)
      If (m_hWnd <> 0) Then
         ShowScrollBar m_hWnd, SB_CTL, 1
         If (lStyle And SBS_SIZEBOX) <> SBS_SIZEBOX Then
            Min = 0
            Max = 255
            SmallChange = 1
            LargeChange = 32
         Else
            UserControl.BackColor = vbButtonFace
         End If
      End If
   End If
End Sub
Private Sub pDestroyScrollControl()
   If (m_hWnd <> 0) Then
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
   End If
End Sub

Private Sub pResizeButtons()
Dim lPos1 As Long
Dim lPos2 As Long
Dim lBtn As Long
Dim lExtent As Long
   
    On Error Resume Next
    
    If (m_eScrollType = esbcHorizontal) Then
       lExtent = GetSystemMetrics(SM_CYHSCROLL)
       lPos2 = UserControl.Width - lExtent * Screen.TwipsPerPixelX
    ElseIf (m_eScrollType = esbcVertical) Then
       lExtent = GetSystemMetrics(SM_CXVSCROLL)
       lPos2 = UserControl.Height - lExtent * Screen.TwipsPerPixelY
    Else
       Exit Sub
    End If
    
    For lBtn = 1 To m_iButtonCount
       With m_tButtons(lBtn)
          If (.ctlThis.Visible) Then
             If (.ePosition = esbcButtonPositionLeftTop) Then
                If (m_eScrollType = esbcHorizontal) Then
                   .ctlThis.Move lPos1, 0, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                   lPos1 = lPos1 + lExtent * Screen.TwipsPerPixelX
                Else
                   .ctlThis.Move 0, lPos1, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                   lPos1 = lPos1 + lExtent * Screen.TwipsPerPixelY
                End If
             Else
                If (m_eScrollType = esbcHorizontal) Then
                   .ctlThis.Move lPos2, 0, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                   lPos2 = lPos2 - lExtent * Screen.TwipsPerPixelX
                Else
                   .ctlThis.Move 0, lPos2, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                   lPos2 = lPos2 - lExtent * Screen.TwipsPerPixelY
                End If
             End If
          End If
       End With
    Next lBtn
    m_lPos1 = lPos1
    If (m_eScrollType = esbcHorizontal) Then
       m_lPos2 = lPos2 + lExtent * Screen.TwipsPerPixelX
    Else
       m_lPos2 = lPos2 + lExtent * Screen.TwipsPerPixelY
    End If
    Err.Clear
End Sub
Private Sub pResizeScroll()
Dim X As Long, Y As Long
Dim cx As Long, cy As Long

   If (m_hWnd <> 0) Then
      If (m_eScrollType = esbcHorizontal) Then
         Y = 0
         X = m_lPos1 \ Screen.TwipsPerPixelX
         cx = m_lPos2 \ Screen.TwipsPerPixelX - X
         cy = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
      ElseIf (m_eScrollType = esbcVertical) Then
         X = 0
         Y = m_lPos1 \ Screen.TwipsPerPixelY
         cy = m_lPos2 \ Screen.TwipsPerPixelY - Y
         cx = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
      Else
         X = 0
         Y = 0
         cx = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
         cy = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
      End If
      MoveWindow m_hWnd, X, Y, cx, cy, 1
   End If
End Sub

Public Sub Resize()
Dim tR As RECT
Dim bLeftScroll As Boolean
Dim bVert As Boolean
Dim bHorz As Boolean
Dim lStyle As Long
Dim lSize As Long

   GetClientRect m_hWndParent, tR
   ' Determine what other scroll bars on the parent:
   bVert = (GetProp(m_hWndParent, "vbalScrollButtons:VERT") <> 0)
   bHorz = (GetProp(m_hWndParent, "vbalScrollButtons:HORZ") <> 0)
   ' Determine if scroll bars are on the left or right:
   lStyle = GetWindowLong(m_hWndParent, GWL_EXSTYLE)
   If (lStyle And WS_EX_LEFTSCROLLBAR) Then
      bLeftScroll = True
   End If
   
   Select Case m_eScrollType
   Case esbcSizeGripper
      ' Only visible if both horz and vert.
      If (bVert) And (bHorz) And Not (bLeftScroll) Then
         tR.Left = tR.Right - GetSystemMetrics(SM_CXVSCROLL)
         tR.Top = tR.Bottom - GetSystemMetrics(SM_CYHSCROLL)
         MoveWindow UserControl.hwnd, tR.Left, tR.Top, (tR.Right - tR.Left), (tR.Bottom - tR.Top), 1
         UserControl_Resize
      End If
   Case esbcHorizontal
      ' We resize to the bottom of form.  Horizontal
      ' extent depends on whether Vertical scroll is
      ' visible
      lSize = GetSystemMetrics(SM_CYHSCROLL)
      tR.Top = tR.Bottom - lSize
      If (bVert) Then
         If (bLeftScroll) Then
            tR.Left = tR.Left + GetSystemMetrics(SM_CXVSCROLL)
         Else
            tR.Right = tR.Right - GetSystemMetrics(SM_CXVSCROLL)
         End If
      End If
      MoveWindow UserControl.hwnd, tR.Left, tR.Top, (tR.Right - tR.Left), (tR.Bottom - tR.Top), 1
      UserControl_Resize
      
   Case esbcVertical
      ' We resize to the right or left of form.  Horizontal
      ' extent depends on whether Vertical scroll is
      ' visible
      lSize = GetSystemMetrics(SM_CXVSCROLL)
      If (bLeftScroll) Then
         tR.Right = tR.Left + lSize
      Else
         tR.Left = tR.Right - lSize
      End If
      If (bHorz) Then
         tR.Bottom = tR.Bottom - GetSystemMetrics(SM_CYHSCROLL)
      End If
      MoveWindow UserControl.hwnd, tR.Left, tR.Top, (tR.Right - tR.Left), (tR.Bottom - tR.Top), 1
      UserControl_Resize
   End Select
End Sub

Private Sub pDrawButton(tDis As DRAWITEMSTRUCT)
Dim hBr As Long
Dim lState As Long
Dim bPushed As Boolean
Dim bDisabled As Boolean
Dim bChecked As Boolean
Dim iBtn As Long
Dim iBtnIndex As Long
Dim lSize As Long
Dim X As Long, Y As Long
Dim bXpStyle As Boolean
Dim hTheme As Long
Dim hr As Long

   lState = SendMessageLong(tDis.hwndItem, BM_GETSTATE, 0, 0)
   'Debug.Print lState
   bPushed = ((lState And BST_CHECKED) = BST_CHECKED) Or ((lState And BST_PUSHED) = BST_PUSHED)
      
   For iBtn = 1 To m_iButtonCount
      If (m_tButtons(iBtn).ctlThis.hwnd = tDis.hwndItem) Then
         iBtnIndex = iBtn
         bChecked = (m_tButtons(iBtn).ctlThis.Value = True)
         bPushed = bPushed Or bChecked
         bDisabled = Not (m_tButtons(iBtnIndex).ctlThis.Enabled)
         Exit For
      End If
   Next iBtn
      
   If (m_bXPStyleButtons) Then
      On Error Resume Next
      hTheme = OpenThemeData(hwnd, StrPtr("Button"))
      If (Err.Number <> 0) Or (hTheme = 0) Then
         bXpStyle = False
      Else
         bXpStyle = True
      End If
   End If
   
   If bChecked Then
      hBr = GetSysColorBrush(vb3DHighlight And &H1F&)
   Else
      hBr = GetSysColorBrush(vbButtonFace And &H1F&)
   End If
   FillRect tDis.hdc, tDis.rcItem, hBr
   DeleteObject hBr
   
   If (bXpStyle) Then
      If bDisabled Then
         hr = DrawThemeBackground(hTheme, tDis.hdc, 1, _
                4, tDis.rcItem, tDis.rcItem)
      ElseIf bChecked Or bPushed Then
         hr = DrawThemeBackground(hTheme, tDis.hdc, 1, _
                3, tDis.rcItem, tDis.rcItem)
      Else
         hr = DrawThemeBackground(hTheme, tDis.hdc, 1, _
                1, tDis.rcItem, tDis.rcItem)
      End If
   End If
   
   If (iBtnIndex > 0) Then
      If (m_eScrollType = esbcHorizontal) Then
         lSize = GetSystemMetrics(SM_CYHSCROLL) - 4
      Else
         lSize = GetSystemMetrics(SM_CXVSCROLL) - 4
      End If
      X = 2 + (lSize - m_lIconSizeX) \ 2
      Y = X
      If (bPushed) Then
         X = X + 1
         Y = Y + 1
'         pDrawImage m_ptrVb6ImageList, m_hIml, m_tButtons(iBtnIndex).lIconIndexDown, tDis.hdc, x, y, m_lIconSizeX, m_lIconSizeY, , , bDisabled
      Else
'         pDrawImage m_ptrVb6ImageList, m_hIml, m_tButtons(iBtnIndex).lIconIndexUp, tDis.hdc, x, y, m_lIconSizeX, m_lIconSizeY, , , bDisabled
      End If
   End If
   
   If (bXpStyle) Then
   
   Else
      If (bPushed) Then
         DrawEdge tDis.hdc, tDis.rcItem, BDR_SUNKENOUTER, BF_RECT
      Else
         DrawEdge tDis.hdc, tDis.rcItem, BDR_RAISEDINNER Or BDR_RAISEDOUTER, BF_RECT
      End If
   End If
   
   If (hTheme) Then
      CloseThemeData hTheme
   End If
   Err.Clear
   
End Sub

Private Sub optButton_Click(Index As Integer)
Dim iB As Long
Dim lBtnIndex As Long
   For iB = 1 To m_iButtonCount
      If (m_tButtons(iB).ctlThis Is optButton(Index)) Then
         lBtnIndex = iB
         Exit For
      End If
   Next iB
   If (lBtnIndex > 0) Then
      RaiseEvent ButtonClick(lBtnIndex)
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim iB As Long
Dim lBtnIndex As Long
   For iB = 1 To m_iButtonCount
      If (m_tButtons(iB).ctlThis Is cmdButton(Index)) Then
         lBtnIndex = iB
         Exit For
      End If
   Next iB
   If (lBtnIndex > 0) Then
      RaiseEvent ButtonClick(lBtnIndex)
   End If

End Sub

Private Sub UserControl_Initialize()
   m_bNoFlatScrollBars = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If (UserControl.Ambient.UserMode) Then
        m_hWndControl = UserControl.hwnd
        Subclass1.hwnd = UserControl.hwnd
        Subclass1.Messages(WM_DRAWITEM) = True
        Subclass1.Messages(WM_CTLCOLORSCROLLBAR) = True
        Subclass1.Messages(WM_VSCROLL) = True
        Subclass1.Messages(WM_HSCROLL) = True
        m_hWndParent = UserControl.Extender.Container.hwnd
        UserControl.BorderStyle() = 0
    End If
    ScrollType = PropBag.ReadProperty("ScrollType", esbcHorizontal)
    Visible = PropBag.ReadProperty("Visible", True)
End Sub

Private Sub UserControl_Resize()
    If (m_hWndControl <> 0) Then
        pResizeButtons
        pResizeScroll
    End If
End Sub

Private Sub UserControl_Terminate()
    If (m_hWndControl <> 0) Then
        pDestroyScrollControl
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "ScrollType", ScrollType, esbcHorizontal
   PropBag.WriteProperty "Visible", Visible, True
End Sub

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    Dim tDis As DRAWITEMSTRUCT
    Dim lBar As Long
    Dim lScrollcode As Long
    Dim lV As Long, lSC As Long
    Dim tSI As SCROLLINFO

    Select Case Msg
    Case WM_DRAWITEM
        CopyMemory tDis, ByVal lParam, Len(tDis)
        pDrawButton tDis
        Result = 1
    Case WM_SIZE
        Resize
    Case WM_CTLCOLORSCROLLBAR
        If (wParam = m_hWndControl) Then
           Result = GetSysColorBrush(SystemColorConstants.vbWindowBackground And &H1F)
        End If
    Case WM_VSCROLL, WM_HSCROLL
        lBar = SB_CTL
        lScrollcode = (wParam And &HFFFF&)
        Select Case lScrollcode
        Case SB_THUMBTRACK
           ' Is vertical/horizontal?
           pGetSI tSI, SIF_TRACKPOS
           Value = tSI.nTrackPos
           pRaiseEvent True
        
        Case SB_LEFT, SB_BOTTOM
           Value = Min
           pRaiseEvent False
        
        Case SB_RIGHT, SB_TOP
           Value = Max
           pRaiseEvent False
        
        Case SB_LINELEFT, SB_LINEUP
           'Debug.Print "Line"
           lV = Value
           lSC = m_lSmallChange
           If (lV - lSC < Min) Then
              Value = Min
           Else
              Value = lV - lSC
           End If
           pRaiseEvent False
        
        Case SB_LINERIGHT, SB_LINEDOWN
            'Debug.Print "Line"
           lV = Value
           lSC = m_lSmallChange
           If (lV + lSC > Max) Then
              Value = Max
           Else
              Value = lV + lSC
           End If
           pRaiseEvent False
        
        Case SB_PAGELEFT, SB_PAGEUP
           Value = Value - LargeChange
           pRaiseEvent False
        
        Case SB_PAGERIGHT, SB_PAGEDOWN
           Value = Value + LargeChange
           pRaiseEvent False
        Case SB_ENDSCROLL
           pRaiseEvent False
        End Select
    End Select
End Sub



