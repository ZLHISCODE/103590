VERSION 5.00
Begin VB.UserControl Ruler 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
End
Attribute VB_Name = "Ruler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =========================================================================================
' Filename:       jdHRuler.ctl
' Author:         Jens Duczmal
' Date:           27 April 2000
' Version :       0.9 Beta
'
' Dependencies:
'     at Designtime :   cVBAImageList.cls
'                       jdRuler.Res
'
'     at Runtime  :     NONE
'
' Description:
' MS-Word 2000-Style Ruler-Control
' Allows handling of Different Scales.
' Has Left/Right-Margins, some Indents and Tab-Stop facilities.
' (TAB STOPS NOT IMPLEMENTED YET)
' =========================================================================================
' Properties            Access   Default Values / Description
' -----------------------------------------------------------------------------------------
' FirstLineIndent       G L      {0}
'                                Position of FirstLineIndent in current Ruler-Scale
'
' hwndBound             G L      {0}
'                                Handle to the "bound" Textbox/RTF-Control.
'                                Currently used to draw the dotted line while moving Indents
'
' LeftIndent            G L      {0}
'                                Position of LeftIndent in current Ruler-Scale
'
' LeftMargin            G L      {2}
'                                Position of LeftMargin in current Ruler-Scale
'
' RightIndent           G L      {0}
'                                Position of RightIndent in current Ruler-Scale
'
' RightMargin           G L      {0}
'                                Position of RightMargin in current Ruler-Scale
'
' RulerScale            G L      {7}
'                                Scale Mode of the Ruler (eRulerScale)
'                                       Pixels = 3
'                                       Inches = 5
'                                       Millimeters = 6
'                                       Centimeters = 7
'
' =========================================================================================
' EVENTS
' -----------------------------------------------------------------------------------------
' Event LeftMarginChanged()
' Event RightMarginChanged()
' Event FirstLineIndentChanged()
' Event LeftIndentChanged()
' Event RightIndentChanged()
' =========================================================================================

'Enums and Types
'---------------------------------
Public Enum eRulerScale
   Pixels = 3
   Inches = 5
   Millimeters = 6
   Centimeters = 7
End Enum

Public Enum eAlign
   tLeft = 0
   tRight = 2
   tCentered = 1
End Enum

Public Enum ePermission
   None = 0
   ByCode = 1
   ByUser = 2
End Enum

'Property variables
'---------------------------------
Private m_lLeftMargin         As Long     'Long-Values will store the Margins and
Private m_lRightMargin        As Long     'Indents as Pixels. Used for the API
Private m_lHangingIndent      As Long     'Decided for this solution instead
Private m_lLeftIndent         As Long     'of calculating again and again
Private m_lRightIndent        As Long
Private m_lFirstLineIndent    As Long
Private m_iRulerScale         As Integer  'The selected Ruler-Scale (eRulerScale)
Private m_lRulerLength        As Long

Private m_sglQuantise         As Single

'Working Variables
'---------------------------------
Private m_cIL                 As cImageList 'Handle to VBALImageList

Private mRect                 As RECT     'Rect of the UserControl
Private mHwnd                 As Long     'Hwnd of UserControl
Private mHdc                  As Long     'Hdc of UserControl
Private mTp                   As POINTAPI 'Used for Drawing of Ruler

Private m_bInDev              As Boolean  'In DesignMode ?

Private m_rLeftIndent         As RECT     'Rect-Structures for all Indents
Private m_rRightIndent       As RECT     'Decided to use PtInRect to check
Private m_rFirstLineIndent    As RECT     'if mouse is inside.
Private m_rHangingIndent      As RECT     'Code is cleaner I assume

Private m_sglRulerStep        As Single   'All Used for drawing the Ruler
Private m_lRulerStep          As Long     'in different scales.
Private m_iStepLarge          As Integer  'See pSetRulerScale for explanation
Private m_iStepHalf           As Integer
Private m_iStepSmall          As Integer

Private m_hwndBound           As Long     'Handle to the 'Bound' Control
                                          'in which the dotted Line will be drawn
Private m_bytAllowTabs        As Byte
Private m_bytAllowIndents     As Byte
Private m_bytAllowMargins     As Byte
Private m_lFontHeight         As Long

Private m_bytMoving           As Byte     'Store which Indent/Margin is currently moving
Private m_iTabMoving          As Integer
Private m_iTabCount           As Integer
Private m_arrTabStop()        As Long
Private m_arrTabAlign()       As Byte

'Default Constants
'---------------------------------
Private Const cdefLeftMargin = 1134          'Some Defaults for Margins/Indents.
Private Const cdefRightMargin = 1134         'Dims as per defRulerScale
Private Const cdefRulerScale = 7          'so actually 2 cm (sorry, I'm German)
Private Const cdefRulerLength = 10206

Private Const cdefLeftIndent = 0
Private Const cdefRightIndent = 0
Private Const cdefFirstLineIndent = 0
Private Const cdefHangingIndent = 0

Private Const cdefAllowMargins = 2
Private Const cdefAllowTabs = 2
Private Const cdefAllowIndents = 2

'Working Constants
'---------------------------------
Private Const cMinMaxHeight = 390      'MinMax-Height in Pixels of UserControl

Private Const cLeftMargin = 1          'Constants to be stored in m_bytMoving
Private Const cRightMargin = 2         'to check which Margin/Indent is currently moving
Private Const cFirstLineIndent = 3
Private Const cHangingIndent = 4
Private Const cLeftIndent = 5
Private Const cRightIndent = 6
Private Const cTabStop = 7

Private Const IconX = 16               'Icon Dims
Private Const IconY = 16

'Events
'---------------------------------
Event IndentChanged(LeftIndent As Long, FirstLineIndent As Long, RightIndent As Long)
Event MarginChanged(LeftMargin As Long, RightMargin As Long)
Event TabStopChanged(TabCount As Integer, TabPos() As Long, TabAlign() As Byte)
Event DblClick()

Public Property Let Quantise(Value As Single)
   If Value = 0 Then
    m_sglQuantise = UserControl.ScaleX(m_lRulerStep, vbPixels, vbTwips)

      'If m_iRulerScale = eRulerScale.Inches Then
      '  m_sglQuantise = UserControl.ScaleX(0.1, vbInches, vbTwips)
      'ElseIf eRulerScale.Centimeters Then
      '   m_sglQuantise = UserControl.ScaleX(0.1, vbCentimeters, vbTwips)
      'ElseIf m_iRulerScale = eRulerScale.Millimeters Then
      '  m_sglQuantise = UserControl.ScaleX(1, vbMillimeters, vbTwips)
      'ElseIf eRulerScale.Pixels Then
      '   m_sglQuantise = UserControl.ScaleX(1, vbPixels, vbTwips)
      'Else
      '   m_sglQuantise = Value
      'End If
   Else
      m_sglQuantise = Value
   End If
   PropertyChanged "Quantise"
End Property

Public Property Get Quantise() As Single
   Quantise = m_sglQuantise
End Property

Public Property Set Font(sFont As StdFont)
   Set UserControl.Font = sFont
   m_lFontHeight = CLng(UserControl.TextHeight("8"))
   PropertyChanged "Font"
   pDraw
End Property

Public Property Get Font() As StdFont
   Set Font = UserControl.Font
End Property
Public Property Let AllowTabs(State As ePermission)
   m_bytAllowTabs = State
   PropertyChanged "AllowTabs"
   pDraw
End Property

Public Property Get AllowTabs() As ePermission
   AllowTabs = m_bytAllowTabs
End Property

Public Property Let AllowIndents(State As ePermission)
   m_bytAllowIndents = State
   PropertyChanged "AllowIndents"
   pDraw
End Property

Public Property Get AllowIndents() As ePermission
   AllowIndents = m_bytAllowIndents
End Property

Public Property Let AllowMargins(State As ePermission)
   m_bytAllowMargins = State
   PropertyChanged "AllowMargins"
   pDraw
End Property
Public Property Get AllowMargins() As ePermission
   AllowMargins = m_bytAllowMargins
End Property

Public Property Let FirstLineIndent(nPos As Long)
   m_lFirstLineIndent = UserControl.ScaleX(nPos, vbTwips, vbPixels)
   PropertyChanged "FirstLineIndent"
   pDraw
End Property

Public Property Get FirstLineIndent() As Long
   FirstLineIndent = UserControl.ScaleX(m_lFirstLineIndent, vbPixels, vbTwips)
End Property

Public Property Let hwndBound(hwnd As Long)
   m_hwndBound = hwnd
End Property

Public Property Let LeftIndent(nPos As Long)
   m_lLeftIndent = UserControl.ScaleX(nPos, vbTwips, vbPixels)
   m_lHangingIndent = m_lLeftIndent
   PropertyChanged "LeftIndent"
   pDraw
End Property

Public Property Get LeftIndent() As Long
   LeftIndent = UserControl.ScaleX(m_lLeftIndent, vbPixels, vbTwips)
End Property

Public Property Let LeftMargin(nPos As Long)
   m_lLeftMargin = UserControl.ScaleX(nPos, vbTwips, vbPixels)

   PropertyChanged "LeftMargin"
   pDraw
End Property

Public Property Get LeftMargin() As Long
   LeftMargin = UserControl.ScaleX(m_lLeftMargin, vbPixels, vbTwips)
End Property

Public Property Let RightIndent(nPos As Long)
   m_lRightIndent = UserControl.ScaleX(nPos, vbTwips, vbPixels)
   PropertyChanged "RightIndent"
   pDraw
End Property

Public Property Get RightIndent() As Long
   RightIndent = UserControl.ScaleX(m_lRightIndent, vbPixels, vbTwips)
End Property

Public Property Let RightMargin(nPos As Long)
   m_lRightMargin = UserControl.ScaleX(nPos, vbTwips, vbPixels)
   PropertyChanged "RightMargin"
   pDraw
End Property

Public Property Get RightMargin() As Long
   RightMargin = UserControl.ScaleX(m_lRightMargin, vbPixels, vbTwips)
End Property

Public Property Let RulerLength(nLength As Long)
   m_lRulerLength = UserControl.ScaleX(nLength, vbTwips, vbPixels)
   UserControl.Width = nLength
   PropertyChanged "RulerLength"
   pDraw
End Property

Public Property Get RulerLength() As Long
   RulerLength = UserControl.ScaleX(m_lRulerLength, vbPixels, vbTwips)
End Property

Public Property Let RulerScale(iScale As eRulerScale)
   m_iRulerScale = iScale
   PropertyChanged "RulerScale"
   pSetRulerScale
   pDraw
End Property

Public Property Get RulerScale() As eRulerScale
   RulerScale = m_iRulerScale
End Property

Public Sub SetTabs(iCount As Integer, TabStop() As Long, TabAlign() As Byte)
    If m_bytAllowTabs = 0 Then Exit Sub

    ReDim m_arrTabStop(0)
    ReDim m_arrTabAlign(0)

    m_iTabCount = iCount
    m_arrTabStop = TabStop
    m_arrTabAlign = TabAlign
    
    pDraw
    
    RaiseEvent TabStopChanged(iCount, TabStop, TabAlign)
End Sub

Private Sub SortTabs()
Dim arrPos() As Long
ReDim arrPos(UBound(m_arrTabStop))

arrPos = m_arrTabStop
QuickSort m_arrTabStop, 0, UBound(m_arrTabStop) - 1

Dim iCnt As Integer
Dim iTabCnt As Integer

    For iCnt = 0 To UBound(m_arrTabStop) - 1
        For iTabCnt = 0 To UBound(arrPos) - 1
            If arrPos(iTabCnt) = m_arrTabStop(iCnt) Then
                m_arrTabAlign(iCnt) = m_arrTabAlign(iTabCnt)
            End If
        Next
    Next

pDraw
End Sub

Private Sub QuickSort(vArray As Variant, l As Integer, r As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim X
    Dim Y
    i = l
    j = r
    X = vArray((l + r) / 2)
    While (i <= j)
        While (vArray(i) < X And i < r)
            i = i + 1
        Wend
        While (X < vArray(j) And j > l)
            j = j - 1
        Wend
        If (i <= j) Then
            Y = vArray(i)
            vArray(i) = vArray(j)
            vArray(j) = Y
            i = i + 1
            j = j - 1
        End If
    Wend
    If (l < j) Then QuickSort vArray, l, j
    If (i < r) Then QuickSort vArray, i, r
End Sub

Private Sub AddTab(Position As Long, Optional Align As eAlign = eAlign.tLeft)
Dim iCnt As Long
Dim iNewTab As Integer

ReDim Preserve m_arrTabStop(UBound(m_arrTabStop) + 1)
ReDim Preserve m_arrTabAlign(UBound(m_arrTabAlign) + 1)
m_iTabCount = m_iTabCount + 1
On Error Resume Next
iNewTab = -1

    For iCnt = 0 To UBound(m_arrTabStop) - 1
        If m_arrTabStop(iCnt) > Position Then
            iNewTab = iCnt
            Exit For
        End If
    Next iCnt
    
    If iNewTab = -1 Then
        m_arrTabStop(UBound(m_arrTabStop) - 1) = Position
        m_arrTabAlign(UBound(m_arrTabAlign) - 1) = Align
    Else
        For iCnt = UBound(m_arrTabStop) - 1 To iNewTab Step -1
            m_arrTabStop(iCnt) = m_arrTabStop(iCnt - 1)
            m_arrTabAlign(iCnt) = m_arrTabAlign(iCnt - 1)
        Next
        m_arrTabStop(iNewTab) = Position
        m_arrTabAlign(iNewTab) = Align
   End If
   
    pDraw
End Sub

Private Sub RemoveTab(Index As Integer)
Dim iCnt As Long
m_iTabCount = m_iTabCount - 1
On Error Resume Next

    For iCnt = Index To UBound(m_arrTabStop) - 1
        m_arrTabStop(iCnt) = m_arrTabStop(iCnt + 1)
        m_arrTabAlign(iCnt) = m_arrTabAlign(iCnt + 1)
    Next iCnt

ReDim Preserve m_arrTabStop(UBound(m_arrTabStop) - 1)
ReDim Preserve m_arrTabAlign(UBound(m_arrTabAlign) - 1)
If m_arrTabStop(0) = 0 Then m_iTabMoving = -1
    pDraw
End Sub

Private Sub pSetRulerScale()
'Prepare to draw Ruler in selected Scalemode
'Explanation follows for Pixels and CM
   Select Case m_iRulerScale
      Case 3
         m_sglRulerStep = 6      'Every 6 Pixels, something need to be drawn
         m_iStepSmall = 1        'The small Step is the small Line. Every 1 x 6 Pixels
         m_iStepHalf = 0         'The half-sized line does not exist with Pixels. 0 x 6 = 0
         m_iStepLarge = 6        'LargeStep draws the Number itself. Every 6 x 6 = 36 Pixels
      Case 5
         m_sglRulerStep = 0.125  'Inches have much smaller steps compared with cm
         m_iStepSmall = 1
         m_iStepHalf = 4
         m_iStepLarge = 8
      Case 6
         m_sglRulerStep = 2.5
         m_iStepSmall = 1
         m_iStepHalf = 4
         m_iStepLarge = 8
      Case 7
         m_sglRulerStep = 0.25   'With CM, we draw something every 0.25 cm
         m_iStepSmall = 1        'means every 1 * 0.25 = 0.25 a small line
         m_iStepHalf = 2         'or every 0.5 a half-sized-line
         m_iStepLarge = 4        'or every 4 * 0.25 = 1 cm the Number
      Case Else
         Exit Sub
   End Select
   
   'Finally we must calculate the Small-Stepping in Pixels. Used later in For-Next-Loop
   m_lRulerStep = CLng(UserControl.ScaleX(m_sglRulerStep, m_iRulerScale, vbPixels))
End Sub

Private Sub CalcIconPositions()
'We will calculate some RECT-Structures here for the Indents.
'On Usercontrol_MouseMove we have to move the Pics for the Indents
'To allow quick check and clear code I decided to use
'PtInRect-API to check whether Cursors is within this area
'So I need Rect-Structures.

'Note that the Icons have 16x16 pixels but the
'Pictures byself got only appx. 8 x 9 pixels.
If m_bytAllowIndents = ePermission.ByUser Then
   m_rHangingIndent.Left = mRect.Left + m_lLeftMargin + m_lHangingIndent - (IconX / 4)
   m_rHangingIndent.Top = mRect.Bottom - 8
   m_rHangingIndent.Right = m_rHangingIndent.Left + IconX
   m_rHangingIndent.Bottom = mRect.Bottom
   
   m_rLeftIndent.Left = mRect.Left + m_lLeftMargin + m_lLeftIndent - (IconX / 4)
   m_rLeftIndent.Top = mRect.Bottom - 9 + 6
   m_rLeftIndent.Right = m_rLeftIndent.Left + IconX
   m_rLeftIndent.Bottom = m_rLeftIndent.Top + 9
   
   m_rFirstLineIndent.Left = mRect.Left + m_lLeftMargin + m_lLeftIndent + m_lFirstLineIndent - (IconX / 4)
   m_rFirstLineIndent.Top = mRect.Bottom - 9 - 8
   m_rFirstLineIndent.Right = m_rFirstLineIndent.Left + IconX
   m_rFirstLineIndent.Bottom = mRect.Bottom - 9
   
   m_rRightIndent.Left = mRect.Right - m_lRightMargin - m_lRightIndent - (IconX / 4)
   m_rRightIndent.Top = mRect.Bottom - 9
   m_rRightIndent.Right = m_rRightIndent.Left + IconX
   m_rRightIndent.Bottom = mRect.Bottom
End If
End Sub

Private Sub pDraw()
'Drawing of Ruler to be done all here

Dim lBrush As Long         'Handle for the Brush (FillColor) we create
Dim lBrushOld As Long      'Handle for OriginalBrush
Dim lPen As Long           'Handle for the Pen (LineColor) we create
Dim lpenOld As Long        'Handle for OriginalPen
Dim rText As RECT

Dim sglCount As Single
Dim lngPos As Long         'Current Position to draw in the Ruler
Dim sglPos As Single
Dim lngLength As Long      'Length of the Ruler
Dim bytStepCount As Byte   'Counter from 1 to 4 in order to determine what to draw
Dim lCount As Long         'Counter for cm. Needed to draw the Number into the ruler
Static lngMoveStart    As Long

   'Clear Control first
   UserControl.Cls
   
   'Now save the Hwnd / Hdc-Properties.
   mHwnd = UserControl.hwnd
   mHdc = UserControl.hdc

   'Get Dimensions of UserControl
   GetClientRect mHwnd, mRect
   
   'Increase as 6 Pixels around (Ruler is smaller than Control)
   InflateRect mRect, 0, -6
   
   'Create a White Brush (FillColor) and save the Original one
   lBrush = CreateSolidBrush(&HFFFFFF)
   lBrushOld = SelectObject(mHdc, lBrush)
   
   'Same with Pen in White
   lPen = CreatePen(0, 1, &HFFFFFF)
   lpenOld = SelectObject(mHdc, lPen)
   If m_bytAllowMargins = ePermission.None Then m_lLeftMargin = 0: m_lRightMargin = 0
   
   'Draw White Rectangle less Left/Right margins if any. Plus/Minus 2 Pixels for optical matters.
   Rectangle mHdc, m_lLeftMargin, mRect.Top, mRect.Right - m_lRightMargin, mRect.Bottom

   'Now clean up the Brush -> Select Original and delete the new
   'This order is quite Important, otherwise your Ressources will be dramatically reduced
   SelectObject mHdc, lBrushOld
   DeleteObject lBrush
   
   'Pen must be deleted as well
   SelectObject mHdc, lpenOld
   DeleteObject lPen
   
   
   
   'If any LeftMargin so draw now in DarkGrey
   If m_lLeftMargin > 0 Then
      lBrush = CreateSolidBrush(&H808080) 'Dark Grey Brush
      lBrushOld = SelectObject(mHdc, lBrush)
      lPen = CreatePen(0, 1, &H808080)
      lpenOld = SelectObject(mHdc, lPen)
      'Draw left margin (darkgrey)
      'but leave 2 Pixels space (will be LightGrey to match Optic with MS-Word)
      Rectangle mHdc, mRect.Left, mRect.Top, m_lLeftMargin - 2, mRect.Bottom
      SelectObject mHdc, lBrushOld
      DeleteObject lBrush
      SelectObject mHdc, lpenOld
      DeleteObject lPen
   End If
   
   'Do same with RightMargin
   If m_lRightMargin > 0 Then
      lBrush = CreateSolidBrush(&H808080) 'Dark Grey Brush
      lBrushOld = SelectObject(mHdc, lBrush)
      lPen = CreatePen(0, 1, &H808080)
      lpenOld = SelectObject(mHdc, lPen)
      'Draw Right Margin (Dark Grey)
      'but leave 2 Pixels space (will be LightGrey to match Optic with MS-Word)
      Rectangle mHdc, mRect.Right - m_lRightMargin + 2, mRect.Top, mRect.Right, mRect.Bottom
      SelectObject mHdc, lBrushOld
      DeleteObject lBrush
      SelectObject mHdc, lpenOld
      DeleteObject lPen
   End If
   
   'We are now going to draw the Ruler-Scale.
   
   'First, reset some Counter-Variables (Remember : Different RulerScales allowed)
   bytStepCount = 1
   sglCount = 0

   'We will draw from the White-Area to the Right first
   'Left of Usercontrol + LeftMargin if any
   
   'For-Next-Loop will loop through SGL-Values (depends on RulerScale)
   'Drawing to be done, of course in Pixels. Must be handled like this
   'in order to avoid offset after 2-3 Inches
   
   For sglPos = m_lLeftMargin + m_lRulerStep To m_lRulerLength Step m_lRulerStep
   
      'We got now the Position of next Scale-Part but in RulerScale
      'So recalculate in Pixels
      lngPos = sglPos
      'sglCount will store the Number to be drawn later as text
      'Could allow drawing of even numbers as well.
      sglCount = sglCount + m_sglRulerStep
      'Now decide what shall be drawn and just do it.

      Select Case bytStepCount
         Case m_iStepLarge
            rText.Top = mRect.Top
            rText.Bottom = mRect.Bottom
            rText.Left = mRect.Left + lngPos - (m_lRulerStep * 2)
            rText.Right = mRect.Left + lngPos + (m_lRulerStep * 2)
            DrawText mHdc, CStr(sglCount), Len(CStr(sglCount)), rText, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
         Case m_iStepHalf
            MoveToEx mHdc, mRect.Left + lngPos, mRect.Top + 4, mTp
            LineTo mHdc, mRect.Left + lngPos, mRect.Top + 10
         Case Else
            MoveToEx mHdc, mRect.Left + lngPos, mRect.Top + 6, mTp
            LineTo mHdc, mRect.Left + lngPos, mRect.Top + 8
      End Select
      
      'Increase Counter for later decision whether SmallLine,HalfLine or Number
      'shall be drawn. Differs from Scale to Scale
      bytStepCount = bytStepCount + 1
      If bytStepCount > m_iStepLarge Then bytStepCount = 1
   Next sglPos
   
   
   'Now we will draw the Part for the LeftMargin.
   'NOTICE : Scale-Numbers in Descending Order !!!
   If m_lLeftMargin > 0 Then
      bytStepCount = 1
      sglCount = 0
      'Make it Descending !!!!!
      For sglPos = m_lLeftMargin - m_lRulerStep To 0 Step -m_lRulerStep
         'lngPos = CLng(UserControl.ScaleX(sglPos, m_iRulerScale, vbPixels))
         lngPos = sglPos
         sglCount = sglCount + m_sglRulerStep
         Select Case bytStepCount
            Case m_iStepHalf
               MoveToEx mHdc, mRect.Left + lngPos, mRect.Top + 4, mTp
               LineTo mHdc, mRect.Left + lngPos, mRect.Top + 10
            Case m_iStepLarge
               rText.Top = mRect.Top
               rText.Bottom = mRect.Bottom
               rText.Left = mRect.Left + lngPos - (m_lRulerStep * 2)
               rText.Right = mRect.Left + lngPos + (m_lRulerStep * 2)
               DrawText mHdc, CStr(sglCount), Len(CStr(sglCount)), rText, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
            Case Else
               MoveToEx mHdc, mRect.Left + lngPos, mRect.Top + 6, mTp
               LineTo mHdc, mRect.Left + lngPos, mRect.Top + 8
         End Select
         bytStepCount = bytStepCount + 1
         If bytStepCount > m_iStepLarge Then bytStepCount = 1
      Next sglPos
   End If
   
   'Last thing missing are the icons of the Indents / Tabstops
   'Here is a good place to calculate the Rect-Structures
   CalcIconPositions

   'Start drawing the 'Shadows' of Indents/Tabs if something is moving
   Select Case m_bytMoving
      Case 0
         lngMoveStart = 0
         
      Case cFirstLineIndent
         If lngMoveStart = 0 Then lngMoveStart = m_rFirstLineIndent.Left
         m_cIL.DrawImage 6, mHdc, lngMoveStart, m_rFirstLineIndent.Top - 7
      Case cHangingIndent
         If lngMoveStart = 0 Then lngMoveStart = m_rLeftIndent.Left
         m_cIL.DrawImage 4, mHdc, lngMoveStart, m_rHangingIndent.Top - 7
         m_cIL.DrawImage 2, mHdc, lngMoveStart, m_rLeftIndent.Top - 7
      Case cLeftIndent
         If lngMoveStart = 0 Then lngMoveStart = m_rLeftIndent.Left
         m_cIL.DrawImage 2, mHdc, lngMoveStart, m_rLeftIndent.Top - 7
         m_cIL.DrawImage 4, mHdc, lngMoveStart, m_rHangingIndent.Top - 7
         m_cIL.DrawImage 6, mHdc, lngMoveStart + m_lFirstLineIndent, m_rFirstLineIndent.Top - 7
      Case cRightIndent
         If lngMoveStart = 0 Then lngMoveStart = m_rRightIndent.Left
         m_cIL.DrawImage 4, mHdc, lngMoveStart, m_rRightIndent.Top - 7
      Case cTabStop
         If lngMoveStart = 0 Then lngMoveStart = m_lLeftMargin + UserControl.ScaleX(m_arrTabStop(m_iTabMoving), vbTwips, vbPixels)
         Select Case m_arrTabAlign(m_iTabMoving)
            Case eAlign.tLeft
               m_cIL.DrawImage 8, mHdc, lngMoveStart, mRect.Bottom - IconY
            Case eAlign.tRight
               m_cIL.DrawImage 10, mHdc, lngMoveStart, mRect.Bottom - IconY
            Case eAlign.tCentered
               m_cIL.DrawImage 12, mHdc, lngMoveStart, mRect.Bottom - IconY
         End Select
   End Select
   
   'Now draw the Images. Top-7 because Icon = 16 px but Picture only 9 px high
   If m_bytAllowIndents <> ePermission.None Then
      m_cIL.DrawImage 3, mHdc, m_rHangingIndent.Left, m_rHangingIndent.Top - 7
      m_cIL.DrawImage 1, mHdc, m_rLeftIndent.Left, m_rLeftIndent.Top - 7
      m_cIL.DrawImage 5, mHdc, m_rFirstLineIndent.Left, m_rFirstLineIndent.Top - 7
      m_cIL.DrawImage 3, mHdc, m_rRightIndent.Left, m_rRightIndent.Top - 7
   End If
   
  'And Finally we have to show the TabStops if any defined
  If m_bytAllowTabs <> ePermission.None Then
    On Error Resume Next
       If UBound(m_arrTabStop) > 0 Then
          Dim intX As Integer
          For intX = 0 To UBound(m_arrTabStop) - 1
             Select Case m_arrTabAlign(intX)
                Case eAlign.tLeft
                   m_cIL.DrawImage 7, mHdc, UserControl.ScaleX(m_arrTabStop(intX), vbTwips, vbPixels) + m_lLeftMargin, mRect.Bottom - IconY
                Case eAlign.tRight
                   m_cIL.DrawImage 9, mHdc, UserControl.ScaleX(m_arrTabStop(intX), vbTwips, vbPixels) + m_lLeftMargin, mRect.Bottom - IconY
                Case eAlign.tCentered
                   m_cIL.DrawImage 11, mHdc, UserControl.ScaleX(m_arrTabStop(intX), vbTwips, vbPixels) + m_lLeftMargin, mRect.Bottom - IconY
             End Select
          Next intX
       End If
   End If
End Sub



Private Sub pDrawLine(Pos As Long)
'While an Indent is moving, MS-Word draws a dotted line in the Text-Area
'We are trying to do same here, although it looks slightly different
'(3-dotted-Line does not exist as API-Constant)

Dim hdc As Long
Dim rClient As RECT
Static oldPos As Long
Dim lCount As Long

    
   If m_hwndBound = 0 Then Exit Sub
   GetClientRect m_hwndBound, rClient
   InflateRect rClient, 0, -2
   hdc = GetDC(m_hwndBound)
 
   If Pos = 0 Then
      rClient.Left = oldPos - 1
      rClient.Right = oldPos + 1
      InvalidateRect m_hwndBound, rClient, False
   Else
      For lCount = rClient.Top + 2 To rClient.Bottom - 2 Step 2
         If lCount Mod 8 > 0 Then
            SetPixel hdc, Pos, lCount, vbBlack
         End If
      Next lCount
   End If
   oldPos = Pos
End Sub



Private Property Get InDev() As Boolean
   'Original Code comes from VBAccelerator
   ' This function is called from a debug.assert call
   ' so m_bIndev is only ever set in DesignTime -
   ' debug.assert is not compiled into executables.
   m_bInDev = True
   InDev = m_bInDev
End Property
Private Sub pCreateImageList()
'Original Code comes from VBAccelerator
'Used to Create the ImageList from a Ressource-Picture-Strip

Dim idRes As Long
   Set m_cIL = New cImageList
      m_cIL.IconSizeX = IconX
   m_cIL.IconSizeY = IconY
   m_cIL.Create
   m_cIL.ColourDepth = ILC_COLOR4
   idRes = 101


   Debug.Assert (InDev() = True)
   If (m_bInDev) Then
      Dim stdPic As New StdPicture
      Set stdPic = LoadResPicture(idRes, vbResBitmap)
      m_cIL.AddFromHandle stdPic.Handle, IMAGE_BITMAP, , &HFF00FF
      Set stdPic = Nothing
   Else
      m_cIL.AddFromResourceID idRes, App.hInstance, IMAGE_BITMAP, , False, &HFF00FF
   End If

End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'If Left-Button has been clicked,
'determine whether Cursor is just over one of the Margins/Indents/TabStops
'm_bytMoving will contain the currently moved 'Object'

'I use PtInRect for the Indents/Tabs.
'I personally find it cleaner instead of using something like
'x>= m_lRightIndent - (iconX /2) and x <= m_lRightIndent + (iconX / 2) and y>= ......

   If Button = vbLeftButton Then
      If PtInRect(m_rRightIndent, ByVal CLng(X), ByVal CLng(Y)) Then
         If m_bytAllowIndents = ePermission.ByUser Then m_bytMoving = cRightIndent
      ElseIf PtInRect(m_rLeftIndent, CLng(X), CLng(Y)) Then
         If m_bytAllowIndents = ePermission.ByUser Then m_bytMoving = cLeftIndent
      ElseIf PtInRect(m_rFirstLineIndent, CLng(X), CLng(Y)) Then
         If m_bytAllowIndents = ePermission.ByUser Then m_bytMoving = cFirstLineIndent
      ElseIf PtInRect(m_rHangingIndent, CLng(X), CLng(Y)) Then
         If m_bytAllowIndents = ePermission.ByUser Then m_bytMoving = cHangingIndent
      ElseIf X >= m_lLeftMargin - 4 And X <= m_lLeftMargin + 4 Then
         If m_bytAllowMargins = ePermission.ByUser Then m_bytMoving = cLeftMargin
      ElseIf X >= mRect.Right - m_lRightMargin - 4 And X <= mRect.Right - m_lRightMargin + 4 Then
         If m_bytAllowMargins = ePermission.ByUser Then m_bytMoving = cRightMargin
      Else
         If m_bytAllowTabs = ePermission.ByUser Then
            Dim intX As Integer
            If UBound(m_arrTabStop) > 0 Then
                For intX = 0 To UBound(m_arrTabStop) - 1
                   If X >= UserControl.ScaleX(m_arrTabStop(intX), vbTwips, vbPixels) + m_lLeftMargin - 8 And X <= UserControl.ScaleX(m_arrTabStop(intX), vbTwips, vbPixels) + m_lLeftMargin + 8 And Y <= mRect.Bottom And Y >= mRect.Bottom - IconY - 8 Then
                      m_bytMoving = cTabStop
                      m_iTabMoving = intX
                      Exit For
                   End If
                Next
            End If
            If m_iTabMoving = -1 Then
               AddTab UserControl.ScaleX((mRect.Left - m_lLeftMargin) + X, vbPixels, vbTwips), tLeft
               m_iTabMoving = UBound(m_arrTabStop) - 1
               m_bytMoving = cTabStop
            End If
         End If
      End If
      pDraw
   End If
   
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sPos As Single

'Determine if and what is currently moved

'The (x Mod m_lRulerStep) part assures snapping of the Indents to the RulerScale
'For Example every 0.125 Inch or 6 Pixels
   If PtInRect(m_rHangingIndent, CLng(X), CLng(Y)) Then
      UserControl.Extender.ToolTipText = "Hanging Indent"
   ElseIf PtInRect(m_rLeftIndent, CLng(X), CLng(Y)) Then
      UserControl.Extender.ToolTipText = "Left Indent"
   ElseIf PtInRect(m_rRightIndent, CLng(X), CLng(Y)) Then
      UserControl.Extender.ToolTipText = "Right Indent"
   ElseIf PtInRect(m_rFirstLineIndent, CLng(X), CLng(Y)) Then
      UserControl.Extender.ToolTipText = "First Line Indent"
   ElseIf X >= m_lLeftMargin - 4 And X <= m_lLeftMargin + 4 Then
      UserControl.Extender.ToolTipText = "Left Margin"
   ElseIf X >= mRect.Right - m_lRightMargin - 4 And X <= mRect.Right - m_lRightMargin + 4 Then
      UserControl.Extender.ToolTipText = "Right Margin"
   Else
      UserControl.Extender.ToolTipText = ""
   End If
   
   
   Select Case m_bytMoving
      Case 0
         'Nothing is moved so time to Set MousePointer
         If X >= m_lLeftMargin - 4 And X <= m_lLeftMargin + 4 Then
            If PtInRect(m_rHangingIndent, CLng(X), CLng(Y)) Then
               Screen.MousePointer = 0
            ElseIf PtInRect(m_rFirstLineIndent, CLng(X), CLng(Y)) Then
               Screen.MousePointer = 0
            ElseIf PtInRect(m_rLeftIndent, CLng(X), CLng(Y)) Then
               Screen.MousePointer = 0
            Else
               Screen.MousePointer = vbSizeWE
            End If
         ElseIf X >= mRect.Right - m_lRightMargin - 4 And X <= mRect.Right - m_lRightMargin + 4 Then
            If PtInRect(m_rRightIndent, CLng(X), CLng(Y)) Then
               Screen.MousePointer = 0
            Else
               Screen.MousePointer = vbSizeWE
            End If
         Else
            Screen.MousePointer = 0
         End If
         'Now exit Sub in order to avoid redrawing of Ruler at the End of this Sub
         Exit Sub
      Case cLeftMargin
         Screen.MousePointer = 0
         If X > 0 And X <= mRect.Right - m_lRightMargin - 115 Then
            sPos = Fix(UserControl.ScaleX(X, vbPixels, vbTwips) / m_sglQuantise) * m_sglQuantise
            sPos = Fix(sPos / m_sglQuantise) * m_sglQuantise + m_sglQuantise
            LeftMargin = sPos
         End If
      Case cRightMargin
         Screen.MousePointer = 0
         If X <= mRect.Right And X >= m_lLeftMargin + 115 Then
            sPos = UserControl.ScaleX(mRect.Right, vbPixels, vbTwips) - (Fix(UserControl.ScaleX(X, vbPixels, vbTwips) / m_sglQuantise) * m_sglQuantise)
            sPos = Fix(sPos / m_sglQuantise) * m_sglQuantise + m_sglQuantise
            RightMargin = sPos
         End If
      Case cFirstLineIndent
         sPos = UserControl.ScaleX(mRect.Left - m_lLeftMargin - m_lLeftIndent, vbPixels, vbTwips) + (Fix(UserControl.ScaleX(X, vbPixels, vbTwips) / m_sglQuantise) * m_sglQuantise)
         sPos = Fix(sPos / m_sglQuantise) * m_sglQuantise + m_sglQuantise
         If FirstLineIndent <> sPos Then
            pDrawLine 0
            FirstLineIndent = sPos
         End If
         pDrawLine mRect.Left + m_lLeftMargin + m_lLeftIndent + m_lFirstLineIndent
      Case cHangingIndent
         sPos = UserControl.ScaleX(mRect.Left - m_lLeftMargin, vbPixels, vbTwips) + (Fix(UserControl.ScaleX(X, vbPixels, vbTwips) / m_sglQuantise) * m_sglQuantise)
          sPos = Fix(sPos / m_sglQuantise) * m_sglQuantise + m_sglQuantise
         If LeftIndent <> sPos Then
            pDrawLine 0
            FirstLineIndent = FirstLineIndent + (LeftIndent - sPos)
            LeftIndent = sPos

         End If
         pDrawLine mRect.Left + m_lLeftMargin + m_lLeftIndent
      Case cLeftIndent
        
         sPos = UserControl.ScaleX(mRect.Left - m_lLeftMargin, vbPixels, vbTwips) + (Fix(UserControl.ScaleX(X, vbPixels, vbTwips) / m_sglQuantise) * m_sglQuantise)
         sPos = Fix(sPos / m_sglQuantise) * m_sglQuantise + m_sglQuantise
         If LeftIndent <> sPos Then
             pDrawLine 0
            LeftIndent = sPos
         End If
         pDrawLine mRect.Left + m_lLeftMargin + m_lLeftIndent

      Case cRightIndent 'Handled different because auf Position-Calculation
        sPos = UserControl.ScaleX(mRect.Right - m_lRightMargin, vbPixels, vbTwips) - (Fix(UserControl.ScaleX(X, vbPixels, vbTwips) / m_sglQuantise) * m_sglQuantise)
        sPos = Fix(sPos / m_sglQuantise) * m_sglQuantise
         If RightIndent <> sPos Then
            pDrawLine 0
            RightIndent = sPos
         End If
         pDrawLine mRect.Right - m_lRightMargin - m_lRightIndent
      Case cTabStop
         sPos = UserControl.ScaleX(mRect.Left - m_lLeftMargin, vbPixels, vbTwips) + (Fix(UserControl.ScaleX(X, vbPixels, vbTwips) / m_sglQuantise) * m_sglQuantise)
         sPos = Fix(sPos / m_sglQuantise) * m_sglQuantise + m_sglQuantise
         
         If Y >= mRect.Bottom Or Y <= mRect.Bottom - IconY - 8 Then
               pDrawLine 0
               m_arrTabStop(m_iTabMoving) = mRect.Left - (LeftMargin * 2)
         Else
            If m_arrTabStop(m_iTabMoving) <> sPos Then
               pDrawLine 0
               m_arrTabStop(m_iTabMoving) = sPos
            End If
         End If
         pDrawLine m_lLeftMargin + UserControl.ScaleX(m_arrTabStop(m_iTabMoving), vbTwips, vbPixels)
   End Select
   
   'Redraw the Ruler. If nothing move sub has already been exited
   pDraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Check whats moved and Raise the specified Events
   Select Case m_bytMoving
      Case 0
         Exit Sub
      Case cLeftMargin, cRightMargin
         RaiseEvent MarginChanged(UserControl.ScaleX(m_lLeftMargin, vbPixels, vbTwips), UserControl.ScaleX(m_lRightMargin, vbPixels, vbTwips))
      Case cLeftIndent, cRightIndent, cFirstLineIndent, cHangingIndent
        RaiseEvent IndentChanged(UserControl.ScaleX(m_lLeftIndent, vbPixels, vbTwips), _
            UserControl.ScaleX(m_lFirstLineIndent, vbPixels, vbTwips), _
            UserControl.ScaleX(m_lRightIndent, vbPixels, vbTwips))
      Case cTabStop
         If m_arrTabStop(m_iTabMoving) = mRect.Left - (LeftMargin * 2) Then
            RemoveTab m_iTabMoving
         Else
            SortTabs
            RaiseEvent TabStopChanged(m_iTabCount, m_arrTabStop, m_arrTabAlign)
            m_iTabMoving = -1
         End If
   End Select
   
   'Mousebutton Raised so no moving any longer
   m_bytMoving = 0
   'Clear the remaining Line in the Bound-Textbox
   pDrawLine 0
   pDraw
End Sub

Private Sub UserControl_Resize()
'User may resize the Ruler only in Width.
   If UserControl.Height <> cMinMaxHeight Then UserControl.Height = cMinMaxHeight
   m_lRulerLength = UserControl.ScaleWidth
   pDraw
End Sub


Private Sub UserControl_Initialize()
   pCreateImageList
End Sub
Private Sub UserControl_Terminate()
   Set m_cIL = Nothing
End Sub

Private Sub UserControl_InitProperties()
   m_bytAllowIndents = cdefAllowIndents
   m_bytAllowTabs = cdefAllowTabs
   m_bytAllowMargins = cdefAllowMargins
   
   m_iRulerScale = cdefRulerScale
   pSetRulerScale
   
   m_lRulerLength = cdefRulerLength
   LeftMargin = cdefLeftMargin
   RightMargin = cdefRightMargin
   LeftIndent = cdefLeftIndent
   RightIndent = cdefRightIndent
   FirstLineIndent = cdefFirstLineIndent
   m_sglQuantise = m_lRulerStep

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
       ReDim m_arrTabStop(0)
   ReDim m_arrTabAlign(0)
   m_iRulerScale = PropBag.ReadProperty("RulerScale", cdefRulerScale)
   pSetRulerScale
   
   m_bytAllowMargins = PropBag.ReadProperty("AllowMargins", cdefAllowMargins)
   m_bytAllowTabs = PropBag.ReadProperty("AllowTabs", cdefAllowTabs)
   m_bytAllowIndents = PropBag.ReadProperty("AllowIndents", cdefAllowIndents)

   
   Dim sFont As New StdFont
   sFont.Name = "Univers Condensed"
   sFont.Size = 8
   Set UserControl.Font = PropBag.ReadProperty("Font", sFont)
   Set sFont = Nothing
   
   
   m_lRulerLength = PropBag.ReadProperty("RulerLength", cdefRulerLength)
   LeftMargin = PropBag.ReadProperty("LeftMargin", cdefLeftMargin)
   RightMargin = PropBag.ReadProperty("RightMargin", cdefRightMargin)
   LeftIndent = PropBag.ReadProperty("LeftIndent", cdefLeftIndent)
   RightIndent = PropBag.ReadProperty("RightIndent", cdefRightIndent)
   FirstLineIndent = PropBag.ReadProperty("FirstLineIndent", cdefFirstLineIndent)
   m_sglQuantise = PropBag.ReadProperty("Quantise", m_sglRulerStep)
   m_iTabMoving = -1
   

   
   pDraw
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "RulerScale", m_iRulerScale, cdefRulerScale
   PropBag.WriteProperty "RulerLength", m_lRulerLength, cdefRulerLength
   PropBag.WriteProperty "RightMargin", RightMargin, cdefRightMargin
   PropBag.WriteProperty "LeftMargin", LeftMargin, cdefLeftMargin
   PropBag.WriteProperty "LeftIndent", LeftIndent, cdefLeftIndent
   PropBag.WriteProperty "RightIndent", RightIndent, cdefRightIndent
   PropBag.WriteProperty "FirstLineIndent", FirstLineIndent, cdefFirstLineIndent
   PropBag.WriteProperty "AllowMargins", m_bytAllowMargins, cdefAllowMargins
   PropBag.WriteProperty "AllowTabs", m_bytAllowTabs, cdefAllowTabs
   PropBag.WriteProperty "AllowIndents", m_bytAllowIndents, cdefAllowIndents
   PropBag.WriteProperty "Quantise", m_sglQuantise, m_sglRulerStep
   
   Dim sFont As New StdFont
   sFont.Name = "Univers Condensed"
   sFont.Size = 8
   PropBag.WriteProperty "Font", Font, sFont
   Set sFont = Nothing
End Sub

