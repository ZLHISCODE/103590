VERSION 5.00
Begin VB.UserControl ZLScrollBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "ZLScrollBar.ctx":0000
   ScaleHeight     =   2040
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ZLScrollBar.ctx":0014
   Begin VB.Shape shpMove 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H000000FF&
      Height          =   1935
      Left            =   2040
      Top             =   0
      Width           =   144
   End
End
Attribute VB_Name = "ZLScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum ScrollStyleType
    sstHScroll = 0
    sstVScroll = 1
End Enum

Public Enum ColorFillType
    cftDefault = 0
    cftB = 1
    cftBF = 2
    cftShade = 3
    cft3D = 4
End Enum

'Public Enum AlignType
'    atAlignNone = 0
'    atAlignTop = 1
'    atAlignBottom = 2
'    atAlignLeft = 3
'    atAlignRight = 4
'    atAlignClient = 5
'End Enum

Private Type POINT
    X As Long
    Y As Long
End Type



Private mStartMovePoint As POINT
Private mStartMoveBlockPostion As POINT
Private mStartMovePostion As Long

Private mMouseState As Integer

Private mAllowMouseChange As Integer
Private mAutoShowBlock As Boolean

Private mScrollStyle As ScrollStyleType   '滚动条样式
Private mColorFillStyle As ColorFillType  '颜色填充样式

Private mMax As Long '最大值
Private mMin As Long '最小值
Private mPosition As Long '当前位置

Private mBeginColor As OLE_COLOR '开始眼神
Private mEndColor As OLE_COLOR '结束颜色

Private mUnitFillLen As Long   '每次填充的长度
Private mUnitSplitLen As Long  '颜色填充的间隔距离

Private mstrHint As String

Private mtoolTip As New clsToolTip

'Private mAlign As AlignType

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Public Event OnPositionChange(lngOldPosition As Long, lngNewPostion As Long)

Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event OnKeyDown(KeyCode As Integer, Shift As Integer)
Public Event OnKeyUp(KeyCode As Integer, Shift As Integer)
Public Event OnKeyPress(KeyAscii As Integer)

Public Event OnClick()


Private Const DEFAULT_BLOCK_WIDTH As Long = 145



'=========================================================================================================

'Property Get Align() As AlignType
'    Align = mAlign
'End Property
'
'Property Let Align(value As AlignType)
'    mAlign = value
'End Property

Property Get AllowMouseChange() As Boolean
    AllowMouseChange = mAllowMouseChange
End Property

Property Let AllowMouseChange(value As Boolean)
    mAllowMouseChange = value
End Property


Property Get AllowShowBlock() As Boolean
    AllowShowBlock = shpMove.Visible
End Property

Property Let AllowShowBlock(value As Boolean)
    shpMove.Visible = value
End Property


Property Get AutoShowBlock() As Boolean
    AutoShowBlock = mAutoShowBlock
End Property

Property Let AutoShowBlock(value As Boolean)
    mAutoShowBlock = value
    
    If mAutoShowBlock Then shpMove.Visible = False
End Property


Property Get ScrollStyle() As ScrollStyleType
    ScrollStyle = mScrollStyle
End Property

Property Let ScrollStyle(value As ScrollStyleType)
    mScrollStyle = value
End Property


Property Get ColorFillStyle() As ColorFillType
    ColorFillStyle = mColorFillStyle
End Property

Property Let ColorFillStyle(value As ColorFillType)
    mColorFillStyle = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property


Property Get Max() As Long
    Max = mMax
End Property

Property Let Max(value As Long)
    mMax = IIf(value < mPosition, mPosition, IIf(value <= mMin, mMin + 1, value))
    
    If Enabled Then
        Call SetBlockSize
        Call DrawScrollBar
    End If
End Property


Property Get Min() As Long
    Min = mMin
End Property

Property Let Min(value As Long)
    mMin = IIf(value > mPosition, mPosition, IIf(value >= mMax, mMax - 1, value))
    
    If Enabled Then
        Call SetBlockSize
        Call DrawScrollBar
    End If
End Property


Property Get Position() As Long
    Position = mPosition
End Property

Property Let Position(value As Long)

    If mMouseState <> 1 Then
        
        Dim lngOldPostion As Long
        
        lngOldPostion = mPosition
        mPosition = IIf(value > mMax, mMax, IIf(value < mMin, mMin, value))
        
        If Enabled Then
            Call DrawScrollBar
            'RaiseEvent OnPositionChange(lngOldPostion, value)
        End If
    End If
End Property


Property Get BeginColor() As OLE_COLOR
    BeginColor = mBeginColor
End Property

Property Let BeginColor(value As OLE_COLOR)
    mBeginColor = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property


Property Get EndColor() As OLE_COLOR
    EndColor = mEndColor
End Property

Property Let EndColor(value As OLE_COLOR)
    mEndColor = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property


Property Get UnitFillLen() As Long
    UnitFillLen = mUnitFillLen
End Property

Property Let UnitFillLen(value As Long)
    mUnitFillLen = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property


Property Get UnitSplitLen() As Long
    UnitSplitLen = mUnitSplitLen
End Property

Property Let UnitSplitLen(value As Long)
    mUnitSplitLen = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property


Property Get Hint() As String
    Hint = mstrHint
End Property

Property Let Hint(value As String)
    mstrHint = value
End Property


Property Get BlockColor() As OLE_COLOR
    BlockColor = shpMove.BorderColor
End Property

Property Let BlockColor(value As OLE_COLOR)
    shpMove.BorderColor = value
End Property



'=========================================================================================================


'appearance---
Property Get AppearanceStyle() As AppearanceStyleType
    AppearanceStyle = Appearance
End Property


Property Let AppearanceStyle(value As AppearanceStyleType)
    Appearance = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property

'autoredraw---
Property Get IsAutoRedraw() As Boolean
    IsAutoRedraw = AutoRedraw
End Property

Property Let IsAutoRedraw(value As Boolean)
    AutoRedraw = value
End Property

'backstyle---
Property Get BackSty() As BackStyleType
    BackSty = BackStyle
End Property

Property Let BackSty(value As BackStyleType)
    BackStyle = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property

'borderstyle---
Property Get BorderSty() As BorderStyleType
    BorderSty = BorderStyle
End Property

Property Let BorderSty(value As BorderStyleType)
    BorderStyle = value
End Property

'enable---
Property Get IsEnable() As Boolean
    IsEnable = Enabled
End Property

Property Let IsEnable(value As Boolean)
    Enabled = value
    
    If Enabled Then
        Call SetBlockSize
        Call DrawScrollBar
    End If
End Property


'scaleheight---
Property Get ScaleH() As Long
    ScaleH = ScaleHeight
End Property

Property Let ScaleH(value As Long)
    ScaleHeight = value
End Property

'scalewidth---
Property Get ScaleW() As Long
    ScaleW = ScaleWidth
End Property

Property Let ScaleW(value As Long)
    ScaleWidth = value
End Property

'scaleleft
Property Get ScaleL() As Long
    ScaleL = ScaleLeft
End Property

Property Let ScaleL(value As Long)
    ScaleLeft = value
End Property

'scaletop---
Property Get ScaleT() As Long
    ScaleT = ScaleTop
End Property

Property Let ScaleT(value As Long)
    ScaleTop = value
End Property

'scalemode---
Property Get ScaleType() As ScaleModeConstants
    ScaleType = ScaleMode
End Property

Property Let ScaleType(value As ScaleModeConstants)
    ScaleMode = value
End Property

'mousepointer---
Property Get MousePointerType() As MousePointerConstants
    MousePointerType = MousePointer
End Property

Property Let MousePointerType(value As MousePointerConstants)
    MousePointer = value
End Property

'mouseicon---
Property Get MouseIco() As StdPicture
    Set MouseIco = MouseIcon
End Property

Property Set MouseIco(value As StdPicture)
    Set MouseIcon = value
End Property

'backcolor---
Property Get BkColor() As OLE_COLOR
    BkColor = BackColor
End Property

Property Let BkColor(value As OLE_COLOR)
    BackColor = value
    
    If Enabled Then
        Call DrawBackGround
        Call DrawColor
    End If
End Property


'hwnd---
Property Get Handle() As OLE_HANDLE
    Handle = hWnd
End Property


Private Sub UserControl_Click()
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnClick
End Sub

'=========================================================================================================


Private Sub UserControl_Initialize()
    mMouseState = 0
    
    mScrollStyle = sstHScroll
    mColorFillStyle = cftBF
    
    mMax = 100
    mMin = 0
    mPosition = 0
    
    mBeginColor = vbGreen
    mEndColor = vbRed
    
    mUnitFillLen = 30
    mUnitSplitLen = 30
    
    mAllowMouseChange = True
    mAutoShowBlock = True
End Sub


Private Sub Draw3D(x1 As Long, y1 As Long, x2 As Long, y2 As Long, c As OLE_COLOR, colorFillSty As ColorFillType)
On Error Resume Next

    Dim dx As Long
    Dim cdx As Double, mb As Boolean
    
  
    Dim i As Integer
    Dim j As Double
    
    Dim R As Long, g As Long, b As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    
    Const k = 128
    
    dx = y2 - y1
    cdx = k / dx
    
    Call ToRgb(c, R, g, b)
    
    Select Case colorFillSty
        Case cft3D:
            j = 0
            For i = y1 To y2 / 2
                j = j + cdx
                
                r1 = Int(R - j) + 1
                g1 = Int(g - j) + 1
                b1 = Int(b - j) + 1
                
                r1 = IIf(r1 > 255, 255, IIf(r1 < 0, 0, r1))
                g1 = IIf(g1 > 255, 255, IIf(g1 < 0, 0, g1))
                b1 = IIf(b1 > 255, 255, IIf(b1 < 0, 0, b1))
                
                UserControl.Line (x1, i)-(x2, i), RGB(r1, g1, b1), B
            Next i
            
            For i = y2 / 2 To y2
                j = j - cdx
                
                r1 = Int(R - j) + 1
                g1 = Int(g - j) + 1
                b1 = Int(b - j) + 1
                
                r1 = IIf(r1 > 255, 255, IIf(r1 < 0, 0, r1))
                g1 = IIf(g1 > 255, 255, IIf(g1 < 0, 0, g1))
                b1 = IIf(b1 > 255, 255, IIf(b1 < 0, 0, b1))
                
                UserControl.Line (x1, i)-(x2, i), RGB(r1, g1, b1), B
            Next i
        Case cftShade:
            j = k
            For i = y1 To y2
                j = j - cdx
    
                r1 = R - Int(j) + 1
                g1 = g - Int(j) + 1
                b1 = b - Int(j) + 1
    
                r1 = IIf(r1 > 255, 255, IIf(r1 < 0, 0, r1))
                g1 = IIf(g1 > 255, 255, IIf(g1 < 0, 0, g1))
                b1 = IIf(b1 > 255, 255, IIf(b1 < 0, 0, b1))
    
                UserControl.Line (x1, i)-(x2, i), RGB(r1, g1, b1), BF
            Next i
        Case cftDefault:
            Line (x1, y1)-(x2, y2), RGB(R, g, b)
        Case cftB:
            Line (x1, y1)-(x2, y2), RGB(R, g, b), B
        Case cftBF:
            Line (x1, y1)-(x2, y2), RGB(R, g, b), BF
    End Select
    
    
End Sub


Private Sub DrawBackGround()
On Error Resume Next
    Dim R As Long, g As Long, b As Long
    
    Call ToRgb(BackColor, R, g, b)
    
    '按钮背景只存在三种方式，3D背景、渐变背景、填充背景
    Call Draw3D(0, 0, Width, Height, RGB(R, g, b), IIf(mColorFillStyle = cft3D, cft3D, IIf(mColorFillStyle = cftShade, cftShade, cftBF)))
End Sub


Private Sub ToRgb(c As OLE_COLOR, ByRef lngR As Long, ByRef lngG As Long, ByRef lngB As Long)
    lngR = (c And &HFF&)
    lngG = (c And &HFF00&) \ 256&
    lngB = (c And &HFF0000) \ 65536
End Sub


Private Function FillColor(lngCurX As Long, lngFillLen As Long, c As OLE_COLOR, fillSty As ColorFillType) As Long
On Error Resume Next

    Dim R As Long, g As Long, b As Long
    
    Call Draw3D(lngCurX, 0, lngCurX + lngFillLen, Height, c, fillSty)
    
    FillColor = lngCurX + mUnitFillLen
End Function


Private Sub DrawColor()
On Error Resume Next

    Dim lngDistance As Long
    Dim lngCurX As Long
    Dim lngUnitPixels As Double
    Dim lngValuePixels As Double
    
    Dim lngTmR As Long, lngTmG As Long, lngTmB As Long
    
    Dim lngSR As Double, lngSG As Double, lngSB As Double
    Dim lngER As Double, lngEG As Double, lngEB As Double
    Dim lngRStep As Double, lngGStep As Double, lngBStep As Double
    
    If mPosition <= mMin Then Exit Sub
    
    Call ToRgb(mBeginColor, lngTmR, lngTmG, lngTmB)
    lngSR = lngTmR: lngSG = lngTmG: lngSB = lngTmB
    
    Call ToRgb(mEndColor, lngTmR, lngTmG, lngTmB)
    lngER = lngTmR: lngEG = lngTmG: lngEB = lngTmB
        
    
    lngDistance = Abs(mMax - mMin) '求两数字的差的绝对值
    lngUnitPixels = Width / lngDistance '当前值对应的控件像素长度
    
    If mPosition = mMax Then
        lngValuePixels = Width
    Else
        lngValuePixels = (mPosition - IIf(mMin < mMax, mMin, mMax)) * lngUnitPixels
    End If
        
    lngRStep = (lngER - lngSR) / (Width / (mUnitFillLen + mUnitSplitLen))
    lngGStep = (lngEG - lngSG) / (Width / (mUnitFillLen + mUnitSplitLen))
    lngBStep = (lngEB - lngSB) / (Width / (mUnitFillLen + mUnitSplitLen))
    
    lngCurX = 0
    While (lngCurX + mUnitFillLen <= lngValuePixels)
        lngSR = IIf(lngSR > 255, 255, IIf(lngSR < 0, 0, lngSR))
        lngSG = IIf(lngSG > 255, 255, IIf(lngSG < 0, 0, lngSG))
        lngSB = IIf(lngSB > 255, 255, IIf(lngSB < 0, 0, lngSB))
        
        If lngCurX + mUnitFillLen + mUnitSplitLen >= lngValuePixels Then
            lngCurX = FillColor(lngCurX, Round(lngValuePixels - lngCurX), RGB(Round(lngSR), Round(lngSG), Round(lngSB)), mColorFillStyle)
        Else
            lngCurX = FillColor(lngCurX, mUnitFillLen, RGB(Round(lngSR), Round(lngSG), Round(lngSB)), mColorFillStyle)
        End If
        
        lngCurX = lngCurX + mUnitSplitLen
        
        lngSR = lngSR + lngRStep
        lngSG = lngSG + lngGStep
        lngSB = lngSB + lngBStep
    Wend
    
    If lngCurX + mUnitFillLen > lngValuePixels Then
        Call FillColor(lngCurX, Round(lngValuePixels - lngCurX), RGB(Round(lngSR), Round(lngSG), Round(lngSB)), mColorFillStyle)
    End If
    
End Sub

'Private Sub SetMeSize()
'    If Parent Is Nothing Then Exit Sub
'
'    Dim pControls As Object
'
'    Set pControls = Parent.Controls
'
'
'    Select Case mAlign
'        Case atAlignNone: Exit Sub
'
'    End Select
'End Sub


Private Sub SetBlockSize()
    Dim lngUnitPixelLen As Double
    
    If Appearance = 1 Then
        shpMove.Height = Height - 60
    Else
        shpMove.Height = Height - 30
    End If
    
    
    lngUnitPixelLen = Width / Abs(mMax - mMin)
    
    If lngUnitPixelLen > DEFAULT_BLOCK_WIDTH Then
        shpMove.Width = lngUnitPixelLen
    Else
        shpMove.Width = DEFAULT_BLOCK_WIDTH
    End If
End Sub

Private Sub SetBlockPostion(value As Long)
On Error Resume Next
    
    Dim lngUnitPixelLen As Double
    Dim lngValuePixels As Long
    
    lngUnitPixelLen = Width / Abs(mMax - mMin)
    
    lngValuePixels = Round((mPosition - IIf(mMin < mMax, mMin, mMax)) * lngUnitPixelLen)

    shpMove.Top = 0
    shpMove.Left = lngValuePixels - Round((shpMove.Width / 2)) + 20

    If shpMove.Left <= 0 Then shpMove.Left = 30
    If shpMove.Left + shpMove.Width >= Width Then shpMove.Left = Width - shpMove.Width - 30
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    If X > shpMove.Left - 30 And X < shpMove.Left + shpMove.Width + 30 And mAllowMouseChange Then
        mStartMoveBlockPostion.X = shpMove.Left
        mStartMoveBlockPostion.Y = shpMove.Top
        
        mStartMovePoint.X = X
        mStartMovePoint.Y = Y
        
        mStartMovePostion = mPosition
        
        mMouseState = 1
    End If
    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
End Sub


Private Sub DrawScrollBar()
    Call DrawBackGround
    Call DrawColor
    
    Call SetBlockPostion(mPosition)
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
       
    If Not Enabled Then Exit Sub
    
    If mAutoShowBlock Then
        Dim blnMouseEnter As Boolean
        blnMouseEnter = (X > 10) And (X < Width - 10) And (Y > 10) And (Y < Height - 10)

        If blnMouseEnter Then
            shpMove.Visible = True
            SetCapture hWnd
        Else
            shpMove.Visible = False
            mMouseState = 0

            ReleaseCapture
        End If
    End If
    
    
    If mMouseState = 1 Then
    
'        If mAutoShowBlock Then
'            shpMove.Visible = True
'        End If
        
        Dim lngMoveDistance As Long
        Dim lngMovePixelLen As Long
        
        lngMoveDistance = X - mStartMovePoint.X
        lngMovePixelLen = Width / (mMax - mMin)
        
'        shpMove.Left = mStartMoveBlockPostion.x + lngMoveDistance
        
'        If shpMove.Left <= 0 Then
'            shpMove.Left = 30
'            mPosition = mMin
'
'            Call DrawScrollBar
'
'            RaiseEvent OnPositionChange(mStartMovePostion, mPosition)
'            RaiseEvent OnMouseMove(Button, Shift, X, Y)
'
'            Exit Sub
'        End If
'
'        If shpMove.Left + shpMove.Width >= Width Then
'            shpMove.Left = Width - shpMove.Width - 30
'            mPosition = mMax
'
'            Call DrawScrollBar
'
'            RaiseEvent OnPositionChange(mStartMovePostion, mPosition)
'            RaiseEvent OnMouseMove(Button, Shift, X, Y)
'
'            Exit Sub
'        End If
        
        
        If Abs(lngMoveDistance) >= lngMovePixelLen Then
            mPosition = mStartMovePostion + Round(lngMoveDistance / lngMovePixelLen) ' + IIf(lngMoveDistance < 0, -0.5, 0.5))
            If mPosition < Min Then mPosition = mMin
            If mPosition > Max Then mPosition = mMax
            
            Call DrawScrollBar
            
            RaiseEvent OnPositionChange(mStartMovePostion, mPosition)
        End If
    End If
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    mMouseState = 0
    
'    If mAutoShowBlock Then
'        Dim blnMouseEnter As Boolean
'        blnMouseEnter = (X > 10) And (X < Width - 10) And (Y > 10) And (Y < Height - 10)
'
'        If Not blnMouseEnter Then shpMove.Visible = False
'    End If
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Call DrawBackGround

    If Enabled Then
        Call DrawColor
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Appearance = PropBag.ReadProperty("Appearance", 1)
    AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    BackStyle = PropBag.ReadProperty("BackStyle", 1)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Enabled = PropBag.ReadProperty("Enabled", True)
    
    ScaleHeight = PropBag.ReadProperty("ScaleHeight", 960)
    ScaleWidth = PropBag.ReadProperty("ScaleWidth", 960)
    ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    
    MousePointer = PropBag.ReadProperty("MousePointer", 0)
    BackColor = PropBag.ReadProperty("BackColor", vbBlack)
    mstrHint = PropBag.ReadProperty("Hint", "")
    
    
    mScrollStyle = PropBag.ReadProperty("ScrollStyle", sstHScroll)
    mColorFillStyle = PropBag.ReadProperty("ColorFillStyle", cftBF)
    
    mMax = PropBag.ReadProperty("Max", 100)
    mMin = PropBag.ReadProperty("Min", 0)
    mPosition = PropBag.ReadProperty("Position", 0)
    
    mBeginColor = PropBag.ReadProperty("BeginColor", vbGreen)
    mEndColor = PropBag.ReadProperty("EndColor", vbRed)
    
    mUnitFillLen = PropBag.ReadProperty("UnitFillLen", 30)
    mUnitSplitLen = PropBag.ReadProperty("UnitSplitLen", 30)
    
    shpMove.Visible = PropBag.ReadProperty("ShpMoveVisible", True)
    shpMove.BorderColor = PropBag.ReadProperty("ShpMoveBorderColor", vbYellow)
    mAllowMouseChange = PropBag.ReadProperty("AllowMouseChange", True)
    mAutoShowBlock = PropBag.ReadProperty("AutoShowBlock", True)
    
    
    
    
    If mAutoShowBlock Then shpMove.Visible = False
End Sub

Private Sub UserControl_Resize()
    Call DrawBackGround
    
    If Enabled Then
        Call DrawColor
        
        Call SetBlockSize
        Call SetBlockPostion(mPosition)
    End If
End Sub

Private Sub UserControl_Show()
    Call DrawBackGround
    
    If Enabled Then
        Call DrawColor
        
        Call SetBlockSize
        Call SetBlockPostion(mPosition)
        
        If mAutoShowBlock Then shpMove.Visible = False
    End If
    
    Call mtoolTip.CreateBalloon(Handle, mstrHint, szClassic, False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", Appearance, 1)
    Call PropBag.WriteProperty("AutoRedraw", AutoRedraw, False)
    Call PropBag.WriteProperty("BackStyle", BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    
    Call PropBag.WriteProperty("ScaleHeight", ScaleHeight, 960)
    Call PropBag.WriteProperty("ScaleWidth", ScaleWidth, 960)
    Call PropBag.WriteProperty("ScaleLeft", ScaleLeft, 960)
    Call PropBag.WriteProperty("ScaleTop", ScaleTop, 960)
    Call PropBag.WriteProperty("ScaleMode", ScaleMode, 960)
    
    Call PropBag.WriteProperty("MousePointer", MousePointer, 0)
    Call PropBag.WriteProperty("BackColor", BackColor, vbBack)
    Call PropBag.WriteProperty("Hwnd", hWnd, 0)
    Call PropBag.WriteProperty("Hint", Hint, "")
    
    Call PropBag.WriteProperty("ScrollStyle", mScrollStyle, sstHScroll)
    Call PropBag.WriteProperty("ColorFillStyle", mColorFillStyle, cftBF)
    
    Call PropBag.WriteProperty("Max", mMax, 100)
    Call PropBag.WriteProperty("Min", mMin, 0)
    Call PropBag.WriteProperty("Position", mPosition, 0)
    
    Call PropBag.WriteProperty("BeginColor", mBeginColor, vbGreen)
    Call PropBag.WriteProperty("EndColor", mEndColor, vbRed)
    
    Call PropBag.WriteProperty("UnitFillLen", mUnitFillLen, 30)
    Call PropBag.WriteProperty("UnitSplitLen", mUnitSplitLen, 30)
    
    Call PropBag.WriteProperty("ShpMoveVisible", shpMove.Visible, True)
    Call PropBag.WriteProperty("ShpMoveBorderColor", shpMove.BorderColor, vbYellow)
    Call PropBag.WriteProperty("AllowMouseChange", mAllowMouseChange, True)
    Call PropBag.WriteProperty("AutoShowBlock", mAutoShowBlock, True)
End Sub
