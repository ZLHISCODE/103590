VERSION 5.00
Begin VB.UserControl SpeedButton 
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1170
   ScaleWidth      =   3600
   Begin VB.PictureBox picSpeedButton 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "SpeedButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BF_SOFT = &H1000
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2 '浅凹下
Private Const BDR_RAISEDINNER = &H4 '浅凸起
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下

Private Const SRCCOPY = &HCC0020
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Size
    X As Long
    Y As Long
End Type

Private Enum enuFloat
    平面 = 0
    凸起 = 1
    凹陷 = 2
End Enum

Public Enum enuAlignPic
    PicLeft = 0
    PicCenter = 1
    PicRight = 2
End Enum

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 20
End Type

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Event Click()

Private mintStyle As Integer
Private mstrCaption As String
Private mblnEnabled As Boolean
Private mbytPictureAlign As Byte
Private mfntFont As New StdFont
Private mblnShowCaption As Boolean

Public Property Get Font() As StdFont
    Set Font = mfntFont
End Property
Public Property Set Font(ByVal fntVar As StdFont)
    Dim typFont As LOGFONT
    Dim lngFont As Long
        
    Set mfntFont = fntVar
    
    '字体大小
    With typFont
        If mfntFont Is Nothing Then
            .lfHeight = 12
            .lfFaceName = "宋体" & vbNullString
        Else
            .lfHeight = mfntFont.Size
            .lfCharSet = mfntFont.Charset
            .lfFaceName = mfntFont.Name
        End If
    End With
    
    lngFont = CreateFontIndirect(typFont)
    SelectObject picSpeedButton.hdc, lngFont
    'DeleteObject picSpeedButton.hdc
    
    '重绘
    picSpeedButton.Cls
    Call picSpeedButton_Paint
End Property

Public Property Get Caption() As String
    Caption = mstrCaption
End Property
Public Property Let Caption(ByVal strValue As String)
    mstrCaption = strValue
    picSpeedButton.Cls
    Call picSpeedButton_Paint
End Property

Public Property Get PictureAlign() As enuAlignPic
    PictureAlign = mbytPictureAlign
End Property
Public Property Let PictureAlign(ByVal bytValue As enuAlignPic)
    mbytPictureAlign = bytValue
    picSpeedButton.Cls
    Call picSpeedButton_Paint
End Property

Public Property Get Picture() As StdPicture
    Set Picture = picPicture.Picture
End Property
Public Property Set Picture(ByVal picValue As StdPicture)
    Set picPicture.Picture = picValue
    picSpeedButton.Cls
    Call picSpeedButton_Paint
End Property

Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
Public Property Let Enabled(ByVal blnValue As Boolean)
    Dim lngColor As Long
    
    lngColor = picSpeedButton.BackColor
    mblnEnabled = blnValue
    mintStyle = 0
    With picSpeedButton
        If blnValue Then
            .Appearance = 1
            .BorderStyle = 0
        Else
            .Appearance = 0
            .BorderStyle = 1
            '.BackColor = BackColor
        End If
        .Enabled = blnValue
        .BackColor = lngColor
    End With
    
    Call picSpeedButton_Paint
End Property

Public Property Get ShowCaption() As Boolean
    ShowCaption = mblnShowCaption
End Property
Public Property Let ShowCaption(ByVal blnShowCaption As Boolean)
    mblnShowCaption = blnShowCaption
    picSpeedButton.Cls
    Call picSpeedButton_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = picSpeedButton.BackColor
End Property
Public Property Let BackColor(ByVal lngColor As OLE_COLOR)
    picSpeedButton.BackColor = lngColor
    Call picSpeedButton_Paint
End Property

Private Sub DrawSpeedButton(picBox As PictureBox, Optional intStyle As Integer)
'功能：将PictureBox模拟成3D平面按钮
'参数：
'  intStyle：按钮的凹凸
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    
    mintStyle = intStyle
    
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
            DrawEdge .hdc, PicRect, CLng(IIf(intStyle = enuFloat.凸起, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
    End With
End Sub

Private Sub picSpeedButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = False Then Exit Sub

    If Button = vbLeftButton Then
        DrawSpeedButton picSpeedButton, enuFloat.凹陷
    End If
    Call picSpeedButton_Paint
End Sub

Private Sub picSpeedButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = False Then Exit Sub

    If Val(picSpeedButton.Tag) = 1 Then
        If X < 0 Or Y < 0 Or X > Width Or Y > Height Then
            '离开按钮区域
            ReleaseCapture
            picSpeedButton.Tag = ""
            If Button = vbLeftButton Then
                '离开按钮区域，但鼠标左键按下
                DrawSpeedButton picSpeedButton, enuFloat.凸起
            Else
                DrawSpeedButton picSpeedButton, enuFloat.平面
            End If
            Call picSpeedButton_Paint
        Else
            '进入按钮区域
            picSpeedButton.Tag = ""
        End If
    Else
        '首次进入按钮区域
        SetCapture picSpeedButton.hwnd
        If Button = vbLeftButton And Val(picSpeedButton.Tag) = 0 Then
            DrawSpeedButton picSpeedButton, enuFloat.凹陷
        Else
            DrawSpeedButton picSpeedButton, enuFloat.凸起
        End If
        picSpeedButton.Tag = "1"
        Call picSpeedButton_Paint
    End If
    
End Sub

Private Sub picSpeedButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = False Then Exit Sub

'    If Val(picSpeedButton.Tag) = 1 Then
'        DrawSpeedButton picSpeedButton, enuFloat.凸起
'    Else
        DrawSpeedButton picSpeedButton, enuFloat.平面
'    End If
    Call picSpeedButton_Paint

    RaiseEvent Click
End Sub

Private Sub picSpeedButton_Paint()
    Dim lngResult As Long
    Dim intScaleMode As Integer
    Dim typPoint As Size
    Dim intPicX As Integer, intPicY As Integer
    Dim intTxtX As Integer, intTxtY As Integer
    Dim intFloat As Integer
    
    '获取文本的高宽
    lngResult = GetTextExtentPoint32(picSpeedButton.hdc, mstrCaption, LenB(StrConv(mstrCaption, vbFromUnicode)), typPoint) '或者API的prvStringLenth()获取字串长度
    
    '字体颜色
    lngResult = SetTextColor(picSpeedButton.hdc, IIf(Enabled, vbBlack, RGB(100, 100, 100)))
    
    intScaleMode = picSpeedButton.ScaleMode
    picSpeedButton.ScaleMode = vbPixels
    picPicture.ScaleMode = vbPixels
    
    Select Case mintStyle
        Case enuFloat.凹陷
            intFloat = 1
        Case enuFloat.凸起
            intFloat = -1
        Case Else   '默认平面
    End Select
    
    If mblnShowCaption Then
        Select Case PictureAlign
            Case enuAlignPic.PicCenter
                intPicX = (picSpeedButton.ScaleWidth - picPicture.ScaleWidth) \ 2 + intFloat
                intPicY = (picSpeedButton.ScaleHeight - picPicture.ScaleHeight - 10 - typPoint.Y) \ 2 + intFloat
                intTxtX = (picSpeedButton.ScaleWidth - typPoint.X) \ 2 + intFloat
                intTxtY = intPicY + typPoint.Y + 10 + intFloat
            Case enuAlignPic.PicRight
                intTxtX = (picSpeedButton.ScaleWidth - picPicture.ScaleWidth - 2 - typPoint.X) \ 2 + intFloat
                intTxtY = (picSpeedButton.ScaleHeight - typPoint.Y) \ 2 + intFloat
                intPicX = intTxtX + typPoint.X + 2 + intFloat
                intPicY = (picSpeedButton.ScaleHeight - picPicture.ScaleHeight) \ 2 + intFloat
            Case Else   '默认图片居左
                intPicX = (picSpeedButton.ScaleWidth - picPicture.ScaleWidth - 2 - typPoint.X) \ 2 + intFloat
                intPicY = (picSpeedButton.ScaleHeight - picPicture.ScaleHeight) \ 2 + intFloat
                intTxtX = intPicX + picPicture.ScaleWidth + 2 + intFloat
                intTxtY = (picSpeedButton.ScaleHeight - typPoint.Y) \ 2 + intFloat
        End Select
    Else
        intPicX = (picSpeedButton.ScaleWidth - picPicture.ScaleWidth) \ 2 + intFloat
        intPicY = (picSpeedButton.ScaleHeight - picPicture.ScaleHeight) \ 2 + intFloat
        intTxtX = picSpeedButton.ScaleWidth
        intTxtY = picSpeedButton.ScaleHeight
    End If
    
    Call TextOut(picSpeedButton.hdc, intTxtX, intTxtY, mstrCaption, LenB(StrConv(mstrCaption, vbFromUnicode)))
    
    If Picture <> 0 Then
        lngResult = BitBlt(picSpeedButton.hdc, intPicX, intPicY, _
                            picSpeedButton.ScaleWidth - 1, _
                            picSpeedButton.ScaleHeight - 1, _
                            picPicture.hdc, 1, 1, SRCCOPY)
    End If
    picSpeedButton.ScaleMode = intScaleMode
    picPicture.ScaleMode = intScaleMode
    
    'ReleaseDC picSpeedButton.hwnd, picSpeedButton.hdc
End Sub

Private Sub picSpeedButton_Resize()
    picSpeedButton.Cls
    Call picSpeedButton_Paint
End Sub

Private Sub UserControl_Initialize()
    Enabled = True
    ShowCaption = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim fntTmp As New StdFont
    
    Caption = PropBag.ReadProperty("Caption", "")
    Enabled = PropBag.ReadProperty("Enabled", True)
    PictureAlign = PropBag.ReadProperty("PictureAlign", enuAlignPic.PicLeft)
    ShowCaption = PropBag.ReadProperty("ShowCaption", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set fntTmp = PropBag.ReadProperty("Font", Nothing)
    If fntTmp Is Nothing Then
        With fntTmp
            .Name = "宋体"
            .Size = 12
        End With
    End If
    Set Font = fntTmp
    
    BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    picSpeedButton.BackColor = BackColor
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    picSpeedButton.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", BackColor
    PropBag.WriteProperty "Font", Font
    PropBag.WriteProperty "Caption", Caption
    PropBag.WriteProperty "Enabled", Enabled
    PropBag.WriteProperty "PictureAlign", PictureAlign
    PropBag.WriteProperty "Picture", Picture
    PropBag.WriteProperty "ShowCaption", ShowCaption
End Sub
