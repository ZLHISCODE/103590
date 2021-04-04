VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl ColorPicker 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   2190
   ScaleWidth      =   2190
   ToolboxBitmap   =   "ColorPicker.ctx":0000
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   370
      Left            =   45
      ScaleHeight     =   375
      ScaleWidth      =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   2085
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1915
      Width           =   200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   0
      Picture         =   "ColorPicker.ctx":0312
      ScaleHeight     =   1350
      ScaleWidth      =   2160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   470
      Width           =   2160
      Begin VB.Shape shpValue 
         BorderColor     =   &H00C56A31&
         FillColor       =   &H00FF8080&
         Height          =   270
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00C56A31&
         FillColor       =   &H00FF8080&
         Height          =   270
         Left            =   1890
         Top             =   1080
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1800
      Top             =   1935
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblColor 
      Caption         =   "&HFFFFFF"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1918
      UseMnemonic     =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum tomEnum
    tomFalse = 0
    tomTrue = -1
    tomUndefined = -9999999
    tomToggle = -9999998
    tomAutoColor = -9999997
    tomDefault = -9999996
    tomSuspend = -9999995
    tomResume = -9999994
    tomApplyNow = 0
    tomApplyLater = 1
    tomTrackParms = 2
    tomCacheParms = 3
    tomApplyTmp = 4
    tomBackward = -1073741823
    tomForward = 1073741823
    tomMove = 0
    tomExtend = 1
    tomNoSelection = 0
    tomSelectionIP = 1
    tomSelectionNormal = 2
    tomSelectionFrame = 3
    tomSelectionColumn = 4
    tomSelectionRow = 5
    tomSelectionBlock = 6
    tomSelectionInlineShape = 7
    tomSelectionShape = 8
    tomSelStartActive = 1
    tomSelAtEOL = 2
    tomSelOvertype = 4
    tomSelActive = 8
    tomSelReplace = 16
    tomEnd = 0
    tomStart = 32
    tomCollapseEnd = 0
    tomCollapseStart = 1
    tomClientCoord = 256
    tomAllowOffClient = 512
    tomNone = 0
    tomWords = 2
    tomDouble = 3
    tomDotted = 4
    tomDash = 5
    tomDashDot = 6
    tomDashDotDot = 7
    tomWave = 8
    tomThick = 9
    tomHair = 10
    tomDoubleWave = 11
    tomHeavyWave = 12
    tomLongDash = 13
    tomThickDash = 14
    tomThickDashDot = 15
    tomThickDashDotDot = 16
    tomThickDotted = 17
    tomThickLongDash = 18
    tomLineSpaceSingle = 0
    tomLineSpace1pt5 = 1
    tomLineSpaceDouble = 2
    tomLineSpaceAtLeast = 3
    tomLineSpaceExactly = 4
    tomLineSpaceMultiple = 5
    tomAlignLeft = 0
    tomAlignCenter = 1
    tomAlignRight = 2
    tomAlignJustify = 3
    tomAlignDecimal = 3
    tomAlignBar = 4
    tomAlignInterWord = 3
    tomAlignInterLetter = 4
    tomAlignScaled = 5
    tomAlignGlyphs = 6
    tomAlignSnapGrid = 7
    tomSpaces = 0
    tomDots = 1
    tomDashes = 2
    tomLines = 3
    tomThickLines = 4
    tomEquals = 5
    tomTabBack = -3
    tomTabNext = -2
    tomTabHere = -1
    tomListNone = 0
    tomListBullet = 1
    tomListNumberAsArabic = 2
    tomListNumberAsLCLetter = 3
    tomListNumberAsUCLetter = 4
    tomListNumberAsLCRoman = 5
    tomListNumberAsUCRoman = 6
    tomListNumberAsSequence = 7
    tomListParentheses = 65536
    tomListPeriod = 131072
    tomListPlain = 196608
    tomCharacter = 1
    tomWord = 2
    tomSentence = 3
    tomParagraph = 4
    tomLine = 5
    tomStory = 6
    tomScreen = 7
    tomSection = 8
    tomColumn = 9
    tomRow = 10
    tomWindow = 11
    tomCell = 12
    tomCharFormat = 13
    tomParaFormat = 14
    tomTable = 15
    tomObject = 16
    tomPage = 17
    tomMatchWord = 2
    tomMatchCase = 4
    tomMatchPattern = 8
    tomUnknownStory = 0
    tomMainTextStory = 1
    tomFootnotesStory = 2
    tomEndnotesStory = 3
    tomCommentsStory = 4
    tomTextFrameStory = 5
    tomEvenPagesHeaderStory = 6
    tomPrimaryHeaderStory = 7
    tomEvenPagesFooterStory = 8
    tomPrimaryFooterStory = 9
    tomFirstPageHeaderStory = 10
    tomFirstPageFooterStory = 11
    tomNoAnimation = 0
    tomLasVegasLights = 1
    tomBlinkingBackground = 2
    tomSparkleText = 3
    tomMarchingBlackAnts = 4
    tomMarchingRedAnts = 5
    tomShimmer = 6
    tomWipeDown = 7
    tomWipeRight = 8
    tomAnimationMax = 8
    tomLowerCase = 0
    tomUpperCase = 1
    tomTitleCase = 2
    tomSentenceCase = 4
    tomToggleCase = 5
    tomReadOnly = 256
    tomShareDenyRead = 512
    tomShareDenyWrite = 1024
    tomPasteFile = 4096
    tomCreateNew = 16
    tomCreateAlways = 32
    tomOpenExisting = 48
    tomOpenAlways = 64
    tomTruncateExisting = 80
    tomRTF = 1
    tomText = 2
    tomHTML = 3
    tomWordDocument = 4
    tomBold = -2147483647
    tomItalic = -2147483646
    tomUnderline = -2147483644
    tomStrikeout = -2147483640
    tomProtected = -2147483632
    tomLink = -2147483616
    tomSmallCaps = -2147483584
    tomAllCaps = -2147483520
    tomHidden = -2147483392
    tomOutline = -2147483136
    tomShadow = -2147482624
    tomEmboss = -2147481600
    tomImprint = -2147479552
    tomDisabled = -2147475456
    tomRevised = -2147467264
    tomNormalCaret = 0
    tomKoreanBlockCaret = 1
    tomIncludeInset = 1
    tomIgnoreCurrentFont = 0
    tomMatchFontCharset = 1
    tomMatchFontSignature = 2
    tomCharset = -2147483648#
    tomRE10Mode = 1
    tomUseAtFont = 2
    tomTextFlowMask = 12
    tomTextFlowES = 0
    tomTextFlowSW = 4
    tomTextFlowWN = 8
    tomTextFlowNE = 12
    tomUsePassword = 16
    tomNoIME = 524288
    tomSelfIME = 262144
End Enum

Private mvarColor As OLE_COLOR
Private mvarAutoColor As OLE_COLOR

Public Event pOK()
Public Event pCancel()



Public Property Get AutoColor() As OLE_COLOR
    AutoColor = mvarAutoColor
End Property

Public Property Let AutoColor(vData As OLE_COLOR)
    mvarAutoColor = vData
    PropertyChanged "AutoColor"
End Property
    
Public Property Get Tag() As String
    Tag = UserControl.Tag
End Property

Public Property Let Tag(vData As String)
    UserControl.Tag = vData
    PropertyChanged "Tag"
End Property
    
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Color() As OLE_COLOR
    Color = mvarColor
End Property

Public Property Let Color(vData As OLE_COLOR)
    mvarColor = vData
    Dim lRow As Long, lCol As Long
    shpValue.Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor = "���ɫ"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor = "����"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor = "����"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor = "����"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor = "����"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor = "��ɫ-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor = "���"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor = "���"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor = "��-��"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor = "��ɫ-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor = "��ɫ"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor = "ǳ��ɫ"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor = "���ɫ"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor = "����"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor = "ǳ��"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor = "������"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor = "��ɫ-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor = "�ۺ�"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor = "����"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor = "����"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor = "����"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor = "÷��"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor = "õ���"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor = "ǳ����"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor = "����"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor = "����"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 7
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
        shpValue.Visible = False
    End Select
    shpValue.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    If vData = tomAutoColor Or vData = -1 Then
    
    Else
        picColor.BackColor = vData
    End If
    If picColor.Visible Then picColor.SetFocus
    If Color = tomAutoColor Then
        DrawButton 2
    Else
        DrawButton 0
    End If
    
    PropertyChanged "Color"
End Property

Private Sub picColor_Click()
'    SendKeys "{ESCAPE}"
'    DoEvents
'    dlgThis.Color = IIf(mvarColor = tomAutoColor, vbBlack, mvarColor)
'    dlgThis.CancelError = True
'    On Error GoTo LL
'    dlgThis.ShowColor
'    mvarColor = dlgThis.Color
'    RaiseEvent pOK
'    Exit Sub
'LL:
'    RaiseEvent pCancel
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    lblColor.Caption = "������ɫ..."
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        RaiseEvent pCancel
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < Picture1.ScaleWidth And Y > 0 And Y < Picture1.ScaleHeight Then
        SetCapture Picture1.hWnd
        shpBorder.Visible = True
    Else
        ReleaseCapture
        Color = mvarColor
        shpBorder.Visible = False
    End If

    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    
    shpBorder.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If Picture1.Point(lX, lY) = -1 Then Exit Sub
    picColor.BackColor = Picture1.Point(lX, lY)
    Select Case CStr(Hex(picColor.BackColor))
    Case "0"
        lblColor = "��ɫ"
    Case "3399"
        lblColor = "��ɫ"
    Case "3333"
        lblColor = "���ɫ"
    Case "3300"
        lblColor = "����"
    Case "663300"
        lblColor = "����"
    Case "800000"
        lblColor = "����"
    Case "993333"
        lblColor = "����"
    Case "333333"
        lblColor = "��ɫ-80%"
    Case "80"
        lblColor = "���"
    Case "66FF"
        lblColor = "��ɫ"
    Case "8080"
        lblColor = "���"
    Case "8000"
        lblColor = "��ɫ"
    Case "808000"
        lblColor = "��ɫ"
    Case "FF0000"
        lblColor = "��ɫ"
    Case "996666"
        lblColor = "��-��"
    Case "808080"
        lblColor = "��ɫ-50%"
    Case "FF"
        lblColor = "��ɫ"
    Case "99FF"
        lblColor = "ǳ��ɫ"
    Case "CC99"
        lblColor = "���ɫ"
    Case "669933"
        lblColor = "����"
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
    Case "FF6633"
        lblColor = "ǳ��"
    Case "800080"
        lblColor = "������"
    Case "999999"
        lblColor = "��ɫ-40%"
    Case "FF00FF"
        lblColor = "�ۺ�"
    Case "CCFF"
        lblColor = "��ɫ"
    Case "FFFF"
        lblColor = "��ɫ"
    Case "FF00"
        lblColor = "����"
    Case "FFFF00"
        lblColor = "����"
    Case "FFCC00"
        lblColor = "����"
    Case "663399"
        lblColor = "÷��"
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
    Case "CC99FF"
        lblColor = "õ���"
    Case "99CCFF"
        lblColor = "��ɫ"
    Case "99FFFF"
        lblColor = "ǳ��"
    Case "CCFFCC"
        lblColor = "ǳ��"
    Case "FFFFCC"
        lblColor = "ǳ����"
    Case "FFCC99"
        lblColor = "����"
    Case "FF99CC"
        lblColor = "����"
    Case "FFFFFF"
        lblColor = "��ɫ"
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    
    Color = Picture1.Point(lX, lY)
    RaiseEvent pOK
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton 3
    Picture2.Tag = "Down"
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And X <= Picture2.ScaleWidth And Y >= 0 And Y <= Picture2.ScaleHeight Then
        SetCapture Picture2.hWnd         '����ToolTipText���������ˣ�
        '������룡����
        If Picture2.Tag = "Down" Then
            DrawButton 3
        Else
            DrawButton 1
        End If
    Else
        If Picture2.Tag <> "" Then
            DrawButton 3
        Else
            '����Ƴ�������                 '����ToolTipText���������ˣ�
            ReleaseCapture
            If Color = tomAutoColor Then
                DrawButton 2
            Else
                DrawButton 0
            End If
        End If
    End If
End Sub

Private Sub DrawButton(lDrawStyle As Long)
    '0:��ͨ &H8000000F    1:�ƶ�  &HEED2C1   2:ѡ�� &HE8E6E1    3:����  &HE2B598          �߿�:&HC56A31
    On Error Resume Next
    If lDrawStyle = 2 Then lDrawStyle = 0
    Cls
    Select Case lDrawStyle
    Case 0  '��ͨ
        Picture2.BackColor = &H8000000F
    Case 1  '�ƶ�
        Picture2.BackColor = &HEED2C1
        Picture2.Line (0, 0)-(Picture2.ScaleWidth - Screen.TwipsPerPixelX, Picture2.ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
    Case 2  'ѡ��
        shpBorder.Visible = False
        shpValue.Visible = False
        Picture2.BackColor = &HE8E6E1
        Picture2.Line (0, 0)-(Picture2.ScaleWidth - Screen.TwipsPerPixelX, Picture2.ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
    Case 3  '����
        Picture2.BackColor = &HE2B598
        Picture2.Line (0, 0)-(Picture2.ScaleWidth - Screen.TwipsPerPixelX, Picture2.ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
    End Select
    
    Picture2.Line (90, 90)-(290, 290), AutoColor, BF
    Picture2.Line (90, 90)-(290, 290), RGB(133, 133, 133), B
    Picture2.CurrentX = 900
    Picture2.CurrentY = 90
    Picture2.Print "�Զ�"
    Refresh
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2.Tag = ""
    mvarColor = tomAutoColor
    DrawButton 3
    RaiseEvent pOK
End Sub

Private Sub UserControl_Initialize()
    Color = vbWhite
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        RaiseEvent pCancel
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Color = PropBag.ReadProperty("Color", vbWhite)
    AutoColor = PropBag.ReadProperty("AutoColor", vbBlack)
    If mvarColor = tomAutoColor Then
        DrawButton 2
    Else
        DrawButton 0
    End If
End Sub

Private Sub UserControl_Resize()
    Width = 2190
    Height = 2190
End Sub

Private Sub UserControl_Show()
'    If mvarColor = tomAutoColor Then
'        DrawButton 2
'    Else
'        DrawButton 0
'    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Color", Color, vbWhite
    PropBag.WriteProperty "AutoColor", AutoColor, vbBlack
    
    PropertyChanged "Color"
    PropertyChanged "AutoColor"
End Sub
