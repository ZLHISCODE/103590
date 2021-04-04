VERSION 5.00
Begin VB.Form frmMipComAlert 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3795
      Top             =   0
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4320
      Top             =   0
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3315
      Top             =   0
   End
   Begin VB.PictureBox picBackground 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   390
      ScaleHeight     =   2055
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   735
      Width           =   4425
      Begin VB.Shape shpLink 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         Height          =   15
         Left            =   75
         Top             =   1620
         Width           =   4305
      End
      Begin VB.Label lblMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMipComAlert.frx":0000
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   60
         TabIndex        =   4
         Top             =   75
         Width           =   4320
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品付款单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   90
         MouseIcon       =   "frmMipComAlert.frx":0184
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1755
         Width           =   1080
      End
   End
   Begin VB.PictureBox picCaption 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   390
      ScaleHeight     =   285
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   435
      Width           =   4425
      Begin VB.Image imgClose 
         Height          =   195
         Left            =   4140
         MouseIcon       =   "frmMipComAlert.frx":048E
         MousePointer    =   99  'Custom
         Picture         =   "frmMipComAlert.frx":0798
         Top             =   45
         Width           =   195
      End
      Begin VB.Label lblTopic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒消息"
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   60
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   45
         MouseIcon       =   "frmMipComAlert.frx":0A1E
         MousePointer    =   99  'Custom
         Picture         =   "frmMipComAlert.frx":0D28
         Top             =   45
         Width           =   195
      End
   End
   Begin VB.Image imgMessage 
      Height          =   240
      Left            =   4410
      Picture         =   "frmMipComAlert.frx":0FAE
      Top             =   1215
      Width           =   240
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000003&
      FillColor       =   &H00404040&
      Height          =   1845
      Left            =   0
      Top             =   0
      Width           =   2595
   End
End
Attribute VB_Name = "frmMipComAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''######################################################################################################################
'
'Private mblnStartUp As Boolean
'Private mlngMaxHeight As Long
'Private mclsMipSystemData As clsMipSystemData
'Private mbytLinkType As Byte
'Private mblnReaded As Boolean
'Private mblnShowing As Boolean
'Public Event OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
'Public Event BeforeShowMessage()
'Public Event AfterShowMessage()
'Public Event ReadMessage()
'
''######################################################################################################################
'Public Property Get Showing() As Boolean
'    Showing = mblnShowing
'End Property
'
'Public Function ShowAlert(ByVal strMessageTopic As String, _
'                            ByVal strMessageText As String, _
'                            ByVal strMessageLinkType As String, _
'                            ByVal strMessageLinkTitle As String, _
'                            ByVal strMessageLinkPara As String, _
'                            ByVal lngMessageWave As Long, _
'                            ByVal lngMessageAlert As Long) As Boolean
'    '---------------------------------------------------------------------------------------
'    '功能： 弹出提示窗口
'    '参数： strMessage-信息集
'    '       lngDelay-每条信息停止的时间
'    '---------------------------------------------------------------------------------------
'    Dim lngScreenX As Long
'    Dim lngScreenY As Long
'    Dim lngScaleX As Long
'    Dim lngScaleY As Long
'    Dim varTmp As Variant
'    Dim varTmp2 As Variant
'    Dim lngCount As Long
'
'    mblnShowing = True
'
'    mblnStartUp = True
'    mblnReaded = False
'
'    RaiseEvent BeforeShowMessage
'
'    lngScreenX = GetSystemMetrics(SM_CXFULLSCREEN)
'    lngScreenY = GetSystemMetrics(SM_CYFULLSCREEN)
'
'    lngScaleX = Me.Width - Me.ScaleWidth
'    lngScaleY = Me.Height - Me.ScaleHeight
'
'    mlngMaxHeight = picBackground.Height + picCaption.Height + 15
'
'    shp.Top = 0
'    shp.Left = 0
'    shp.Width = picBackground.Width + lngScaleX + 15
'    shp.Height = mlngMaxHeight
'
'    With picCaption
'        .Left = 15
'        .Top = 15
'        .Width = picBackground.Width
'    End With
'
'    With picBackground
'        .Left = 15
'        .Top = picCaption.Top + picCaption.Height
'    End With
'
'    Me.Height = 90
'    Me.Width = picBackground.Width + lngScaleX + 15
'    Me.Left = lngScreenX * Screen.TwipsPerPixelX - Me.Width - 15
'    Me.Top = (lngScreenY * Screen.TwipsPerPixelY) + 160
'
'    If strMessageLinkType <> "" Then
'        Select Case strMessageLinkType
'        Case "报表"
'            Call ShowMessage(strMessageTopic, strMessageText, 1, strMessageLinkTitle, strMessageLinkPara, lngMessageWave, lngMessageAlert)
'        Case "模块"
'            Call ShowMessage(strMessageTopic, strMessageText, 2, strMessageLinkTitle, strMessageLinkPara, lngMessageWave, lngMessageAlert)
'        Case Else
'            Call ShowMessage(strMessageTopic, strMessageText, 3, strMessageLinkTitle, strMessageLinkPara, lngMessageWave, lngMessageAlert)
'        End Select
'    Else
'        Call ShowMessage(strMessageTopic, strMessageText, , , , lngMessageWave, lngMessageAlert)
'    End If
'
'    ShowWindow Me.hWnd, 4
'    SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H10 Or &H1
'
'    tmrOpen.Enabled = True
'
'End Function
'
'Private Sub ShowMessage(ByVal strTopic As String, _
'                        ByVal strMessageText As String, _
'                        Optional ByVal bytLinkType As Byte, _
'                        Optional ByVal strLinkTitle As String, _
'                        Optional ByVal strLinkPara As String, _
'                        Optional ByVal lngWave As Long, _
'                        Optional ByVal lngAlert As Long)
'    '******************************************************************************************************************
'    '功能：显示通用消息（报表和模块不能同时显示）
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    mbytLinkType = bytLinkType
'
'    tmrAlert.Interval = lngAlert * 1000
'
'    If lngWave > 0 Then Call PlayWave(lngWave)
'    lblMessage.Caption = strMessageText
'
'    If lblMessage.Height > 705 Then lblMessage.Height = 705
'
'    lblTopic.Caption = IIf(strTopic = "", "提醒消息", strTopic)
'    lblLink.Caption = ""
'    lblLink.Tag = ""
'
'    If bytLinkType > 0 Then
'        lblLink.Caption = strLinkTitle
'        lblLink.Tag = strLinkPara
'    End If
'    lblLink.Visible = (bytLinkType > 0)
'    shpLink.Visible = (lblLink.Caption <> "")
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    '
'End Sub
'
'Private Sub imgClose_Click()
'    tmrClose.Interval = 1
'    tmrClose.Enabled = True
'End Sub
'
'Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        imgClose.Left = imgClose.Left + 15
'        imgClose.Top = imgClose.Top + 15
'    End If
'End Sub
'
'Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    tmrAlert.Enabled = False
'    tmrAlert.Enabled = True
'End Sub
'
'Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        imgClose.Left = imgClose.Left - 15
'        imgClose.Top = imgClose.Top - 15
'    End If
'End Sub
'
'Private Sub imgMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    tmrAlert.Enabled = False
'    tmrAlert.Enabled = True
'End Sub
'
'Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    tmrAlert.Enabled = False
'    tmrAlert.Enabled = True
'
'    If mblnReaded = False Then
'        mblnReaded = True
'        RaiseEvent ReadMessage
'    End If
'End Sub
'
'Private Sub lblLink_Click()
'
'    RaiseEvent OpenLink(mbytLinkType, lblLink.Tag)
'
'End Sub
'
'Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    tmrAlert.Enabled = False
'    tmrAlert.Enabled = True
'End Sub
'
'Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    tmrAlert.Enabled = False
'    tmrAlert.Enabled = True
'End Sub
'
'Private Sub picBackground_Paint()
'    Call DrawColorToColor(picBackground, &HFFFFFF, &HFFC0C0)
'    Call PicShowFlat(picBackground, 1)
'End Sub
'
'Private Sub picCaption_Paint()
'    Call DrawColorToColor(picCaption, &HFFFFFF, &HFFC0C0)
'    Call PicShowFlat(picCaption, 1)
'End Sub
'
'Private Sub tmrAlert_Timer()
'
'    '显示信息记录
'
'    tmrAlert.Enabled = False
'    tmrClose.Interval = 1
'    tmrClose.Enabled = True
'
'End Sub
'
'Private Sub tmrClose_Timer()
'
'    Dim lngHeight As Long
'    Dim lngTop As Long
'
'    lngHeight = Me.Height
'
'    If lngHeight > 90 Then
'
'        lngHeight = lngHeight - 30
'        lngTop = Me.Top + 30
'
'        On Error Resume Next
'
'        MoveWindow Me.hWnd, Me.Left / 15, lngTop / 15, Me.Width / 15, lngHeight / 15, 1
'        SetWindowPos Me.hWnd, -1, Me.Left / 15, lngTop / 15, Me.Width / 15, lngHeight / 15, &H10 Or &H1
'
'    Else
'        RaiseEvent AfterShowMessage
'        mblnShowing = False
'        Unload Me
'    End If
'End Sub
'
'Private Sub tmrOpen_Timer()
'
'    Dim lngHeight As Long
'    Dim lngNewHeight As Long
'    Dim lngScaleY As Long
'
'    Dim lngH As Long
'    Dim lngTop As Long
'
'    lngScaleY = Me.Height - Me.ScaleHeight
'    lngHeight = Me.Height
'
'    If lngHeight < mlngMaxHeight + lngScaleY Then
'        lngNewHeight = lngHeight + 30
'
'        If lngNewHeight > mlngMaxHeight + lngScaleY Then lngNewHeight = mlngMaxHeight + lngScaleY
'
'        lngH = Me.Height + (lngNewHeight - lngHeight)
'        lngTop = Me.Top - (lngNewHeight - lngHeight)
'
'        On Error Resume Next
'
'        MoveWindow Me.hWnd, Me.Left / 15, lngTop / 15, Me.Width / 15, lngH / 15, 1
'        SetWindowPos Me.hWnd, -1, Me.Left / 15, lngTop / 15, Me.Width / 15, lngH / 15, &H10 Or &H1
'
'    Else
'        tmrOpen.Enabled = False
'        DoEvents
'        tmrAlert.Enabled = True
'    End If
'End Sub
'
'
