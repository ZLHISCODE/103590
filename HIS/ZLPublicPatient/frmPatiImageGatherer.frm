VERSION 5.00
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Begin VB.Form frmPatiImageGatherer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "图像采集"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   Icon            =   "frmPatiImageGatherer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   0
      ScaleHeight     =   4485
      ScaleWidth      =   5445
      TabIndex        =   4
      Top             =   0
      Width           =   5475
      Begin ZLDSVideoProcess.DSCapture DSCap 
         Height          =   4440
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5400
         Object.Visible         =   -1  'True
         AutoScroll      =   0   'False
         AutoSize        =   0   'False
         AxBorderStyle   =   2
         Caption         =   ""
         Color           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         KeyPreview      =   -1  'True
         PixelsPerInch   =   96
         PrintScale      =   1
         Scaled          =   -1  'True
         DropTarget      =   0   'False
         HelpFile        =   ""
         ScreenSnap      =   0   'False
         SnapBuffer      =   10
         DoubleBuffered  =   0   'False
         Enabled         =   -1  'True
         IsStretch       =   0   'False
         IsShowState     =   -1  'True
         IsFullScreen    =   0   'False
         IsAdjustWindowSize=   0   'False
         IsFit           =   0   'False
         IsEscKeyQuitFullScreen=   -1  'True
         IsDblClickQuitFullScreen=   0   'False
         IsClickQuitFullScreen=   0   'False
         CurWidth        =   360
         CurHeight       =   296
         CurVideoWidth   =   356
         CurVideoHeight  =   274
         ShowModel       =   0
         CapParameterWindPos=   8
         SnatchWay       =   0
         ParameterCfgFileName=   ""
         HideCfgItem     =   0
         AppHandle       =   0
         Begin VB.PictureBox picRect 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   1170
            ScaleHeight     =   2055
            ScaleWidth      =   2385
            TabIndex        =   6
            Top             =   750
            Visible         =   0   'False
            Width           =   2385
         End
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "视频设置"
      Height          =   350
      Index           =   4
      Left            =   5940
      TabIndex        =   3
      Top             =   1350
      Width           =   950
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "采集"
      Height          =   350
      Index           =   0
      Left            =   5940
      TabIndex        =   0
      Top             =   90
      Width           =   950
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "确定"
      Height          =   350
      Index           =   2
      Left            =   5940
      TabIndex        =   1
      Top             =   510
      Width           =   950
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "返回"
      Height          =   350
      Index           =   3
      Left            =   5940
      TabIndex        =   2
      Top             =   930
      Width           =   950
   End
   Begin VB.Label lblShowInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "图像预览："
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   2820
      Width           =   915
   End
   Begin VB.Image imgPerson 
      BorderStyle     =   1  'Fixed Single
      Height          =   1425
      Left            =   5490
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1800
   End
End
Attribute VB_Name = "frmPatiImageGatherer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnNotLoad As Boolean
Private mstrPictureFilePath As String
Private mblnOK As Boolean '是否采集成功
Private mblnFirst As Boolean
Private Enum Enum_Button
    EM_FUN_采集 = 0
    EM_FUN_确定 = 2
    EM_FUN_返回 = 3
    EM_FUN_视频设置 = 4
End Enum
Private mlngButton As Long

Public Function ShowMe(ByVal frmParent As Form, ByRef strPictureFilePath As String) As Boolean
    '-----------------------------------------------
    '功能：窗体入口，显示图像采集窗口
    '参数：
    '   frmParent：父窗体
    '   strPictureFilePath：采集图片保存位置
    '返回：返回是否采集成功
    '编制：冉俊明
    '日期：2014-6-26
    '-----------------------------------------------
    On Error GoTo ErrHand
    mstrPictureFilePath = App.Path & "\person.bmp"
    mblnOK = False
    Me.Show 1, frmParent
    strPictureFilePath = mstrPictureFilePath
    ShowMe = mblnOK
    Exit Function
ErrHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Private Sub DSCap_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    mlngButton = 2
End Sub

Private Sub DSCap_OnMouseMove(ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    Dim blnMouseIn As Boolean '鼠标是否在picRect内
    
    If Not picRect.Visible Then Exit Sub
    
    X = ScaleX(X, vbPixels, vbTwips)
    Y = ScaleX(Y, vbPixels, vbTwips)
    blnMouseIn = picRect.Left < X And X <= picRect.Left + picRect.Width And picRect.Top < Y And Y <= picRect.Top + picRect.Height
    If blnMouseIn Then
        SetCapture picRect.hWnd
        picRect.MousePointer = vbSizePointer
        If mlngButton = vbLeftButton Then
            X = ReleaseCapture()
            SendMessage picRect.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        End If
    Else
        picRect.MousePointer = vbDefault
        ReleaseCapture
    End If
    picRect.Refresh
End Sub

Private Sub DSCap_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    mlngButton = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    DSCap.Width = picBack.Width
    DSCap.Height = picBack.Height
    If mblnFirst = True Then
        Call DSCap.RePreview
    End If
End Sub

Private Sub picRect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnMouseIn As Boolean '鼠标是否在picRect内
    If Not picRect.Visible Then Exit Sub

    blnMouseIn = (0 < X) And (X <= ScaleX(picRect.Width, vbTwips, vbPixels)) And (0 < Y) And (Y <= ScaleX(picRect.Height, vbTwips, vbPixels))
    If blnMouseIn Then
        SetCapture picRect.hWnd
        picRect.MousePointer = vbSizePointer
        If Button = vbLeftButton Then
            X = ReleaseCapture()
            SendMessage picRect.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        End If
    Else
        picRect.MousePointer = vbDefault
        ReleaseCapture
    End If
    picRect.Refresh
End Sub

Private Sub Form_Load()
    Dim strErr As String
    mblnNotLoad = False
    mblnFirst = False
    picRect.ScaleMode = vbPixels
    cmdButton(EM_FUN_确定).Enabled = False
    On Error GoTo ErrHand
    Call SetRectangle '制作一个透明矩形框
    DSCap.ReadParameterFromFile
    '进入预览模式
    strErr = DSCap.StartPreview
    If strErr <> "" Then GoTo ErrHand
    mblnFirst = True
    Exit Sub
ErrHand:
    mblnNotLoad = True
    MsgBox strErr, vbExclamation, gstrSysName
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnNotLoad = True Then Unload Me
End Sub

Private Sub cmdButton_Click(Index As Integer)
    Dim strErr As String, strFile As String

    Select Case Index
    Case EM_FUN_采集
        '将采集图象保存为BMP文件
        strErr = DSCap.CaptureBmpImageToFile(mstrPictureFilePath)
        If strErr <> "" Then GoTo ErrHand
        '显示采集图像
        imgPerson.Picture = DSCap.CaptureBmpImage
        cmdButton(EM_FUN_确定).Enabled = True
        DSCap.ShowModel = smAutoFitCut
    Case EM_FUN_确定
        mblnOK = True: Unload Me
    Case EM_FUN_返回
        Unload Me
    Case EM_FUN_视频设置
        '显示采集参数配置对话框
        strErr = DSCap.ShowCaptureParameterCfgDialog(0)
        If strErr <> "" Then GoTo ErrHand
        '重新进入预览模式
        Call DSCap.RePreview
    End Select
    Exit Sub
ErrHand:
    MsgBox strErr, vbExclamation, gstrSysName
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '释放资源
    Call DSCap.FreeRes
End Sub

Private Sub SetRectangle()
    '--------------------------------------------
    '功能：制作一个透明矩形框
    '--------------------------------------------
    Const RGN_DIFF = 4
    Dim lngOuterRgn As Long, lngInnerRgn As Long, lngCombinedRgn As Long
    Dim sinWidth As Single, sinHeight As Single
    Dim sinBorderWidth As Single, sinTitleHeight As Single
    If Not picRect.Visible Then Exit Sub


    '获取控件宽度和高度
    sinWidth = ScaleX(picRect.Width, vbTwips, vbPixels)
    sinHeight = ScaleY(picRect.Height, vbTwips, vbPixels)
    '外矩形框
    lngOuterRgn = CreateRectRgn(0, 0, sinWidth, sinHeight)
    '计算内矩形框宽度和高度
    sinBorderWidth = (sinWidth - picRect.ScaleWidth) / 2
    sinTitleHeight = sinHeight - sinBorderWidth - picRect.ScaleHeight
    '内矩形框
    lngInnerRgn = CreateRectRgn(sinBorderWidth + 3, sinTitleHeight + 3, picRect.ScaleWidth - 3, picRect.ScaleHeight - 3)
    '从窗体中去除创建“洞”的区域
    lngCombinedRgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn lngCombinedRgn, lngOuterRgn, lngInnerRgn, RGN_DIFF
    '限制窗口到区域
    SetWindowRgn picRect.hWnd, lngCombinedRgn, False
End Sub
