VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmWin 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   8040
   ControlBox      =   0   'False
   Icon            =   "Frm主窗体.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8040
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock winSock 
      Left            =   5520
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdateConnect 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimeToolTipText 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   690
   End
   Begin VB.PictureBox PicToolTipText 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4860
      ScaleHeight     =   225
      ScaleWidth      =   1485
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   1515
      Begin VB.Label LblToolTipText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "提示信息"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.PictureBox PicBackBitmap 
      AutoRedraw      =   -1  'True
      Height          =   585
      Left            =   360
      Picture         =   "Frm主窗体.frx":1CFA
      ScaleHeight     =   525
      ScaleWidth      =   1605
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.PictureBox PicRollUp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   390
      ScaleHeight     =   165
      ScaleWidth      =   2505
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   2505
      Begin VB.Image ImgRollUp 
         Height          =   240
         Index           =   0
         Left            =   1110
         Picture         =   "Frm主窗体.frx":DEB2
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   240
      End
   End
   Begin VB.PictureBox PicRollDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   390
      ScaleHeight     =   165
      ScaleWidth      =   2505
      TabIndex        =   17
      Top             =   1890
      Visible         =   0   'False
      Width           =   2505
      Begin VB.Image ImgRollDown 
         Height          =   240
         Index           =   0
         Left            =   1110
         Picture         =   "Frm主窗体.frx":DFFC
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   240
      End
   End
   Begin VB.Timer TimePass 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3780
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer SetZorder 
      Interval        =   10
      Left            =   3420
      Top             =   1950
   End
   Begin VB.PictureBox Pic任务条 
      AutoRedraw      =   -1  'True
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7965
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   8025
      Begin VB.PictureBox Pic任务 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   345
         Index           =   0
         Left            =   4650
         ScaleHeight     =   285
         ScaleWidth      =   2055
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   25
         Visible         =   0   'False
         Width           =   2115
         Begin VB.Label Lbl任务 
            Height          =   165
            Index           =   0
            Left            =   480
            TabIndex        =   11
            Top             =   90
            Width           =   1365
         End
         Begin VB.Image Img任务 
            Height          =   285
            Index           =   0
            Left            =   90
            Stretch         =   -1  'True
            Top             =   60
            Width           =   285
         End
      End
      Begin VB.PictureBox Pic分隔 
         AutoRedraw      =   -1  'True
         Height          =   315
         Left            =   990
         MousePointer    =   9  'Size W E
         ScaleHeight     =   255
         ScaleWidth      =   0
         TabIndex        =   5
         Top             =   60
         Width           =   60
      End
      Begin VB.PictureBox Pic开始 
         AutoRedraw      =   -1  'True
         Height          =   345
         Left            =   30
         ScaleHeight     =   285
         ScaleWidth      =   795
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   855
         Begin VB.PictureBox PicImg 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   405
            TabIndex        =   3
            Top             =   30
            Width           =   405
            Begin VB.Image Img开始 
               Height          =   240
               Left            =   60
               Picture         =   "Frm主窗体.frx":E146
               Stretch         =   -1  'True
               Top             =   30
               Width           =   270
            End
         End
         Begin VB.Label Lbl开始 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "开始"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   435
            TabIndex        =   4
            Top             =   90
            Width           =   360
         End
      End
      Begin ComctlLib.StatusBar Sbar 
         Height          =   375
         Left            =   6810
         TabIndex        =   12
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   1
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Object.Width           =   2117
               MinWidth        =   2117
               TextSave        =   "15:49"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Pic常用功能 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   570
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   345
         Begin VB.Image Img常用功能 
            Height          =   285
            Index           =   0
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Width           =   285
         End
      End
      Begin VB.Line LineRight 
         BorderColor     =   &H80000005&
         X1              =   990
         X2              =   990
         Y1              =   60
         Y2              =   360
      End
      Begin VB.Line LineLeft 
         BorderColor     =   &H80000003&
         X1              =   960
         X2              =   960
         Y1              =   60
         Y2              =   360
      End
   End
   Begin VB.Frame FraSplit 
      Height          =   30
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   2370
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.PictureBox PicBackDesktop 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   0
      Left            =   390
      ScaleHeight     =   555
      ScaleWidth      =   2505
      TabIndex        =   7
      Top             =   2550
      Visible         =   0   'False
      Width           =   2500
      Begin VB.PictureBox Pic菜单 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   1995
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   ";FileExit"
         Top             =   30
         Visible         =   0   'False
         Width           =   1995
         Begin VB.Label Lbl快键 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   930
            TabIndex        =   15
            Top             =   150
            Width           =   90
         End
         Begin VB.Image Img菜单 
            Height          =   480
            Index           =   0
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Lbl菜单 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "空"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   615
            TabIndex        =   9
            Top             =   150
            Width           =   180
         End
         Begin VB.Image Img菜单指示 
            Height          =   150
            Index           =   0
            Left            =   1200
            Picture         =   "Frm主窗体.frx":E710
            Stretch         =   -1  'True
            Top             =   210
            Visible         =   0   'False
            Width           =   120
         End
      End
   End
   Begin VB.PictureBox Pic标识 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   345
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   345
      Begin VB.Label Lbl标识 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   16
         Top             =   2460
         Width           =   90
      End
      Begin VB.Image Img标识 
         Height          =   3495
         Left            =   30
         Stretch         =   -1  'True
         Tag             =   "1"
         Top             =   120
         Width           =   285
      End
   End
   Begin MSComctlLib.ImageList ImgLvw 
      Left            =   3180
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer RefreshMenu 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1860
      Top             =   0
   End
   Begin MSComctlLib.ListView LvwList 
      Height          =   3795
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   6694
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   0
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   4035
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2223
            MinWidth        =   882
            Picture         =   "Frm主窗体.frx":E85A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuRightMenu 
      Caption         =   "右键菜单(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)"
      End
      Begin VB.Menu MnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolTester 
         Caption         =   "使用SQL速度测试工具(&U)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolIndividuation 
         Caption         =   "使用个性化设置(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolNotify 
         Caption         =   "消息通知(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolShowDisReport 
         Caption         =   "显示停用报表(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolDictonary 
         Caption         =   "字典管理工具(&D)"
      End
      Begin VB.Menu mnuToolMessage 
         Caption         =   "消息收发管理(&M)"
      End
      Begin VB.Menu mnuToolNotice 
         Caption         =   "提醒消息查阅(&T)"
      End
      Begin VB.Menu mnuToolStyle 
         Caption         =   "系统选项(&S)"
      End
      Begin VB.Menu mnuToolExcel 
         Caption         =   "启动&EXCEL报表"
      End
      Begin VB.Menu MnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolHistory 
         Caption         =   "清除历史记录(&H)"
      End
      Begin VB.Menu MnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolSetColor 
         Caption         =   "设置字体颜色(&O)"
      End
      Begin VB.Menu mnuToolSelBackBmp 
         Caption         =   "选择背景图片(&B)"
      End
      Begin VB.Menu mnuToolOutTool 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOutToolSet 
         Caption         =   "外接工具设置(&O)…"
      End
      Begin VB.Menu mnuToolOutToolList 
         Caption         =   "外接工具(&G)"
         Visible         =   0   'False
         Begin VB.Menu mnuToolOutToolExecute 
            Caption         =   "工具(&1)"
            Index           =   0
         End
      End
      Begin VB.Menu MnuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepairIndividuationClear 
         Caption         =   "清除本机界面异常(&L)"
      End
      Begin VB.Menu mnuRepairComponent 
         Caption         =   "检测安装部件(&C)"
      End
      Begin VB.Menu mnuRepairClientUpdate 
         Caption         =   "客户端修复(&U)"
      End
      Begin VB.Menu MnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuReg 
         Caption         =   "注销(&R)"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "FrmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnPress As Boolean                            '开始按钮是否处于压下状态
Private mintLevel As Integer                            '当前处于第几级
Private mintLast序号 As Integer                         '上一次的序号
Private mlngSelectModul As Long                         '选择的模块
Private mlngSelectUsual As Long                         '选择的常用功能
Private mblnMenuOpened As Boolean                       '所选择的菜单是否已经打开
Private mdblMenuWidth As Double                         '当前菜单的高度
Private mdblMenuHeight As Double                        '当前菜单的最大宽度
Private mblnFirst As Boolean                            '第一次启动成功
Private mblnShow As Boolean
Private mstrLastSelectCaption As String                 '上次所选择窗体的标题
Private mlngLastSelectIndex As Long                     '上次所选择窗体对应的任务条的索引
Private mFrmChildObj As Form                            '子窗体对象
Private mCurTime As Date                                '当前预升级时间检查点.
Private mblnAdjustPost As Boolean
Private mcllTemp As Collection
Private marrRoll(256) As String                         '--保存每级菜单允许滚动菜单项的索引
Private mstrTitle As String                             '产品标题
Private mblnHide As Boolean                             '是否显示本窗体
Private Const M_INT_RPTDISABLED As Integer = 242        '禁用报表图标
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Public mclsAppTool As New zl9AppTool.clsAppTool
Private mblnRemote As Boolean '是否开启远程

Public Property Get frmHide() As Boolean
    frmHide = mblnHide
End Property

Public Property Get ObjLogin() As Object
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    Set ObjLogin = gobjRelogin
End Property

Public Property Get mobjEmr() As Object
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    Set mobjEmr = gobjRelogin.EMR
End Property

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    
    mblnFirst = False
    mblnMenuOpened = False
    
    '--装入菜单组--
    If LoadLvw = False Then
        Unload Me
        Exit Sub
    End If
    Call LoadUsual
    
    '此段必须在创建同义词后(因消息通知属于ZlAppTool部件,执行其函数--GetUserInfo时出错)
    MnuToolIndividuation.Checked = IIf(Val(zlDatabase.GetPara("使用个性化风格")) = 0, False, True)
    mnuToolNotify.Checked = IIf(Val(zlDatabase.GetPara("接收邮件消息")) = 0, True, False)
    mnuToolTester.Checked = IIf(GetSetting("ZLSOFT", "公共全局", "SQLTest", 0) = 0, False, True)
    mnuToolShowDisReport.Checked = IIf(Val(zlDatabase.GetPara("显示停用报表")) = 0, False, True)
    mnuToolNotify_Click
    Call SetMainForm(Me)
    Call InitEvn
    
    '如果只有一可用模块,则打开
    On Error Resume Next
    With grsMenus
        .Filter = "模块<>0 And 报表=0"
        If Not .EOF Then
            If .RecordCount = 1 Then
                Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
            End If
        End If
        .Filter = 0
    End With
    
    If mblnShow = False Then MsgBox "显示背景图片时，发生错误！（恢复为缺省图片）", vbInformation, gstrSysName
    Call LoadOutTools(False)
    
    '启动消息服务平台客户端收发服务
    '------------------------------------------------------------------------------------------------------------------
    If ConnectMip(Me.hwnd) = True Then
        Set mclsMipModule = New zl9ComLib.clsMipModule
        Call mclsMipModule.InitMessage(0, 0, "")
        Call AddMipModule(mclsMipModule)
    End If
    
    '启动自动提醒服务
    mclsAppTool.CodeMan 0, 5, gcnOracle, Me, gstrDbUser
    If mblnHide Then Me.Hide '是外部调用，隐藏主窗体,by 陈东
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static StrPass As String                                '输入密码(Open zlReport.ReportMan )
    Dim objItem As ListItem, blnExist As Boolean
    
    If mblnPress And (KeyCode >= vbKeyA And KeyCode < vbKeyZ) Then
        Call FindMenu(KeyCode)
        Exit Sub
    End If
    
    '--隐藏菜单--
    If KeyCode = vbKeyEscape And mblnPress Then ShowMenu
    If KeyCode = vbKeyW And Shift = vbCtrlMask Then Pic开始_MouseDown 1, 0, 0, 0
    If KeyCode = vbKeyF4 And Shift = vbAltMask Then MnuExit_Click: Exit Sub
    
    TimePass.Enabled = False
    If KeyCode = vbKeyF12 And Shift = 7 Then
        StrPass = ""
        Exit Sub
    End If
    
    If KeyCode <> vbKeyReturn Then
        If InStr(1, "1234567890 ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyCode))) <> 0 Then StrPass = StrPass & UCase(Chr(KeyCode))
        
        If StrPass = "OPEN ZLREPORT REPORTMAN" Then
            If OwnerUser(gstrDbUser) Then
                StrPass = ""
                
                If FindWindow(vbNullString, "报表管理") <> 0 Then Exit Sub
                If MsgBox("您确定要运行自定义报表工具吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Call ExecuteFunc(0, "ZL9REPORT", 99999901)
                SetParent FindWindow(vbNullString, "报表管理"), Me.hwnd
            End If
        End If
    End If
    TimePass.Enabled = True
End Sub

Private Sub Form_Load()
    Dim IntKind As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim LngHdl As Long                                  '任务条句柄
    Dim intGrant As Integer
    mblnFirst = True
    mblnAdjustPost = False
    Dim strTitle As String, strTag As String
    
    On Error Resume Next
    '取系统窗体句柄
    LngHdl = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(LngHdl, 0, 0, 0, 0, 0, TOGGLE_HIDEWINDOW)
    
    '判断是否有权限使用消息收发功能
    Call CheckTools
    
    '使用该函数的目的就是用它来自动修改菜单(帮助及关于等)
    RestoreWinState Me
    
    '--检测是否为试用版（改菜单的标识图片）--
    IntKind = IIf(GetSetting("ZLSOFT", "注册信息", "Kind", "") = "试用", -1, 0)
    Set gcllCollMap = New Collection
    
    Me.WindowState = 2
    '设置基准菜单
    菜单基准.功能菜单 = 90000001
    菜单基准.窗口菜单 = 99990001
    菜单基准.其它功能菜单 = 99999901
    菜单基准.分隔菜单 = 99999999
    
    Call CheckWinVersion
    
    strTitle = zlRegInfo("产品标题")
    strTag = ""
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "旗舰版"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "专业版"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    mstrTitle = strTitle & IIf(strTag = "", "", "(" & strTag & ")")
    '传递数据库活动连接给打印部件
    IniPrintMode gcnOracle, gstrDbUser
    
    '首先判断会话中是否有消息服务器名称
    'select 参数值 from zloptions where 参数号 =17
    strSQL = "select 参数值 from zloptions where 参数号 =17"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断轮询服务器是否开启")
    If rsTemp.RecordCount = 1 Then
        If NVL(rsTemp!参数值) <> "" Then
            '开了轮询服务器,关闭TIME
            tmrUpdateConnect.Enabled = False
        Else
            '没开轮询服务器,使用TIME进行 预升级检查
            tmrUpdateConnect.Enabled = True
            tmrUpdateConnect.Interval = 30000
            mCurTime = Now
        End If
    Else
        '没开轮询服务器,使用TIME进行 预升级检查
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If
    
    '外部调用的处理,by 陈东
    mblnHide = False
    If gstrCommand <> "" Then Call DoCommand
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogInAfter
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 LogInAfter 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If

    '初始化监听
    InitWinsock
End Sub

Private Sub Form_Resize()
    Me.WindowState = 2
    With LvwList
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - 405
    End With
    With PicBackBitmap
        .Top = 0
        .Width = Me.ScaleWidth
        .Left = 0
        .Height = Me.ScaleHeight - 405
    End With
    
    With Pic任务条  'Windows 任务条
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth
        .Top = LvwList.Top + LvwList.Height
    End With
    With Sbar
        .Left = Pic任务条.Width - .Width - 50
    End With
    
    zlControl.PicShowFlat Pic任务条, 2, , taCenterAlign
    zlControl.PicShowFlat Pic开始, 2, , taCenterAlign
    zlControl.PicShowFlat Pic分隔, 2, , taCenterAlign
    
    gLngFormID = Me.hwnd
    Dim StrCaption As String
    StrCaption = mstrTitle ' zlProductTitle(GetUnitInfo("发行码"))
    
    '--设置窗体标题--
    Call SetWindowText(Me.hwnd, StrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim LngHdl As Long
    Dim blnCloaseWin As Boolean
    
    On Error Resume Next
    blnCloaseWin = Val(zlDatabase.GetPara("关闭Windows")) <> 0
    '取系统窗体句柄
    LngHdl = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(LngHdl, 0, 0, 0, 0, 0, TOGGLE_UNHIDEWINDOW)
    '清理外挂医保，以及业务窗体
    Call CloseChildWindows(Me)
    '清理消息对象
    Call DisConnectMip
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Call gobjRelogin.Dispose '需要先卸载对象
    Set gobjRelogin = Nothing
    SaveSetting "ZLSOFT", "公共全局", "SQLTest", 0
    '清除缓存的参数值
    zlDatabase.ClearParaCache
    Call ShutDown(blnCloaseWin)
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = 1 Then gcnOracle.Close
        Set gcnOracle = Nothing
    End If
    ReDim Preserve gobjCls(0)
    ReDim Preserve gstrObj(0)
End Sub

Private Sub ImgRollDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollDown_MouseMove(Index, Button, Shift, x, y)
End Sub

Private Sub ImgRollDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollDown_MouseUp(Index, Button, Shift, 0, 0)
End Sub

Private Sub ImgRollUp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollUp_MouseMove(Index, Button, Shift, x, y)
End Sub

Private Sub ImgRollUp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicRollUp_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub Img常用功能_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Pic常用功能_MouseDown(Index, Button, Shift, x, y)
End Sub

Private Sub Img常用功能_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Pic常用功能_MouseMove(Index, Button, Shift, x, y)
End Sub

Private Sub Img常用功能_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Pic常用功能_MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub Img任务_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic任务_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub LblToolTipText_Click()
    Call PicToolTipText_Click
End Sub

Private Sub LblToolTipText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicToolTipText_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Lbl任务_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic任务_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub LvwList_DblClick()
    Dim LngFindWindows As Long                          '目标窗体
    
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem.Key = "HELP" Then
        Shell "hh.exe  zl9start.chm", vbNormalFocus
        Exit Sub
    End If
    If LvwList.SelectedItem.Key = "EXIT" Then
        MnuExit_Click
        Exit Sub
    End If
    
    If LvwList.SelectedItem.Tag = -1 Then
        '--执行该模块--
        With grsMenus
            .MoveFirst
            .Find "编号='" & Mid(LvwList.SelectedItem.Key, 3) & "'"
            
            Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
        End With
    Else
        '--打开该模块--
        Call OpenWindow(Mid(LvwList.SelectedItem.Key, 3), LvwList.SelectedItem.Text)
    End If
End Sub

Public Function LoadLvw() As Boolean
    LoadLvw = False
    
    With grsMenus
        .Filter = "上级=0"
        LvwList.ListItems.Clear
        If .EOF Then
            MsgBox "你没有操作该系统的权限！"
            grsMenus.Filter = 0
            Exit Function
        End If
        
        On Error Resume Next
        Do While Not .EOF '为ImageList装入图标
            ImgLvw.ListImages.Add ImgLvw.ListImages.Count + 1, "K_" & ImgLvw.ListImages.Count + 1, GetPicDisp(!图标, False)
            .MoveNext
        Loop
        
        ImgLvw.ListImages.Add ImgLvw.ListImages.Count + 1, "HELP", GetPicDisp(-1)
        ImgLvw.ListImages.Add ImgLvw.ListImages.Count + 1, "EXIT", GetPicDisp(-3)
        
        Set LvwList.Icons = ImgLvw
        .MoveFirst
        Do While Not .EOF
            LvwList.ListItems.Add , "K_" & !编号, !标题, .AbsolutePosition
            LvwList.ListItems("K_" & !编号).Tag = IIf((!模块) = 0, 0, -1)
            .MoveNext
        Loop
        LvwList.ListItems.Add , "HELP", "帮助", "HELP"
        LvwList.ListItems.Add , "EXIT", "退出系统", "EXIT"
        
        .Filter = 0
    End With
    
    LoadLvw = True
End Function

Private Sub LvwList_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    LvwList.Drag 0
End Sub

Private Sub LvwList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then LvwList_DblClick
End Sub

Private Sub LvwList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnPress Then ShowMenu
    Call Find任务(-99999999)
    If Button = 2 Then PopupMenu MnuRightMenu, 2
End Sub

Private Sub LvwList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then LvwList.Drag 1
End Sub

Private Sub mclsMipModule_ConnectStateChanged(ByVal IsConnected As Boolean)
    '连接状态已经变化
    If IsConnected Then
        tmrUpdateConnect.Enabled = False
    Else
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If
End Sub

Private Sub mclsMipModule_OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara)
End Sub

Private Sub mclsMipModule_OpenReport(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara, True)
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMessageItemKey As String, ByVal strMessageConent As String)
    Select Case UCase(strMessageItemKey)
    '--------------------------------------------------------------------------------------------------------------
    Case "ZLHIS_PUB_005"            '产品升级通知
        Call gobjRelogin.UpdateClient
    End Select

End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuRepairClientUpdate_Click()
    If MsgBox("本操作将重新检测本机部件环境，对本机部件环境进行修复，对修复后的所有部件进行重新注册。你确认要进行客户端修复吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call gobjRelogin.UpdateClient(True)
    End If
End Sub

Public Sub mnuRepairComponent_Click()
    '--清空注册表[本机部件]--
    SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
    MsgBox "部件检测完毕，所有改动在重新登录后生效！", vbInformation, gstrSysName
End Sub

Private Sub mnuRepairIndividuationClear_Click()
    Dim strSQL As String, rsTmp As Recordset
    Dim strAnalyseComputer As String
    
    If MsgBox("本操作将清除ZLHIS相关的注册表参数，以及数据库中存储的本人、本机参数，产品相关功能将按参数缺省值运行，你确定要继续吗？", vbYesNo + vbDefaultButton2 + vbQuestion, "清除本机界面异常") = vbYes Then
        strSQL = "Select Distinct 部件 From zlPrograms Where 部件 Is Not Null"
        On Error GoTo ErrHand
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "清除本机界面异常")
        Do While Not rsTmp.EOF
            Call DelWinState(Me, rsTmp!部件 & "")
            rsTmp.MoveNext
        Loop
        strAnalyseComputer = OS.ComputerName
        strSQL = "Zl_zluserparas_Clear('" & gstrDbUser & "','" & strAnalyseComputer & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, gstrSysName)
        MsgBox "清除成功，请关闭程序重新进入，确认是否解决界面异常问题。", vbInformation, "清除本机界面异常"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuToolDictonary_Click()
    mclsAppTool.CodeMan 0, 1, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuToolExcel_Click()
    Dim ObjExcel As Object, strHaveSys As String
    
    If gstrUserName = "" Then
        MsgBox "请为操作员设置对应的用户后再使用本功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    strHaveSys = gobjRelogin.Systems
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Zl9Excel.ClsExcel")
    If Err <> 0 Then
        MsgBox "无法创建EXCEL部件，您将不能使用EXCEL报表！", vbInformation, gstrSysName
        Exit Sub
    End If
    Call ObjExcel.CodeMan(0, 0, gcnOracle, Me, gstrDbUser)
    Call ObjExcel.SetHaveSys(strHaveSys)
    Call ObjExcel.ExcelReportMain
    Set ObjExcel = Nothing
End Sub

Private Sub mnuToolHistory_Click()
    Call zlDatabase.SetPara("最近使用模块", "")
End Sub

Private Sub MnuToolIndividuation_Click()
    MnuToolIndividuation.Checked = MnuToolIndividuation.Checked Xor True
    Call zlDatabase.SetPara("使用个性化风格", IIf(MnuToolIndividuation.Checked, "1", "0"))
    SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser, "使用个性化风格", IIf(MnuToolIndividuation.Checked, "1", "0")
End Sub

Private Sub mnuToolMessage_Click()
    mclsAppTool.CodeMan 0, 2, gcnOracle, Me, gstrDbUser
End Sub

Private Sub MnuReg_Click()
    If MsgBox("你确定要注销吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call ReLogin
End Sub

Private Sub MnuExit_Click()
    Dim intStyle As Integer
    
    If Frm关闭.ShowMe(intStyle) Then
        If intStyle = 0 Then
            ReLogin
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub mnuToolNotice_Click()
    mclsAppTool.CodeMan 0, 6, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuToolNotify_Click()
    mnuToolNotify.Checked = Not mnuToolNotify.Checked
    Call zlDatabase.SetPara("接收邮件消息", IIf(mnuToolNotify.Checked, "1", "0"))
    mclsAppTool.CodeMan 0, 4, gcnOracle, Me, gstrDbUser, IIf(mnuToolNotify.Checked = True, "Open", "Close")
End Sub

Private Sub mnuToolOutToolExecute_Click(Index As Integer)
    '刘兴宏:2007/08/22
    '增加对外部工具的执行
    Call ExeCuteToolFile(mnuToolOutToolExecute(Index).Tag)
End Sub

Private Sub mnuToolOutToolSet_Click()
    Dim blnApply As Boolean
    '刘兴宏:2007/08/22
    '增加外部工具的设置
     Call frm工具设置.ShowEdit(Me, blnApply)
    If blnApply = False Then Exit Sub
    Call LoadOutTools(False)
End Sub

Private Sub mnuToolShowDisReport_Click()
    mnuToolShowDisReport.Checked = Not mnuToolShowDisReport.Checked
    Call zlDatabase.SetPara("显示停用报表", IIf(mnuToolShowDisReport.Checked, 1, 0))
End Sub

Private Sub mnuToolSelBackBmp_Click()
    Dim BlnShow As Boolean              '能否正常显示
    Dim StrPicPath As String            '背景图片路径
    '--供用户选择背景图片--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .Filter = "背景图片 (*.bmp;*.jpg)|*.bmp;*.jpg"
        .ShowOpen
        
        '用户选择图片,测试是否正常
        On Error Resume Next
        Err = 0
        BlnShow = False
        StrPicPath = .FileName
        Img标识.Picture = LoadPicture(StrPicPath)
        If Err <> 0 Then
            MsgBox "您所选择的图片文件不正常显示！", vbInformation, gstrSysName
        Else
            BlnShow = True
        End If
    End With
    
    PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '保存图片位置供下次提取
    Call zlDatabase.SetPara("zlWinBackPic", StrPicPath)
    '恢复原来设置的图片
    Img标识.Picture = LoadResPicture(101, 0) '菜单标识
ErrHand:
End Sub

Private Sub mnuToolSetColor_Click()
    '--供用户选择字体颜色--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .ShowColor
        
        LvwList.ForeColor = .Color
        
        '保存字体色供下次提取
        Call zlDatabase.SetPara("zlWinFontColor", .Color)
    End With
ErrHand:
End Sub

Public Sub mnuToolStyle_Click()
    mclsAppTool.CodeMan 0, 3, gcnOracle, Me, gstrDbUser, gstrMenuSys
    If Val(zlDatabase.GetPara("允许远程控制")) <> winSock.LocalPort Then
        Call InitWinsock
    End If
    If mclsAppTool.IsRestart Then
        mclsAppTool.IsRestart = False
        Call ReLogin
    Else
        '重新加载常用功能
        Call ShutUsual
        Pic分隔.Left = 990
        Call LoadUsual
    End If
End Sub

Private Sub mnuToolTester_Click()
    mnuToolTester.Checked = mnuToolTester.Checked Xor True
    SaveSetting "ZLSOFT", "公共全局", "SQLTest", IIf(mnuToolTester.Checked, 1, 0)
End Sub

Private Sub PicRollDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MouseOver As Boolean
    On Error Resume Next
    
    '--判断当前鼠标位置是否在菜单上--
    MouseOver = (0 <= x) And (x <= PicRollDown(Index).Width) And (0 <= y) And (y <= PicRollDown(Index).Height)
    If MouseOver Then
        Call ShutMenu(mintLevel)
        Call zlControl.PicShowFlat(PicRollDown(Index), -1, , taCenterAlign)
        Call SetCapture(PicRollDown(Index).hwnd)
    Else
        Call zlControl.PicShowFlat(PicRollDown(Index), 0, , taCenterAlign)
        Call ReleaseCapture
    End If
End Sub

Private Sub PicRollDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    Call RollUpMenu(PicRollDown(Index).Tag, 1)
End Sub

Private Sub PicRollUp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MouseOver As Boolean
    On Error Resume Next
    
    '--判断当前鼠标位置是否在菜单上--
    MouseOver = (0 <= x) And (x <= PicRollUp(Index).Width) And (0 <= y) And (y <= PicRollUp(Index).Height)
    If MouseOver Then
        Call ShutMenu(mintLevel)
        Call zlControl.PicShowFlat(PicRollUp(Index), -1, , taCenterAlign)
        Call SetCapture(PicRollUp(Index).hwnd)
    Else
        Call zlControl.PicShowFlat(PicRollUp(Index), 0, , taCenterAlign)
        Call ReleaseCapture
    End If
End Sub

Private Sub PicRollUp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    Call RollUpMenu(PicRollUp(Index).Tag, 2)
End Sub

Private Sub PicToolTipText_Click()
    PicToolTipText.Visible = False
    TimeToolTipText.Enabled = False
End Sub

Private Sub PicToolTipText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicToolTipText_Click
End Sub

Private Sub Pic菜单_DblClick(Index As Integer)
    '--存在下级菜单则打开--
    Call PicToolTipText_Click
    If Img菜单指示(Index).Visible Then
        '--打开该模块--
        Call OpenWindow(Img菜单(Index).Tag, Mid(Lbl菜单(Index).Caption, 1, IIf(InStr(1, Lbl菜单(Index).Caption, "(") <> 0, Len(Lbl菜单(Index).Caption) - 3, Len(Lbl菜单(Index).Caption))))
        
        If mblnPress Then ShowMenu
    End If
End Sub

Private Sub Pic菜单_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ReleaseCapture
    mlngSelectModul = 0
End Sub

Private Sub Pic常用功能_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '设置为凹下
    If Button <> 1 Then Exit Sub
    mlngSelectUsual = Index
    Call zlControl.PicShowFlat(Pic常用功能(Index), -2, , taCenterAlign)
End Sub

Private Sub Pic常用功能_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '设置为凸出
    Dim MouseOver As Boolean
    '--判断当前鼠标位置是否在菜单上--
    
    If Button = 1 Then Exit Sub
    MouseOver = (0 <= x) And (x <= Pic常用功能(Index).Width) And (0 <= y) And (y <= Pic常用功能(Index).Height)
    If MouseOver Then
        Call zlControl.PicShowFlat(Pic常用功能(Index), 2, , taCenterAlign)
        Call SetCapture(Pic常用功能(Index).hwnd)
        Call ShowToolTipText(Pic常用功能(Index))
    Else
        Call zlControl.PicShowFlat(Pic常用功能(Index), 0, , taCenterAlign)
        Call ReleaseCapture
        Call ShowToolTipText(Pic常用功能(Index), False)
    End If
End Sub

Private Sub Pic常用功能_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lng系统 As Long, str部件 As String, lng模块 As Long
    If Button <> 1 Then Exit Sub
    If Not (Index = mlngSelectUsual) Then Exit Sub
    
    mlngSelectUsual = 0
    '先分解参数
    str部件 = ""
    lng系统 = Split(Pic常用功能(Index).Tag, "♂")(0)
    lng模块 = Split(Pic常用功能(Index).Tag, "♂")(1)
    
    grsMenus.Filter = "系统=" & lng系统 & " And 模块=" & lng模块
    If grsMenus.RecordCount <> 0 Then str部件 = IIf(IsNull(grsMenus!部件), "", grsMenus!部件)
    grsMenus.Filter = 0
    If str部件 = "" Then Exit Sub
    
    '运行该模块
    Call zlControl.PicShowFlat(Pic常用功能(Index), 0, , taCenterAlign)
    Call ExecuteFunc(lng系统, str部件, lng模块)
End Sub

Private Sub Pic任务_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim IntChange As Integer
    Dim LngActiveWindow As Long
    '用于保存集合中的数据
    Dim IntIndex As Integer, LngHdl As Long, intStyle As Integer
    
    '设置任务条模式
    For IntChange = 1 To gcllCollMap.Count
        IntIndex = gcllCollMap("K_" & IntChange)(0)
        LngHdl = gcllCollMap("K_" & IntChange)(1)
        intStyle = gcllCollMap("K_" & IntChange)(2)
        
        gcllCollMap.Remove "K_" & IntChange
        
        If IntIndex <> Index Then
            gcllCollMap.Add Array(IntIndex, LngHdl, 0), "K_" & IntChange
            Call zlControl.PicShowFlat(Pic任务(IntIndex), 2, , taCenterAlign)
        Else
            '激活当前窗体
            If IsIconic(LngHdl) Then
                gcllCollMap.Add Array(IntIndex, LngHdl, 1), "K_" & IntChange
                Call ShowWindow(LngHdl, 9)            '还原指定窗体为原大小
                Call zlControl.PicShowFlat(Pic任务(Index), -2, , taCenterAlign)
            Else
                If intStyle = 0 Then
                    gcllCollMap.Add Array(IntIndex, LngHdl, 1), "K_" & IntChange
                    Call zlControl.PicShowFlat(Pic任务(Index), -2, , taCenterAlign)
                Else
                    gcllCollMap.Add Array(IntIndex, LngHdl, 0), "K_" & IntChange
                    Call CloseWindow(LngHdl)
                    Call zlControl.PicShowFlat(Pic任务(Index), 2, , taCenterAlign)
                End If
            End If
            If Not IsIconic(LngHdl) Then Call SetActiveWindow(LngHdl)
        End If
    Next
End Sub

Private Sub Img菜单_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic菜单_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub Img菜单_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic菜单_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub Img菜单指示_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic菜单_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub Img菜单指示_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic菜单_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub Img开始_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic开始_MouseDown Button, Shift, x, y
End Sub

Private Sub Lbl菜单_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic菜单_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub Lbl菜单_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic菜单_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub Lbl开始_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic开始_MouseDown Button, Shift, x, y
End Sub

Private Sub PicBackDesktop_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mintLevel = Index + 1
End Sub

Private Sub PicImg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic开始_MouseDown Button, Shift, x, y
End Sub

Private Sub Pic菜单_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strCode As String                               '模块编号
    Dim LngFindWindows As Long                          '目标窗体
    
    If Button = 1 Then
        Call PicToolTipText_Click
        If Img菜单指示(Index).Visible Then
            '--存在下级菜单--
            If mblnMenuOpened = False Then
                LoadMenu mintLevel, Index
                mblnMenuOpened = True
            End If
            
        Else
            '--执行模块功能--
            strCode = Img菜单(Index).Tag
            If mblnPress Then ShowMenu

            Select Case Index
            Case "9000"                             '帮助
                Shell "hh.exe  zl9start.chm", vbNormalFocus
            Case "9001"                             '注销
                MnuReg_Click
            Case "9002"                             '退出
                MnuExit_Click
            Case "9100"                             '字典管理
                mnuToolDictonary_Click
            Case "9101"                             '消息收发
                mnuToolMessage_Click
            Case "9102"
                mnuRepairComponent_Click
            Case "9103"
                mnuToolStyle_Click
            Case "9104"
                mnuToolExcel_Click
            Case "9105"                             '提醒消息
                mnuToolNotice_Click
            Case Is >= 9300 And Index <= 9500
                '刘兴宏:目前暂定200个工具
                '小于零的情况下,表示进行的外部工具调用
                If Index = 9301 Then
                    Dim blnApply As Boolean
                    '刘兴宏:2007/08/22
                    '增加外部工具的设置
                     Call frm工具设置.ShowEdit(Me, blnApply)
                    If blnApply = False Then Exit Sub
                    Call LoadOutTools(False)
                Else
                    Err = 0: On Error Resume Next
                    Call ExeCuteToolFile(mcllTemp("K" & Index))
                End If
            Case Else                               '读菜单
                    DoEvents
                    With grsMenus
                        .MoveFirst
                        .Find "编号='" & strCode & "'"
                        Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
                    End With
 
            End Select
        End If
    End If
End Sub

Private Sub Pic菜单_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MouseOver As Boolean
    
    '--将滚动条设置为平面
    On Error Resume Next
    Call zlControl.PicShowFlat(PicRollUp(mintLevel - 1), 0, , taCenterAlign)
    Call zlControl.PicShowFlat(PicRollDown(mintLevel - 1), 0, , taCenterAlign)
    
    '--判断当前鼠标位置是否在菜单上--
    MouseOver = (0 <= x) And (x <= Pic菜单(Index).Width) And (0 <= y) And (y <= Pic菜单(Index).Height)
    If MouseOver Then
        If mlngSelectModul = Index Then Exit Sub
        Pic菜单(Index).BackColor = &H8000000D
        Lbl菜单(Index).ForeColor = &H80000005
        Call SetCapture(Pic菜单(Index).hwnd)
        If mlngSelectModul <> Index Then mblnMenuOpened = False
        mlngSelectModul = Index
        If Button <> 88 Then RefreshMenu.Enabled = True
        Call PicBackDesktop_MouseMove(Pic菜单(Index).Container.Index, Button, Shift, x, y)
        Call ShowToolTipText(Pic菜单(Index))
    Else
        Pic菜单(Index).BackColor = &H8000000F
        Lbl菜单(Index).ForeColor = &H80000008
        Call ReleaseCapture
        mlngSelectModul = 0
        RefreshMenu.Enabled = False
        Call ShowToolTipText(Pic菜单(Index), False)
    End If
End Sub

Private Sub Pic分隔_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnPress Then ShowMenu
End Sub

Private Sub Pic分隔_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim DblTotalWidth As Double
    
    If Button = 1 Then
        With Pic分隔
            
            DblTotalWidth = 2000
            DblTotalWidth = DblTotalWidth * (Pic任务.Count - 1)
            
            If DblTotalWidth > Sbar.Left - Pic分隔.Left - x - Pic分隔.Width - 1000 Then
                '--任务宽度和大于可容纳的空间--
                DblTotalWidth = ((Sbar.Left - Pic分隔.Left - x - Pic分隔.Width - 100) / IIf(Pic任务.Count - 1 = 0, 1, Pic任务.Count - 1)) - 50
                If DblTotalWidth > 2000 Then DblTotalWidth = 2000
            Else
                DblTotalWidth = 2000
            End If
            If DblTotalWidth < 800 Then Exit Sub
            
            If .Left + x > Sbar.Left - 3000 Then .Left = Sbar.Left - 3000: Pic分隔_MouseMove Button, Shift, 0, 0: Exit Sub
            If .Left + x < Pic常用功能(Pic常用功能.Count - 1).Left + Pic常用功能(Pic常用功能.Count - 1).Width + 100 Then .Left = Pic常用功能(Pic常用功能.Count - 1).Left + Pic常用功能(Pic常用功能.Count - 1).Width + 100: Pic分隔_MouseMove Button, Shift, 0, 0: Exit Sub
            
            .Move .Left + x
        End With
        
        With LineRight
            .X1 = Pic分隔.Left - 25
            .X2 = .X1
        End With
        
        With LineLeft
            .X1 = LineRight.X1 - 25
            .X2 = .X1
        End With
        
        Call AdjustPost
    End If
End Sub

Private Sub Pic分隔_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call AdjustPost
End Sub

Private Sub Pic开始_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnFirst Then Exit Sub
    mintLevel = 1
    If Button = 1 Then Call ShowMenu
End Sub

Public Function ShowMenu()
    Dim IntDelLevel As Integer
    '--按下开始按钮,则显示或隐藏菜单,调整相关属性--
    
    mblnPress = mblnPress Xor True
    Call zlControl.PicShowFlat(Pic开始, IIf(mblnPress, -2, 2), , taCenterAlign)
    
    '--显示或隐藏所有菜单--
    If mblnPress Then
        Call LoadMenu(-1)
    Else
        Call ShutMenu
        mlngSelectModul = 0
    End If
    Pic标识.Visible = mblnPress
End Function

Private Sub Pic任务条_Click()
    If mblnPress Then ShowMenu
    Call Find任务(-99999999)
End Sub

Private Sub Pic任务条_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Pic任务条.ZOrder 0
End Sub

Private Sub RefreshMenu_Timer()
    Dim Index As Integer
    
    On Error Resume Next
    RefreshMenu.Enabled = False
    If mlngSelectModul = 0 Then Exit Sub
    
    If Img菜单指示(mlngSelectModul).Visible = False Then
        Call ShutMenu(mintLevel)
        Exit Sub
    End If
    
    Index = mlngSelectModul
    Pic菜单_MouseMove Index, 88, 0, 1, 1
    Pic菜单_MouseDown Index, 1, 0, 1, 1
End Sub

Private Sub LoadMenu(ByVal IntState As Integer, Optional ByVal IntIndex As Integer = 0)
    Dim DblWidthTmp As Double, lngIndexThis As Long, IntDelLevel As Integer
    Dim str系统 As String, str序号 As String
    Dim int系统_Cur As Integer, int系统_Max As Integer
    Dim int序号_Cur As Integer, int序号_Max As Integer
    Dim arr系统, arr序号
    Dim blnRight As Boolean
    '--先卸载下级菜单，再装入指定菜单，并设置其级数--
    
    With grsMenus
        If IntIndex = 0 Then
            .Filter = "上级=0"
        Else
            .Filter = "上级='" & Img菜单(IntIndex).Tag & "'"
        End If
        If .EOF And Img菜单(IntIndex).Tag <> "9003" And Img菜单(IntIndex).Tag <> "9004" Then Exit Sub
        
        On Error Resume Next
        '--先卸载再装入菜单--
        Call ShutMenu(mintLevel)
        Load PicBackDesktop(mintLevel)
        
        mintLast序号 = 0
        mdblMenuHeight = 0
        mdblMenuWidth = 0
        
        If Img菜单(IntIndex).Tag <> "9003" And Img菜单(IntIndex).Tag <> "9004" Then             '不是工具菜单及历史使用记录
            '--装入系统菜单--
            Do While Not .EOF  '统计菜单的最大宽度
                DblWidthTmp = 95 * PicBackDesktop(mintLevel).Font.Size / 9 * LenB(StrConv(!标题, vbFromUnicode)) + 800
                If mdblMenuWidth < DblWidthTmp Then mdblMenuWidth = DblWidthTmp
                .MoveNext
            Loop
            mdblMenuWidth = mdblMenuWidth + Img菜单(0).Width + 400 + Img菜单指示(0).Width
        
            .MoveFirst
            marrRoll(mintLevel) = ""
            Do While Not .EOF
                lngIndexThis = Img菜单.Count
                marrRoll(mintLevel) = marrRoll(mintLevel) & IIf(marrRoll(mintLevel) = "", "", ",") & lngIndexThis
                If !报表 = 1 And Val(!是否停用) = 1 Then
                    If mnuToolShowDisReport.Checked Then
                        Call SetMenuState(lngIndexThis, !编号, M_INT_RPTDISABLED, !标题, IIf(IsNull(!快键), "", !快键), 1, IIf(!模块 = 0, False, True), IIf(IsNull(!说明), "", !说明))
                    End If
                Else
                    Call SetMenuState(lngIndexThis, !编号, !图标, !标题, IIf(IsNull(!快键), "", !快键), 1, IIf(!模块 = 0, False, True), IIf(IsNull(!说明), "", !说明))
                End If
                .MoveNext
            Loop
        Else
            DblWidthTmp = 95 * PicBackDesktop(mintLevel).Font.Size / 9 * LenB(StrConv("小宝小宝小宝小宝", vbFromUnicode)) + 800
            mdblMenuWidth = DblWidthTmp
            mdblMenuWidth = DblWidthTmp + Img菜单(0).Width + 400 + Img菜单指示(0).Width
            If Img菜单(IntIndex).Tag = "9003" Then
                '--装入工具菜单--
                '装入字典管理菜单
                If mnuToolDictonary.Visible Then Call SetMenuState(9100, 9100, -5, "字典管理工具", "D", 1, True)
                '装入消息收发菜单
                If mnuToolMessage.Visible Then Call SetMenuState(9101, 9101, -5, "消息收发管理", "M", 1, True)
                '装入提醒消息查阅菜单
                If mnuToolNotice.Visible Then Call SetMenuState(9105, 9105, -5, "提醒消息查阅", "R", 1, True)
                '装入界面风格选择
                If mnuToolStyle.Visible Then Call SetMenuState(9103, 9103, -5, "系统选项", "S", 1, True)
                '装入启动EXCEL报表
                If mnuToolExcel.Visible Then Call SetMenuState(9104, 9104, -5, "启动EXCEL报表", "E", 1, True)
                '装入检测安装部件
                Call SetMenuState(9102, 9102, -5, "检测安装部件", "C", 1, True)
                
                '装入工具菜单
                Call LoadOutTools(True)
                
            Else
                '--装入历史使用记录菜单--
                Call LoadHistory
            End If
        End If
        
        If IntState = -1 Then
            Load FraSplit(1)
            '装入工具菜单
            Call SetMenuState(9003, 9003, -4, "常规工具", "T", 1, False)
            If Trim(zlDatabase.GetPara("最近使用模块")) <> "" Then
                '装入历史使用记录菜单
                Call SetMenuState(9004, 9004, -4, "历史记录", "O", 1, False)
            End If
            '装入帮助菜单
            Call SetMenuState(9000, 9000, -1, "帮助", "H", 1, True)
            Load FraSplit(2)
            '装入注销菜单
            Call SetMenuState(9001, 9001, -2, "注销", "R", 1, True)
            '装入退出菜单
            Call SetMenuState(9002, 9002, -3, "退出", "X", 1, True)
        End If
        
        '--调整装载容器的位置--
        With PicBackDesktop(mintLevel)
            If mintLevel = 1 Then
                .Left = Pic标识.Left + Pic标识.Width - 20
            Else
                '对.Left进行检查
                blnRight = True
                If mintLevel >= 3 Then
                    blnRight = (PicBackDesktop(mintLevel - 2).Left < PicBackDesktop(mintLevel - 1).Left)
                End If
                If blnRight Then
                    .Left = PicBackDesktop(mintLevel - 1).Left + PicBackDesktop(mintLevel - 1).Width - 100
                    If .Left + mdblMenuWidth > Me.ScaleWidth Then
                        .Left = PicBackDesktop(mintLevel - 1).Left - mdblMenuWidth + 100
                        blnRight = blnRight Xor True
                    End If
                Else
                    .Left = PicBackDesktop(mintLevel - 1).Left - mdblMenuWidth + 100
                    If .Left < 0 Then
                        .Left = PicBackDesktop(mintLevel - 1).Left + PicBackDesktop(mintLevel - 1).Width - 100
                        blnRight = blnRight Xor True
                    End If
                End If
            End If
            .Width = mdblMenuWidth + 50
            .Height = mdblMenuHeight + IIf(mintLevel = 1, 150, 90)
            
            If mintLevel = 1 Then
                '--如果是第一级，则赋原值--
                Pic标识.Height = .Height - 50
                Img标识.Height = Pic标识.Height
                .Top = Pic任务条.Top - .Height
            Else
                '--如果下级菜单，缺省Top为上级菜单的高度--
                Dim DblTop
                DblTop = PicBackDesktop(mintLevel - 1).Top + Pic菜单(IntIndex).Top
                '对.Top进行检查
                If Pic任务条.Top - mdblMenuHeight - 50 < DblTop Then
                    '如果超高(向上显示)
                    .Top = Pic任务条.Top - mdblMenuHeight - 50
                Else
                    .Top = DblTop - 50
                End If
            End If
            .Tag = mintLevel
            .Visible = IIf(.Height > 100, True, False)
            .ZOrder 0
        End With
        
        If mintLevel = 1 Then
            '--如果是第一级菜单，则调整标识--
            With Pic标识
                .Height = mdblMenuHeight + 150
                .Top = Pic任务条.Top - .Height
                .ZOrder 0
            End With
    
            With Img标识
                .Top = Pic标识.Height - .Height
            End With
            
            '设置软件名称
            With Lbl标识
                .AutoSize = True
                .Caption = mstrTitle 'zlProductTitle(GetUnitInfo("发行码"))
                .AutoSize = False
                .Height = .Width
                .Width = 200
                .Left = Img标识.Left + 80
                .Top = Pic标识.Height - .Height - 100
                .ForeColor = IIf(GetSetting("ZLSOFT", "注册信息", "Kind", "") = "试用", &HFF, &HFFFFFF)
            End With
            
            zlControl.PicShowFlat Pic标识, 2, , taCenterAlign
        End If
        
        Call AdjustMenu(mintLevel)
        Call zlControl.PicShowFlat(PicBackDesktop(mintLevel), 2, , taCenterAlign)
        PicBackDesktop(mintLevel).ZOrder 0
    End With
    grsMenus.Filter = 0
End Sub

Private Function SetMenuState(ByVal lngCurID As Long, ByVal strCode As String, ByVal LngIcon As Long, ByVal StrCaption As String, _
ByVal BytLink As String, ByVal intType As Integer, Optional ByVal BlnEndMenu As Boolean = True, Optional ByVal strNote As String = "")
    '--调整各菜单条的宽度等属性--
    'IntType:是否允许滚动 1-允许
    
    Load Img菜单(lngCurID)
    Load Img菜单指示(lngCurID)
    Load Lbl菜单(lngCurID)
    Load Pic菜单(lngCurID)
    Load Lbl快键(lngCurID)
        
    With Img菜单(lngCurID)
        .Left = 100
        .Top = -10
        .Tag = strCode
        Set .Container = Pic菜单(lngCurID)
        .Visible = True
        .Picture = GetPicDisp(LngIcon, BlnEndMenu) '菜单标识
    End With
    
    With Lbl菜单(lngCurID)
        .Left = Img菜单(lngCurID).Left + Img菜单(lngCurID).Width + 200
        Set .Container = Pic菜单(lngCurID)
        .Visible = True
        .Caption = StrCaption
        If BytLink <> "" Then .Caption = .Caption & "(" & BytLink & ")"
    End With
    
    With Lbl快键(lngCurID)
        .Left = Lbl菜单(lngCurID).Left + Lbl菜单(lngCurID).Width - 180
        Set .Container = Pic菜单(lngCurID)
        .Caption = BytLink
        .Visible = (BytLink <> "")
    End With
    
    If LngIcon < 0 Then
        Select Case lngCurID
        Case 9003
            With FraSplit(1)
                .Visible = True
                Set .Container = PicBackDesktop(mintLevel)
                .Left = 0
                .Top = Pic菜单(mintLast序号).Top + Pic菜单(mintLast序号).Height
                .Width = Pic菜单(mintLast序号).Width
            End With
        Case 9001
            With FraSplit(2)
                .Visible = True
                Set .Container = PicBackDesktop(mintLevel)
                .Left = 0
                .Top = Pic菜单(mintLast序号).Top + Pic菜单(mintLast序号).Height
                .Width = Pic菜单(mintLast序号).Width
            End With
        End Select
    End If
    
    With Pic菜单(lngCurID)
        .Left = 25
        If mintLast序号 = 0 Then
            .Top = 50
        Else
            If LngIcon < 0 Then
                Select Case lngCurID
                Case 9003
                    .Top = FraSplit(1).Top + FraSplit(1).Height
                Case 9001
                    .Top = FraSplit(2).Top + FraSplit(2).Height
                Case Else
                    .Top = Pic菜单(mintLast序号).Top + Pic菜单(mintLast序号).Height
                End Select
            Else
                .Top = Pic菜单(mintLast序号).Top + Pic菜单(mintLast序号).Height
            End If
        End If
        .Tag = mintLevel
        Set .Container = PicBackDesktop(mintLevel)
        .Visible = True
        .Width = mdblMenuWidth - 50
        mdblMenuHeight = mdblMenuHeight + .Height
    End With
    
    With Img菜单指示(lngCurID)
        .Left = Pic菜单(lngCurID).Width - .Width - 50
        .Top = (Pic菜单(lngCurID).Height - .Height) / 2
        Set .Container = Pic菜单(lngCurID)
        .Visible = BlnEndMenu Xor True
    End With
    
    With Lbl菜单(lngCurID)
        .Width = Img菜单指示(lngCurID).Left - .Left - 100
    End With
    
    If mintLast序号 <> lngCurID Then mintLast序号 = lngCurID
    
    '设置提示信息
    Call SetToolTipText(Img菜单, lngCurID, 0, strNote)
    Call SetToolTipText(Lbl菜单, lngCurID, 0, strNote)
    Call SetToolTipText(Lbl快键, lngCurID, 0, strNote)
    Call SetToolTipText(Pic菜单, lngCurID, 0, strNote)
    Call SetToolTipText(Img菜单指示, lngCurID, 0, strNote)
End Function

Public Sub ShutMenu(Optional ByVal Level As Integer = 0)
    Dim ObjShut As Object, LngUnloadObjs As Long
    Dim IntDelLevel As Integer
    
    On Error Resume Next
    LngUnloadObjs = 0
    RefreshMenu.Enabled = False
    For Each ObjShut In Me.Controls '删除控件
        Err = 0
        Select Case TypeName(ObjShut)
            Case "Image", "Label", "PictureBox"
                If ObjShut.Index <> 0 And Not (ObjShut.Name Like "*常用功能") Then
                    If Err = 0 Then
                        If Val(ObjShut.Container.Tag) >= Level Then
                            Unload ObjShut
                        End If
                    End If
                End If
            Case "Frame"
                If Level <= 1 Then
                    If ObjShut.Index <> 0 And Not (ObjShut.Name Like "*常用功能") Then
                        If Err = 0 Then
                            If Val(ObjShut.Container.Tag) >= Level Then
                                Unload ObjShut
                            End If
                        End If
                    End If
                End If
        End Select
    Next

    For Each ObjShut In Me.Controls '因容器上次无法删除，再次执行
        Err = 0
        Select Case TypeName(ObjShut)
            Case "PictureBox"
                If ObjShut.Index <> 0 And Not (ObjShut.Name Like "*常用功能") Then
                    If Err = 0 Then
                        If Val(ObjShut.Tag) >= Level Then
                            ObjShut.Visible = False
                            Unload ObjShut
                        End If
                    End If
                End If
        End Select
    Next
    
    For IntDelLevel = mintLevel To PicBackDesktop.Count - 1
        PicBackDesktop(IntDelLevel).Visible = False
        Unload PicBackDesktop(IntDelLevel)
    Next
End Sub

Private Sub FindMenu(ByVal IntKey As Integer)
    Dim ObjShut As Control
    
    On Error Resume Next
    '--在当前级数中查找指定快捷键的菜单--
    For Each ObjShut In Me.Controls '因容器上次无法删除，再次执行
        Err = 0
        If TypeName(ObjShut) = "PictureBox" Then
            If ObjShut.Index <> 0 Then
                If Err = 0 Then
                    If Val(ObjShut.Container.Index) = mintLevel Then
                        If Err = 0 Then
                            If Lbl快键(ObjShut.Index).Caption = UCase(Chr(IntKey)) Then
                                '设置菜单为选中状态并执行
                                RefreshMenu.Enabled = False
                                Pic菜单_MouseMove ObjShut.Index, 88, 0, 0, 0
                                Pic菜单_MouseDown ObjShut.Index, 1, 0, 0, 0
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    
End Sub

Public Sub OpenWindow(ByVal strCode As String, ByVal StrCaption As String)
    Dim FrmMainOpen As FrmExecute, FrmTest As Form
    Dim IntChange As Integer, LngHdl As Long
    Dim strBuffer As String * 256, IntBuffer As Integer
    
    IntBuffer = 255
    '设置任务条模式
    For IntChange = 1 To gcllCollMap.Count
        LngHdl = gcllCollMap("K_" & IntChange)(1)
        Call GetWindowText(LngHdl, strBuffer, IntBuffer)
        If StrCaption = Trim(Replace(strBuffer, Chr(0), "")) Then Exit For
    Next
    
    If StrCaption = Trim(Replace(strBuffer, Chr(0), "")) Then
        '设置任务条模式
        For IntChange = 1 To gcllCollMap.Count
            If gcllCollMap("K_" & IntChange)(1) <> LngHdl Then
                Call zlControl.PicShowFlat(Pic任务(gcllCollMap("K_" & IntChange)(0)), 2, , taCenterAlign)
            Else
                Call zlControl.PicShowFlat(Pic任务(gcllCollMap("K_" & IntChange)(0)), -2, , taCenterAlign)
            End If
        Next
        
        '激活当前窗体
        If IsIconic(LngHdl) Then Call ShowWindow(LngHdl, 9)            '还原指定窗体为原大小
        Call SetActiveWindow(LngHdl)
    Else
        Set FrmMainOpen = New FrmExecute
        With FrmMainOpen
            .窗体名称 = StrCaption
            .Str编号 = strCode
            Set .mrsMenus = grsMenus.Clone
            Call SetParent(.hwnd, gLngFormID)
            .Show 0
        End With
    End If
End Sub

Public Function Show任务(ByVal ChildObj As Object, Optional ByVal strCode As String = "", Optional ByVal StrCaption As String = "")
    Dim LngIcon As Long
    Dim IntIndex As Integer
    Dim DblTotalWidth As Double
    
    If grsMenus.State = 0 Then Exit Function
    If grsMenus.EOF Then Exit Function
    With grsMenus
        .MoveFirst
        .Find "标题='" & ChildObj.Caption & "'"
        If .EOF Then
            .MoveFirst
            
            '如果属于工具
            If ChildObj.Caption = "" Then Exit Function
            If InStr(1, "自定义报表管理,字典管理工具,消息收发管理", ChildObj.Caption) <> 0 Then GoTo Normal
            Exit Function
        End If
    End With
    
Normal:                                                     '正常进入
    Set mFrmChildObj = ChildObj
    Call SetParent(ChildObj.hwnd, Me.hwnd)
    
    If StrCaption = "" Then StrCaption = mFrmChildObj.Caption
    If gcllCollMap.Count = 0 Then
        IntIndex = 1
    Else
        IntIndex = gcllCollMap(gcllCollMap.Count)(0) + 1
    End If
    Load Pic任务(IntIndex)
    Load Lbl任务(IntIndex)
    Load Img任务(IntIndex)
    
    With Pic任务(IntIndex)
        .Tag = -IntIndex
        .Visible = True
'        Set .Container = Pic任务条
    End With
    With Lbl任务(IntIndex)
        .Tag = -IntIndex
        .Caption = StrCaption
        Set .Container = Pic任务(IntIndex)
        .Visible = True
    End With
    With Img任务(IntIndex)
        .Width = 240
        .Height = 240
        .Tag = -IntIndex
        .Picture = ChildObj.Icon
        Set .Container = Pic任务(IntIndex)
        .Visible = True
    End With
    '参数说明          索引       窗体句柄       是否激活
    gcllCollMap.Add Array(IntIndex, mFrmChildObj.hwnd, 1), "K_" & gcllCollMap.Count + 1
    Call AdjustPost
    
    If grsMenus!模块 <> 0 Then Call AddHistory(grsMenus!系统 & "," & grsMenus!模块)
End Function

Public Sub Shut任务(ByVal ObjFrm As Object)
    Dim IntChange As Integer, IntDelete As Integer
    On Error Resume Next
    
    '找到该任务
    IntDelete = 0
    For IntChange = 1 To gcllCollMap.Count
        If gcllCollMap(IntChange)(1) = ObjFrm.hwnd Then
            '删除该任务
            IntDelete = IntChange
            Call Find任务(gcllCollMap(IntChange)(0), True)
            Exit For
        End If
    Next
    
    If IntDelete = 0 Then Exit Sub
    '依次修改后继任务
    For IntChange = IntDelete To gcllCollMap.Count - 1
        gcllCollMap.Remove "K_" & IntChange
        gcllCollMap.Add gcllCollMap("K_" & IntChange + 1), "K_" & IntChange
    Next
    gcllCollMap.Remove "K_" & gcllCollMap.Count
    
    Call AdjustPost
End Sub

Public Sub Find任务(ByVal Index As Long, Optional BlnDel As Boolean = False, Optional BlnState As Boolean = False)
    Dim ObjShut As Object
    
    On Error Resume Next
    
    Call AdjustPost
    For Each ObjShut In Me.Controls '删除控件
        Err = 0
        Select Case TypeName(ObjShut)
            Case "Image", "Label", "PictureBox"
                If ObjShut.Index <> 0 Then
                    If Err = 0 Then
                        If ObjShut.Tag < 0 Then
                            If BlnDel And ObjShut.Tag = -Index Then
                                Unload ObjShut
                                Call AdjustPost
                            End If
                        End If
                    End If
                End If
        End Select
    Next
    For Each ObjShut In Me.Controls '删除控件
        Err = 0
        Select Case TypeName(ObjShut)
            Case "PictureBox"
                If ObjShut.Index <> 0 Then
                    If Err = 0 Then
                        If ObjShut.Tag < 0 Then
                            If BlnDel And ObjShut.Tag = -Index Then
                                Unload ObjShut
                                Call AdjustPost
                            ElseIf BlnState And ObjShut.Tag = -Index Then
                                Call zlControl.PicShowFlat(ObjShut, -2, , taCenterAlign)
                            Else
                                Call zlControl.PicShowFlat(ObjShut, 2, , taCenterAlign)
                            End If
                        End If
                    End If
                End If
        End Select
    Next
    
    If Index = -99999999 Then
        Dim IntChange As Integer
        Dim IntIndex As Integer, LngThisHdl As Long, intStyle As Integer
        mstrLastSelectCaption = ""
        mlngLastSelectIndex = 0
        
        '设置任务条模式
        For IntChange = 1 To gcllCollMap.Count
            IntIndex = gcllCollMap("K_" & IntChange)(0)
            LngThisHdl = gcllCollMap("K_" & IntChange)(1)
            intStyle = gcllCollMap("K_" & IntChange)(2)
            
            gcllCollMap.Remove "K_" & IntChange
            
            gcllCollMap.Add Array(IntIndex, LngThisHdl, 0), "K_" & IntChange
            Call zlControl.PicShowFlat(Pic任务(IntIndex), 2, , taCenterAlign)
        Next
    End If
End Sub

Public Sub AdjustPost()
    Dim IntReAdjust As Integer, DblTotalWidth As Double
    Dim DblPicResizeWidth As Double, DblPicResizeLeft As Double
    Dim DblLblResizeLeft As Double, DblLblResizeWidth As Double
    Dim DblPicResizeTop As Double, DblPicResizeHeight As Double
    On Error Resume Next
    
    DblTotalWidth = 2000
    DblTotalWidth = DblTotalWidth * (Pic任务.Count - 1)
    
    If DblTotalWidth > Sbar.Left - Pic分隔.Left - Pic分隔.Width - 1000 Then
        '--任务宽度和大于可容纳的空间--
        DblTotalWidth = ((Sbar.Left - Pic分隔.Left - Pic分隔.Width - 100) / IIf(Pic任务.Count - 1 = 0, 1, Pic任务.Count - 1)) - 50
        If DblTotalWidth > 2000 Then DblTotalWidth = 2000
    Else
        DblTotalWidth = 2000
    End If
    
    DblPicResizeTop = Pic任务(0).Top / Screen.TwipsPerPixelX
    DblPicResizeHeight = Pic任务(0).Height / Screen.TwipsPerPixelX
    
    DblPicResizeLeft = (Pic分隔.Left + Pic分隔.Width + 50) / Screen.TwipsPerPixelX
    DblPicResizeWidth = (DblTotalWidth) / Screen.TwipsPerPixelX
    Lbl任务(gcllCollMap("K_1")(0)).Width = DblTotalWidth - Lbl任务(gcllCollMap("K_1")(0)).Left - 100
    Call MoveWindow(Pic任务(gcllCollMap("K_1")(0)).hwnd, DblPicResizeLeft, DblPicResizeTop, DblPicResizeWidth, DblPicResizeHeight, 0)
    Call zlControl.PicShowFlat(Pic任务(gcllCollMap("K_1")(0)), 2, , taCenterAlign)
    
    For IntReAdjust = 3 To Pic任务.Count
        DblPicResizeLeft = (Pic任务(gcllCollMap("K_" & IntReAdjust - 2)(0)).Left + Pic任务(gcllCollMap("K_" & IntReAdjust - 2)(0)).Width + 50) / Screen.TwipsPerPixelX
        DblPicResizeWidth = DblTotalWidth / Screen.TwipsPerPixelX
        Lbl任务(gcllCollMap("K_" & IntReAdjust - 1)(0)).Width = DblTotalWidth - Lbl任务(gcllCollMap("K_" & IntReAdjust - 1)(0)).Left - 100
        Call MoveWindow(Pic任务(gcllCollMap("K_" & IntReAdjust - 1)(0)).hwnd, DblPicResizeLeft, DblPicResizeTop, DblPicResizeWidth, DblPicResizeHeight, 0)
        Call zlControl.PicShowFlat(Pic任务(gcllCollMap("K_" & IntReAdjust - 1)(0)), 2, , taCenterAlign)
    Next
    Pic任务条.Refresh
    
    For IntReAdjust = 1 To gcllCollMap.Count
        gcllCollMap("K_" & IntReAdjust)(2) = 0
    Next
End Sub

Private Function SetToolTipText(ByVal ObjCon As Object, ByVal NewIndex As Long, ByVal intStyle As Integer, ByVal strNote As String)
    '编制人:朱玉宝
    '编制日期:2000-11-21
    '功能:增加控件的帮助说明

    Select Case intStyle
    Case 0
        ObjCon(NewIndex).ToolTipText = strNote
    Case -1
        ObjCon(NewIndex).ToolTipText = "获取中联软件系统的在线帮助！"
    Case -2
        ObjCon(NewIndex).ToolTipText = "退出中联软件系统。"
    End Select
End Function

Private Sub SetZorder_Timer()
    Dim LngHdl As Long
    Dim FrmTest As Form
    Dim StrCaption As String * 255
    Dim StrTran As String
    Dim LngCount As Long
    Dim IntChange As Integer
    '用于保存集合中的数据
    Dim IntIndex As Integer, LngThisHdl As Long, intStyle As Integer
    
    Pic任务条.ZOrder
    PicToolTipText.ZOrder

    '获取当前活动窗体,如果不是主窗体,则将菜单关闭
    LngHdl = GetActiveWindow()
    If LngHdl <> Me.hwnd Then
        '关闭菜单
        mblnPress = True
        mblnAdjustPost = False
        Call ShowMenu
        
        '查找该子窗体对应的任务条
        On Error Resume Next
        LngCount = 254
        Call GetWindowText(LngHdl, StrCaption, LngCount)
        StrTran = Trim(Replace(StrCaption, Chr(0), ""))
        If StrTran <> mstrLastSelectCaption Then
            mstrLastSelectCaption = StrTran

            '设置任务条模式
            For IntChange = 1 To gcllCollMap.Count
                IntIndex = gcllCollMap("K_" & IntChange)(0)
                LngThisHdl = gcllCollMap("K_" & IntChange)(1)
                intStyle = gcllCollMap("K_" & IntChange)(2)
                
                gcllCollMap.Remove "K_" & IntChange
                
                If LngThisHdl <> LngHdl Then
                    gcllCollMap.Add Array(IntIndex, LngThisHdl, 0), "K_" & IntChange
                    Call zlControl.PicShowFlat(Pic任务(IntIndex), 2, , taCenterAlign)
                Else
                    gcllCollMap.Add Array(IntIndex, LngThisHdl, 1), "K_" & IntChange
                    Call zlControl.PicShowFlat(Pic任务(IntIndex), -2, , taCenterAlign)
                End If
            Next

            '激活当前窗体
            If IsWindowVisible(LngHdl) <> 0 Then
'                If IsIconic(LngHdl) Then Call ShowWindow(LngHdl, 9)                       '还原指定窗体为原大小
'                Call SetActiveWindow(LngHdl)
                If Not IsIconic(LngHdl) Then Call SetActiveWindow(LngHdl)
            End If
        End If
    Else
        If mblnAdjustPost = False Then
            Call AdjustPost
            mblnAdjustPost = True
        End If
    End If
End Sub

Private Sub InitEvn()
    Dim StrPicPath As String
    Dim LngColor As Long
    
    '--初始装入图标,图片--
    Img标识.Picture = LoadResPicture(101, 0) '菜单标识
    StrPicPath = zlDatabase.GetPara("zlWinBackPic")
    
    If Trim(StrPicPath) <> "" Then
        '用户选择图片,测试是否正常
        On Error Resume Next
        Err = 0
        
        Img标识.Picture = LoadPicture(StrPicPath)
        mblnShow = (Err = 0)
        
        If mblnShow Then PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Else
        mblnShow = True
    End If
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '恢复原来设置的图片
    Img标识.Picture = LoadResPicture(101, 0) '菜单标识
    
    '取字体色
    LngColor = Val(zlDatabase.GetPara("zlWinFontColor"))
    If LngColor <> -1 Then
        LvwList.ForeColor = LngColor
    End If
End Sub

Private Sub TimePass_Timer()
    Call Form_KeyDown(vbKeyF12, 7)  '清除静态变量
End Sub

Private Sub AdjustMenu(ByVal intLevel As Integer)
    Dim lngMin As Long, lngMax As Long, arrIndex
    Dim lngHeight As Long, lngRollHeight As Long, blnVisible As Boolean '菜单项的高度,非固定菜单项的总高度,是否显示该菜单
    Dim lngEnd As Long '最后一个显示的菜单项
    '调整菜单，对于超出显示的菜单项，禁止显示。最后增加滚动条
    
    If marrRoll(intLevel) = "" Then Exit Sub
    If PicBackDesktop(intLevel).Top >= 0 Then Exit Sub
    
    arrIndex = Split(marrRoll(intLevel), ",")
    lngMax = UBound(arrIndex)
    lngHeight = (-1 * PicBackDesktop(intLevel).Top) \ Pic菜单(0).Height + 2
    lngHeight = lngHeight * Pic菜单(0).Height
    
    '获取非固定菜单项的总高度
    lngEnd = 0
    lngRollHeight = 0
    For lngMin = 0 To lngMax
        lngRollHeight = lngRollHeight + Pic菜单(arrIndex(lngMin)).Height
    Next
    lngRollHeight = lngRollHeight - lngHeight + PicRollUp(0).Height
    
    '调整菜单及标识栏高度
    If intLevel = 1 Then
        Pic标识.Height = Pic标识.Height - lngHeight + PicRollUp(0).Height
        Pic标识.Top = Pic任务条.Top - Pic标识.Height
        Pic标识.Visible = True
        
        With Img标识
            .Top = 50
        End With
        
        '设置软件名称
        With Lbl标识
            .AutoSize = True
            .Caption = mstrTitle ' zlProductTitle(GetUnitInfo("发行码"))
            .AutoSize = False
            .Height = .Width
            .Width = 200
            .Left = Img标识.Left + 80
            .Top = Pic标识.Height - .Height - 100
            .ForeColor = IIf(GetSetting("ZLSOFT", "注册信息", "Kind", "") = "试用", &HFF, &HFFFFFF)
        End With
    End If
    PicBackDesktop(intLevel).Height = PicBackDesktop(intLevel).Height - lngHeight + PicRollUp(0).Height
    PicBackDesktop(intLevel).Top = Pic任务条.Top - PicBackDesktop(intLevel).Height
    PicBackDesktop(intLevel).ZOrder 0
    
    Call AddRollMenu(intLevel)
    
    '调整非固定菜单项
    For lngMin = 0 To lngMax
        blnVisible = True
        If lngMin <> 0 Then
            blnVisible = (Pic菜单(arrIndex(lngMin - 1)).Top + Pic菜单(arrIndex(lngMin - 1)).Height < lngRollHeight - PicRollUp(0).Height)
        End If
        Pic菜单(arrIndex(lngMin)).Visible = blnVisible
        If Not blnVisible And lngEnd = 0 Then lngEnd = lngMin - 1
    Next
    PicRollDown(intLevel).Top = Pic菜单(arrIndex(lngEnd)).Top + Pic菜单(arrIndex(lngEnd)).Height
    
    If intLevel <> 1 Then Exit Sub
    
    '调整系统固定菜单项
    Dim blnHistory As Boolean
    FraSplit(1).Top = PicRollDown(intLevel).Top + PicRollDown(intLevel).Height
    Pic菜单(9003).Top = FraSplit(1).Top + FraSplit(1).Height
    blnHistory = Trim(zlDatabase.GetPara("最近使用模块")) <> ""
    If blnHistory Then Pic菜单(9004).Top = Pic菜单(9003).Top + Pic菜单(9003).Height
    Pic菜单(9000).Top = Pic菜单(IIf(blnHistory, 9004, 9003)).Top + Pic菜单(IIf(blnHistory, 9004, 9003)).Height
    FraSplit(2).Top = Pic菜单(9000).Top + Pic菜单(9000).Height
    Pic菜单(9001).Top = FraSplit(2).Top + FraSplit(2).Height
    Pic菜单(9002).Top = Pic菜单(9001).Top + Pic菜单(9001).Height
End Sub

Private Sub AddRollMenu(ByVal intLevel As Integer)
    '增加滚动条
    
    Load PicRollDown(intLevel)
    Load PicRollUp(intLevel)
    Load ImgRollDown(intLevel)
    Load ImgRollUp(intLevel)
    
    With PicRollUp(intLevel)
        Set .Container = PicBackDesktop(intLevel)
        .Left = 50
        .Top = 50
        .Width = PicBackDesktop(intLevel).Width - 80
        .Tag = intLevel
    End With
    With PicRollDown(intLevel)
        Set .Container = PicBackDesktop(intLevel)
        .Left = 50
        .Width = PicBackDesktop(intLevel).Width - 80
        .Visible = True
        .Tag = intLevel
    End With
    With ImgRollDown(intLevel)
        Set .Container = PicRollDown(intLevel)
        .Left = PicRollDown(intLevel).Width / 2 - .Width
        .Visible = True
        .Tag = intLevel
    End With
    With ImgRollUp(intLevel)
        Set .Container = PicRollUp(intLevel)
        .Left = PicRollUp(intLevel).Width / 2 - .Width
        .Visible = True
        .Tag = intLevel
    End With
End Sub

Private Sub RollUpMenu(ByVal intLevel As Integer, Optional ByVal intWay As Integer = 1)
    Dim lngMin As Long, lngMax As Long, lngCur As Long
    Dim lngStart As Long, lngEnd As Long, blnVisible As Boolean
    Dim arrIndex
    '滚动菜单
    'intWay-滚动方向:1-向下;2-向上
    'lngStart保存Visible为真的第一个菜单的索引
    'lngEnd保存Visible为真的最后一个菜单的索引
    
    If marrRoll(intLevel) = "" Then Exit Sub
    
    arrIndex = Split(marrRoll(intLevel), ",")
    lngMax = UBound(arrIndex)
    blnVisible = False
    
    '根据intWay，找到第一个菜单项及最后一个菜单项
    lngCur = 0
    For lngMin = 0 To lngMax
        If Pic菜单(arrIndex(lngMin)).Visible Then
            If Not blnVisible Then
                lngStart = lngMin
                blnVisible = True
            End If
        Else
            If blnVisible Then
                lngEnd = lngCur
                Exit For
            End If
        End If
        lngCur = lngMin
    Next
    If lngEnd = 0 Then lngEnd = lngMax
    
    '重排菜单
    If (lngStart = 0 And intWay <> 1) Or (lngEnd = lngMax And intWay = 1) Then Exit Sub
    lngStart = lngStart + IIf(intWay = 1, 1, -1)
    lngEnd = lngEnd + IIf(intWay = 1, 1, -1)
    PicRollUp(intLevel).Visible = Not (lngStart = 0)
    PicRollDown(intLevel).Visible = Not (lngEnd = lngMax)
    
    For lngMin = 0 To lngMax
        blnVisible = (lngMin >= lngStart And lngMin <= lngEnd)
        Pic菜单(arrIndex(lngMin)).Visible = blnVisible
    Next
    For lngMin = lngStart To lngEnd
        If lngMin = lngStart Then
            Pic菜单(arrIndex(lngMin)).Top = IIf(PicRollUp(intLevel).Visible, PicRollUp(intLevel).Top + PicRollUp(intLevel).Height, 50)
        Else
            Pic菜单(arrIndex(lngMin)).Top = Pic菜单(arrIndex(lngMin - 1)).Top + Pic菜单(arrIndex(lngMin - 1)).Height
        End If
    Next
    
    Call zlControl.PicShowFlat(PicRollUp(intLevel), 0, , taCenterAlign)
    Call zlControl.PicShowFlat(PicRollDown(intLevel), 0, , taCenterAlign)
    If intWay <> 1 Then
        Call zlControl.PicShowFlat(PicRollUp(intLevel), -1, , taCenterAlign)
    Else
        Call zlControl.PicShowFlat(PicRollDown(intLevel), -1, , taCenterAlign)
    End If
End Sub

Private Sub LoadHistory()
    Dim str系统 As String, str序号 As String
    Dim arr系统 As Variant, arr序号 As Variant
    Dim int系统_Cur As Integer, int序号_Cur As Integer
    Dim int系统_Max As Integer, int序号_Max As Integer
    Dim strValue As String
    
    '将历史记录装入菜单
    strValue = zlDatabase.GetPara("最近使用模块")
    If UBound(Split(strValue, "|")) < 1 Then Exit Sub
    str系统 = Trim(Split(strValue, "|")(0))
    str序号 = Trim(Split(strValue, "|")(1))
    If str系统 = "" Or str序号 = "" Then Exit Sub
    
    arr系统 = Split(str系统, ",")
    arr序号 = Split(str序号, ",")
    int系统_Max = UBound(arr系统)
    int序号_Max = UBound(arr序号)
    If int系统_Max > 8 Then int系统_Max = 8 '最多八个历史记录
    
    For int系统_Cur = 0 To int系统_Max
        int序号_Cur = int系统_Cur
        If int序号_Cur > int序号_Max Then Exit For
        
        With grsMenus
            .Filter = "系统=" & IIf(arr系统(int系统_Cur) = "", 0, arr系统(int系统_Cur)) & " And 模块=" & arr序号(int序号_Cur)
            If .RecordCount <> 0 Then
                '设置缺省值
                Call SetMenuState(9200 + int系统_Cur, !编号, !图标, !标题, int系统_Cur + 1, 1, IIf(!模块 = 0, False, True), IIf(IsNull(!说明), "", !说明))
            End If
            .Filter = 0
        End With
    Next
End Sub

Private Sub LoadUsual()
    Dim str系统 As String, str序号 As String, str图标 As String, str标题 As String
    Dim arr系统 As Variant, arr序号 As Variant, arr图标 As Variant, arr标题 As Variant
    Dim int系统_Cur As Integer, int序号_Cur As Integer, int图标_Cur As Integer, int标题_Cur As Integer
    Dim int系统_Max As Integer, int序号_Max As Integer, int图标_Max As Integer, int标题_Max As Integer
    Dim strValue As String
    
    '增加常用功能
    strValue = zlDatabase.GetPara("常用功能模块")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    str系统 = Trim(Split(strValue, "|")(0))
    str序号 = Trim(Split(strValue, "|")(1))
    str图标 = Trim(Split(strValue, "|")(2))
    str标题 = Trim(Split(strValue, "|")(3))
    If str系统 = "" Or str序号 = "" Then Exit Sub
    
    arr系统 = Split(str系统, ",")
    arr序号 = Split(str序号, ",")
    arr图标 = Split(str图标, ",")
    arr标题 = Split(str标题, ",")
    int系统_Max = UBound(arr系统)
    int序号_Max = UBound(arr序号)
    int图标_Max = UBound(arr图标)
    int标题_Max = UBound(arr标题)
    
    For int系统_Cur = 0 To int系统_Max
        int序号_Cur = int系统_Cur
        int图标_Cur = int系统_Cur
        int标题_Cur = int系统_Cur
        If int序号_Cur > int序号_Max Then Exit For
        
        With grsMenus
            .Filter = "系统=" & arr系统(int系统_Cur) & " And 模块=" & arr序号(int序号_Cur)
            If .RecordCount <> 0 Then
                '设置缺省值
                str系统 = !系统
                str序号 = !模块
                If int标题_Cur <= int标题_Max Then
                    str标题 = arr标题(int标题_Cur)
                Else
                    str标题 = !标题
                End If
                If int图标_Cur <= int图标_Max Then
                    str图标 = arr图标(int图标_Cur)
                Else
                    str图标 = !图标
                End If
                Call AddUsualModul(str系统 & "♂" & str序号 & "♂" & str标题 & "♂" & str图标)
                Call Pic分隔_MouseMove(1, 0, 0, 0)
                If Pic分隔.Left + Pic常用功能(0).Width >= Sbar.Left - 3000 Then Exit Sub
            End If
            .Filter = 0
        End With
    Next
End Sub

Private Sub AddUsualModul(ByVal strModul As String)
    Dim lngAdd As Long
    Dim lng系统 As Long, lng模块 As Long, lng图标 As Long, str标题 As String
    '增加指定的常用功能控件
    
    '先分解参数
    lng系统 = Split(strModul, "♂")(0)
    lng模块 = Split(strModul, "♂")(1)
    str标题 = Split(strModul, "♂")(2)
    lng图标 = Split(strModul, "♂")(3)
    
    '加载控件
    lngAdd = Pic常用功能.Count
    Load Pic常用功能(lngAdd)
    Load Img常用功能(lngAdd)
    With Pic常用功能(lngAdd)
        Set .Container = Pic任务条
        .Left = Pic常用功能(lngAdd - 1).Left + Pic常用功能(lngAdd - 1).Width
        .Tag = strModul
        .Visible = True
    End With
    With Img常用功能(lngAdd)
        Set .Container = Pic常用功能(lngAdd)
        .Picture = GetPicDisp(lng图标)
        .Visible = True
    End With
    Call SetToolTipText(Pic常用功能, lngAdd, 0, str标题)
    Call SetToolTipText(Img常用功能, lngAdd, 0, str标题)
End Sub

Private Sub ShutUsual()
    '删除所有常用功能
    Dim ObjShut As Object, LngUnloadObjs As Long
    
    On Error Resume Next
    LngUnloadObjs = 0
    For Each ObjShut In Me.Controls '删除控件
        Err = 0
        Select Case TypeName(ObjShut)
            Case "Image", "PictureBox"
                If (ObjShut.Name Like "*常用功能") Then
                    If Err = 0 Then
                        Unload ObjShut
                    End If
                End If
        End Select
    Next

    For Each ObjShut In Me.Controls '因容器上次无法删除，再次执行
        Err = 0
        Select Case TypeName(ObjShut)
            Case "PictureBox"
                If (ObjShut.Name Like "*常用功能") Then
                    If Err = 0 Then
                        ObjShut.Visible = False
                        Unload ObjShut
                    End If
                End If
        End Select
    Next
End Sub

Private Sub ShowToolTipText(ByVal ObjCon As Object, Optional ByVal blnVisible As Boolean = True)
    Static ObjCon_Last As Object
    '1秒延迟后显示；4秒延迟后消失
    'blnVisible:强制关闭
    On Error Resume Next
    
    If Not blnVisible Then
        TimeToolTipText.Enabled = False
        PicToolTipText.Visible = False
        Exit Sub
    End If
    
    '如果对象不同，则重新计时（1秒）
    If Not ObjCon_Last Is Nothing Then
        If ObjCon_Last.Name = ObjCon.Name And ObjCon_Last.Index = ObjCon.Index Then
            '有可能上次选择的是菜单，而菜单可能已经关闭，所以判断是否发生错误
            If Err = 0 Then
                TimeToolTipText.Interval = 1000
                TimeToolTipText.Enabled = True
                Exit Sub
            End If
        End If
    End If
    Set ObjCon_Last = ObjCon
    
    PicToolTipText.Visible = False
    With TimeToolTipText
        .Enabled = False
        .Interval = 1000
        .Enabled = True
        .Tag = ObjCon.Container.Left + ObjCon.Left & "♂" & ObjCon.Container.Top + ObjCon.Top & "♂" & ObjCon.ToolTipText
        Exit Sub
    End With
End Sub

Private Sub TimeToolTipText_Timer()
    Dim BlnShowToolTipText As Boolean
    'BlnShowToolTipText:假-消失;真-显示
    
    With TimeToolTipText
        BlnShowToolTipText = (.Interval = 1000)
        .Enabled = False
        If BlnShowToolTipText Then
            .Interval = 4000
            .Enabled = True
        End If
    End With
    If Trim(TimeToolTipText.Tag) = "" Then Exit Sub
    LblToolTipText.Caption = Split(TimeToolTipText.Tag, "♂")(2)
    If Trim(LblToolTipText.Caption) = "" Then Exit Sub
    With PicToolTipText
        .Visible = BlnShowToolTipText
        .Left = Split(TimeToolTipText.Tag, "♂")(0) + 250
        .Top = Split(TimeToolTipText.Tag, "♂")(1) + 500
        .Width = LblToolTipText.Width + 80
        .ZOrder
        '如果超出边界，则反取
        If .Left < 0 Then
            .Left = 0
        End If
        If .Left > Me.Width Then
            .Left = Me.Width - .Width
        End If
        If .Top < 0 Then
            .Top = 0
        End If
        If .Top > Me.Height Then
            .Top = Split(TimeToolTipText.Tag, "♂")(1) - .Height
        End If
    End With
End Sub

Private Sub CheckTools()
    Dim blnSplit As Boolean         '是否显示分隔条
    '消息收发与EXCEL报表的权限控制：
    '1、如果授权码中含有此功能
    '2、如果该用户拥有此权限
    '3、显示这两个功能
    '其他工具模块仅判断该用户是否拥有此权限
    
    '工具对应说明
    '打印、预览、输出EXCEL  ,10,'导航功能清单','基本'
    'mnuToolDictonary       ,11,'字典管理工具','基本'
    'mnuToolMessage         ,12,'消息收发工具','基本,发送消息'
    'mnuTooleSelect         ,13,'系统选项设置','基本'
    'mnuToolExcel           ,14,'EXCEL报表工具','基本,报表增删,报表计算,所有系统'
    'mnuToolUp              ,15,'本地参数上传' ,'基本'
    
    Dim intGrant As Integer
    
    '导航功能清单
    'Excel报表工具
    mnuToolExcel.Visible = False
    '消息收发工具
    mnuToolMessage.Visible = False
    mnuToolNotify.Visible = False
    '系统选项设置
    mnuToolStyle.Visible = False
    '字典管理工具
    mnuToolDictonary.Visible = False
    
    '当然,分隔条一定是要禁止的,只要存在其中一个功能（字典管理、消息收发、EXCEL报表或系统选项），就需要显示分隔条
    blnSplit = False
    
    intGrant = zlRegTool '(GetUnitInfo("注册码"))
    If ((intGrant And 4) = 4) Then
        If InStr(1, GetPrivFunc(0, 工具清单.消息收发工具), "基本") <> 0 Then
            mnuToolMessage.Visible = True
            mnuToolNotify.Visible = True
            blnSplit = True
        Else
            Call zlDatabase.SetPara("接收邮件消息", "0")
        End If
    End If
    If ((intGrant And 8) = 8) Then
        If InStr(1, GetPrivFunc(0, 工具清单.EXCEL报表工具), "基本") Then
            mnuToolExcel.Visible = True
            blnSplit = True
        End If
    End If
    
    
    If InStr(1, GetPrivFunc(0, 工具清单.系统选项设置), "基本") Then
        mnuToolStyle.Visible = True
        blnSplit = True
    End If
    If InStr(1, GetPrivFunc(0, 工具清单.字典管理工具), "基本") Then
        mnuToolDictonary.Visible = True
        blnSplit = True
    End If
    MnuBar3.Visible = blnSplit
End Sub

Public Sub RunModual(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strPara As String, Optional ByVal blnReport As Boolean)
    '------------------------------------------------------------------------------------------------------
    '功能:调用执行报表,此功能是为自动提醒调用而写,by 陈福容
    '参数:lngSys 系统编号;lngModual 模块号
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHand
    
    With grsMenus
        If blnReport Then
            .Filter = "系统=" & lngSys & " AND 模块=" & lngModual & " And 报表=1"
        Else
            .Filter = "系统=" & lngSys & " AND 模块=" & lngModual
        End If
        If .RecordCount = 0 Then .Filter = 0: Exit Sub
        If .Fields("模块").Value <> 0 Then
            Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value, strPara)
        End If
        .Filter = 0
    End With
    
ErrHand:
    
End Sub

Private Function LoadOutTools(ByVal blnMenu As Boolean) As Boolean
    '-----------------------------------------------------------------------------------
    '功能:加载外部工具
    '参数:blnMenu-点击开始菜单显示列表
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim i As Long
    Dim strReg As String, arrTemp As Variant, ArrTool As Variant
    Dim objButton As ButtonMenu
    Err = 0: On Error Resume Next
    
    '先清除外部工具菜单
    
    For i = 1 To mnuToolOutToolExecute.UBound
        Unload mnuToolOutToolExecute(i)
    Next
    mnuToolOutToolList.Visible = False
    
    '加载工具菜单
    strReg = GetSetting("ZLSOFT", "公共全局\TOOLS", "TOOLFILES", "")
    Set mcllTemp = New Collection
    
    
    If strReg = "" Then Exit Function
    
    ArrTool = Split(strReg, "|")
    If blnMenu = True Then
        Call SetMenuState(9301, 9301, -4, "外接工具设置", i, 1, True, "外接工具设置")
    End If
    
    For i = 0 To UBound(ArrTool)
        arrTemp = Split(ArrTool(i) & ",", ",")
        If arrTemp(0) <> "" And arrTemp(1) <> "" And i <= 199 Then
            If i = 0 Then
                With mnuToolOutToolExecute(0)
                    .Caption = arrTemp(0) & "(&1)"
                    .Tag = arrTemp(1)
                    .Visible = True
                    mnuToolOutToolList.Visible = True
                End With
            Else
                Load mnuToolOutToolExecute(i)
                With mnuToolOutToolExecute(i)
                    .Caption = arrTemp(0) & IIf(i + 1 > 9, "", "(&" & i + 1 & ")")
                    .Tag = arrTemp(1)
                    .Visible = True
                End With
            End If
            If blnMenu = True Then
                Call SetMenuState(9300 + i + 2, -1 * 9300 + i + 2, 0, arrTemp(0), i, 1, True, arrTemp(1))
            End If
            mcllTemp.Add arrTemp(1), "K" & 9300 + i + 2
        End If
    Next
    LoadOutTools = True
End Function


Private Sub ExeCuteToolFile(ByVal strFile As String)
    '-----------------------------------------------------------------------------------
    '功能:执行工具文件
    '参数:strFile-文件名
    '编制:刘兴宏
    '日期:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Err = 0: On Error GoTo ErrHand:
    If objFile.FileExists(strFile) = False Then
        MsgBox "工具文件:" & strFile & vbCrLf & "不存在,可能已被删除,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    Shell strFile, vbNormalFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Function GetCommand() As String
    '功能:用于业务部件获取命令行参数,by 陈东
    '参数:无
    GetCommand = gstrCommand
End Function

Private Sub DoCommand()
    '功能：外部调用导航台时，根据传入参数启动业务部件。,by 陈东
    '参数：无
    Dim i As Integer, lngModual As Long
    Dim varCmd As Variant
    On Error GoTo errH
    varCmd = Split(gstrCommand, " ")
    For i = LBound(varCmd) To UBound(varCmd)
        If UCase(varCmd(i)) Like "PROGRAM=*" Then
            lngModual = Val(Split(varCmd(i), "=")(1))
            grsMenus.Filter = "模块=" & lngModual
            If Not grsMenus.EOF Then
                Call RunModual(grsMenus!系统, lngModual, "")
                mblnHide = True
            End If
            grsMenus.Filter = 0
        End If
    Next
    Exit Sub
errH:
    
End Sub

Public Sub UnloadForm()
    '功能：外部调用导航台启动业务部件后，业务部件在退出时，要调用此函数关闭导航台。by 陈东
    '参数：无
    Unload Me
End Sub

Private Sub tmrUpdateConnect_Timer()
    '预升级处理
    If DateAdd("n", -30, Now) >= mCurTime Then '30分钟检查一次
        tmrUpdateConnect.Enabled = False
        Call gobjRelogin.UpdateClient
        mCurTime = Now
        tmrUpdateConnect.Enabled = True
    End If
End Sub

Public Function CloseChildWindows(ByVal frmMain As Object) As Boolean
     '功能:关闭所有子窗口
    Dim FrmThis     As Form, ClsClose As Object, IntCount As Integer, LngErr As Long
    Dim objInsure   As Object
    Dim blnOK       As Boolean
    
    On Error Resume Next
    blnOK = True
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                blnOK = False
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogOutBefore
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            blnOK = False
            MsgBox "zlPlugIn 外挂部件执行 LogOutBefore 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
    On Error Resume Next
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption Then Unload FrmThis
    Next
    '关闭所有部件的窗体
    If Err.Number <> 0 Then Err.Clear
    LngErr = UBound(gstrObj)
    If Err.Number = 0 Then
        For IntCount = 0 To LngErr
            Set ClsClose = gobjCls(IntCount)
            blnOK = blnOK And ClsClose.CloseWindows
            Set gobjCls(IntCount) = Nothing
        Next
    End If
    '关闭应用工具包部件的窗体
    blnOK = blnOK And mclsAppTool.CloseWindows
    '关闭公共部件的窗体
    blnOK = blnOK And CloseWindows
    Set objInsure = GetObject("", "zl9Insure.clsInsure")
    Call objInsure.Releaseme
    If Err.Number <> 0 Then Err.Clear
    CloseChildWindows = blnOK
End Function

Public Function GetPicDisp(Optional ByVal intIcon As Long = 0, Optional ByVal Bln模块 As Boolean = True) As IPictureDisp
    '编制人:朱玉宝
    '编制日期:2000-12-12
    '得到图片对象

    On Error Resume Next
    If intIcon = 0 Then intIcon = IIf(Bln模块, -5, -4)
    Select Case intIcon
    Case -1
        Set GetPicDisp = LoadResPicture("HELP", 1)
    Case -2
        Set GetPicDisp = LoadResPicture("RELOGIN", 1)
    Case -3
        Set GetPicDisp = LoadResPicture("EXIT", 1)
    Case -4
        Set GetPicDisp = LoadResPicture("DIRECTORY", 1)
    Case -5
        Set GetPicDisp = LoadResPicture("MODUL", 1)
    Case Else
        Set GetPicDisp = mclsAppTool.GetIcon(intIcon)
    End Select
End Function

Private Sub InitWinsock()
'功能:获取参数,初始化服务器
    Dim lngPort As Long
            
    On Error Resume Next
    
    lngPort = Val(zlDatabase.GetPara("允许远程控制"))
    mblnRemote = Not lngPort = -1
    winSock.Tag = "1"
    With winSock
        If mblnRemote Then
            .LocalPort = IIf(Val(lngPort) = 0, "1001", Val(lngPort))
            .Listen
        Else
            If .State <> sckClosed Then .Close
        End If
    End With
    winSock.Tag = ""
End Sub

Private Sub winSock_Close()
    If winSock.Tag = "" Then
        If winSock.State <> sckClosed And mblnRemote Then winSock.Close: winSock.Listen  '重新监听
    End If
End Sub

Private Sub winSock_ConnectionRequest(ByVal requestID As Long)
    If winSock.State <> sckClosed Then winSock.Close
    winSock.Accept requestID
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strMsg  As String
    
    winSock.GetData strData
    
    On Error GoTo errH
    If strData = "请求远程" Then
                RunCommand "REG ADD HKLM\SYSTEM\CurrentControlSet\Control\Terminal"" ""Server /v fDenyTSConnections /t REG_DWORD /d 0 /f"
                winSock.SendData "YES"
    End If
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    winSock.Close: winSock.Listen
    If winSock.Tag = "" Then
        Select Case Number
            Case 10053
                MsgBox "由于长时间没有操作，连接自动中断。", vbInformation, gstrSysName
            Case Else
                MsgBox Number & Description, vbInformation, gstrSysName
         End Select
    Else
        winSock.Tag = ""
    End If
End Sub



