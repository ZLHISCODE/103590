VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#1.0#0"; "zlIDKind.ocx"
Begin VB.Form frmStation 
   Caption         =   "影像医技工作站"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   Icon            =   "frmStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   60
      ScaleHeight     =   6015
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   720
      Width           =   4495
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   495
         TabIndex        =   4
         Top             =   60
         Width           =   1485
      End
      Begin VB.TextBox Txt基本信息 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         Height          =   2100
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3285
         Width           =   4205
      End
      Begin XtremeCommandBars.CommandBars cbrdock 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7875
      Top             =   165
   End
   Begin zlIDKind.IDKind IDKind 
      Bindings        =   "frmStation.frx":1CFA
      Height          =   360
      Left            =   6975
      TabIndex        =   1
      Top             =   165
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   635
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStation.frx":1D0E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   2850
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":25A2
            Key             =   "紧急"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":2B3C
            Key             =   "住院"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":3416
            Key             =   "阳性"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":3570
            Key             =   "影像"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":3CEA
            Key             =   "已缴"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":4084
            Key             =   "绿色通道"
            Object.Tag             =   "6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1980
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":41DE
            Key             =   "复选留空"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":4778
            Key             =   "复选选中"
            Object.Tag             =   "90001"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmStation.frx":4D12
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mcol
    Col紧急 = 0: Col来源: Col阳性: Col质量: Col姓名: Col检查号: Col检查过程: Col性别: Col年龄: Col内容: Col部位: Col执行间: Col检查时间: Col开嘱时间: Col开嘱医生
    
    Col身高 = 15: Col体重: Col婴儿: Col登记人: Col报到人: Col完成人: Col打印胶片: Col报告操作: Col绿色通道: Col报告打印: Col报告人: Col复核人: Col检查技师: Col采图时间
    
    Col影像类别 = 29: Col病人ID: Col主页ID: Col挂号单: Col医嘱ID: Col发送号: Col检查UID: Col检查状态: Col转出 '从29列开始不显示
End Enum

Private Enum FilterID
    ID_门诊 = 4001: ID_住院 = 4002: ID_体检 = 4003: ID_外诊 = 4004
    ID_费用 = 4005: ID_已缴 = 4006: ID_未缴 = 4007: ID_登记 = 4008
    ID_报到 = 4009: ID_报告 = 4010: ID_审核 = 4011: ID_完成 = 4012
    ID_过滤方式 = 4013: ID_过滤值 = 4014: ID_开始过滤 = 4015: ID_本次住院 = 4016
End Enum

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0姓名或就诊卡 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
End Enum

Private mlngCur科室ID As Long                               '当前科室ID
Private mstrCur科室 As String                               '当前科室 编码-名称
Private mstrCanUse科室 As String                            '当前可用科室  ID_编码-名称
Private mstrCurFindtype As String                           '定位方式
Private mblnFinishCommit As Boolean                         '无报告完成里,是否无需再次确认
Private mblnCompleteCommit As Boolean                       '审核后无需再次确认
Private mblnInitOK As Boolean                               '初始化完成
Private mblnIgnoreResult As Boolean                         '忽略阴阳性 '=true 忽略
Private mblnShowImgAtReport As Boolean                      '打开报告时打开观片站
Private mblnReportWithImage As Boolean                      '有图像才能写报告，无图像不可写报告
Private mblnReportWithResult As Boolean                     '无影像诊断为阴性
Private mblnLocalizerBackward As Boolean                    '定位片后置
Private mblnPacsReport As Boolean                           '是否使用PACS报告编辑器，Fasle时使用电子病历编辑器
Private mblnPrintCommit As Boolean                          '打印后直接完成

Private mstrRoom As String                                  '只处理执行间内的病人
Private mstrPrivs As String, mlngModul As Long
Private mblnPatTrack As Boolean                             '是否对进病人进行跟踪
Private mbln直接检查 As Boolean                             '登记后直接进入检查
Private mblnNoShowCancel As Boolean                         '不显示取消的检查
Private mBeforeDays As Integer                              '默认查询的天数
Private mblnMoved As Boolean                                '当前时间段内是否被转移过
Private mblnOpenReport As Boolean                           '开始检查自动打开报告
Private mblnTechReptSame As Boolean                         '只能填写自己检查的报告
Private mintResultInput As Integer                          '提示输入阴阳性和影像质量

Private mblnUse3D As Boolean                                '是否启用三维重建功能
Private mstr3DExeDir As String                              '三维重建程序路径
Private mstr3DPara As String                                '三维重建参数
Private mstr3DFunctions As String                           '三维重建功能

'过滤条件变量
Private mdatFBegin As Date
Private mdatFEnd As Date
Private mDatType As Integer                                 '时间查询方式 1=按检查时间、2=按发送时间
Private mstrFNO As String
Private mlngF科室ID As Long
Private mstrF标识号 As String
Private mstrF就诊卡 As String
Private mstrF姓名 As String
Private mdblFChkNO As Double
Private mstr标本部位 As String
Private mstr诊断医生 As String
Private mstr审核医生 As String
Private mstr疾病诊断 As String
Private mbln结果阳性 As Boolean
Private mstr影像质量 As String
Private mstr检查技师 As String
Private mstr检查过程 As String
Private mstr影像类别 As String
Private mstr检查所见 As String
Private mstr诊断意见 As String
Private mstr建议 As String
Private mlngRefreshInterval As Long                         '病人列表自动刷新间隔
Private Sub InitVslist()
    With vsList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 38

        .ColWidth(Col紧急) = 200: .ColWidth(Col来源) = 200: .ColWidth(Col阳性) = 200: .ColWidth(Col质量) = 200: .ColWidth(Col姓名) = 400
        .ColWidth(Col检查号) = 600: .ColWidth(Col检查过程) = 600: .ColWidth(Col性别) = 400: .ColWidth(Col年龄) = 400: .ColWidth(Col内容) = 800
        .ColWidth(Col部位) = 800: .ColWidth(Col执行间) = 600: .ColWidth(Col检查时间) = 1000: .ColWidth(Col开嘱时间) = 1000: .ColWidth(Col开嘱医生) = 600
        .ColWidth(Col身高) = 400: .ColWidth(Col体重) = 400: .ColWidth(Col婴儿) = 400: .ColWidth(Col登记人) = 600: .ColWidth(Col报到人) = 600
        .ColWidth(Col完成人) = 600: .ColWidth(Col打印胶片) = 800: .ColWidth(Col报告操作) = 800: .ColWidth(Col绿色通道) = 800: .ColWidth(Col报告打印) = 800
        .ColWidth(Col报告人) = 600: .ColWidth(Col复核人) = 600: .ColWidth(Col检查技师) = 800: .ColWidth(Col采图时间) = 1000
        
        .ColWidth(Col影像类别) = 0: .ColWidth(Col病人ID) = 0: .ColWidth(Col主页ID) = 0: .ColWidth(Col挂号单) = 0
        .ColWidth(Col医嘱ID) = 0: .ColWidth(Col发送号) = 0: .ColWidth(Col检查UID) = 0: .ColWidth(Col检查状态) = 0: .ColWidth(Col转出) = 0


        .TextMatrix(0, Col紧急) = 200: .TextMatrix(0, Col来源) = 200: .TextMatrix(0, Col阳性) = 200: .TextMatrix(0, Col质量) = 200: .TextMatrix(0, Col姓名) = 400
        .TextMatrix(0, Col检查号) = 600: .TextMatrix(0, Col检查过程) = 600: .TextMatrix(0, Col性别) = 400: .TextMatrix(0, Col年龄) = 400: .TextMatrix(0, Col内容) = 800
        .TextMatrix(0, Col部位) = 800: .TextMatrix(0, Col执行间) = 600: .TextMatrix(0, Col检查时间) = 1000: .TextMatrix(0, Col开嘱时间) = 1000: .TextMatrix(0, Col开嘱医生) = 600
        .TextMatrix(0, Col身高) = 400: .TextMatrix(0, Col体重) = 400: .TextMatrix(0, Col婴儿) = 400: .TextMatrix(0, Col登记人) = 600: .TextMatrix(0, Col报到人) = 600
        .TextMatrix(0, Col完成人) = 600: .TextMatrix(0, Col打印胶片) = 800: .TextMatrix(0, Col报告操作) = 800: .TextMatrix(0, Col绿色通道) = 800: .TextMatrix(0, Col报告打印) = 800
        .TextMatrix(0, Col报告人) = 600: .TextMatrix(0, Col复核人) = 600: .TextMatrix(0, Col检查技师) = 800: .TextMatrix(0, Col采图时间) = 1000
        
        .TextMatrix(0, Col影像类别) = 0: .TextMatrix(0, Col病人ID) = 0: .TextMatrix(0, Col主页ID) = 0: .TextMatrix(0, Col挂号单) = 0
        .TextMatrix(0, Col医嘱ID) = 0: .TextMatrix(0, Col发送号) = 0: .TextMatrix(0, Col检查UID) = 0: .TextMatrix(0, Col检查状态) = 0: .TextMatrix(0, Col转出) = 0
        
        Dim i As Integer
        For i = 0 To .Cols
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        
        .Editable = flexEDNone
    End With
End Sub
Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim i As Integer
    Dim i3DFunc As Integer
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        '.SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    
'菜单定义
'Begin------------------------文件菜单--------------------------------------默认可见
    Me.cbrMain.ActiveMenuBar.Title = "菜单"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)"): cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)"): cbrControl.IconId = 102
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)"): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "批量打印(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "清单打印(&L)"): cbrControl.BeginGroup = True: cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&O)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "影像设备设置(&D)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_SendImg, "发送图像(&T)"): cbrControl.IconId = 3061
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"):: cbrControl.IconId = 191: cbrControl.BeginGroup = True
    End With


'Begin----------------------检查菜单--------------------------------------默认可见
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "检查(&S)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_RequestPrint, "打印申请单据(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "检查登记(&I)"): cbrControl.IconId = 211: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CopyCheck, "复制登记(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "取消登记(&R)"): cbrControl.IconId = 742
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "召回取消(&G)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "修改信息(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "检查报到(&L)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "取消报到(&D)"): cbrControl.IconId = 743
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "影像观片(&S)"): cbrControl.IconId = 8111:  cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "观片对比(&E)"): cbrControl.IconId = 8112
        
        '如果启用三维重建功能，则创建对应菜单
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "三维重建"): cbrControl.ID = conMenu_Img_3D
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Delete, "影像删除(&K)"): cbrControl.IconId = 8113
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Query, "Q/R获取图象(&Q)"): cbrControl.IconId = 8111
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "关联影像(&C)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 505: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "取消关联(&B)"): cbrControl.IconId = 506
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "检查结果(&X)"):  cbrControl.BeginGroup = True: cbrControl.ID = conMenu_Manage_Result
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "阳性(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "阴性(&Y)"): cbrPopControl.IconId = 3507
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "影像质量(&Y)"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "甲级(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "乙级(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_GChannel, "绿色通道(&G)"): cbrControl.ID = conMenu_Manage_GChannel
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelOk, "标记(&J)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelCancel, "取消(&Y)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "无报告完成(&F)"): cbrControl.IconId = 216: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "无报告回退(&U)"):  cbrControl.IconId = 3012
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "检查完成(&E)"): cbrControl.IconId = 225
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消完成(&U)"): cbrControl.IconId = 219
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ChangeDevice, "更换设备"): cbrControl.IconId = 3203
    End With
    
    
'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '二级菜单
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_FindType, "查找方式(&G)"): cbrControl.ID = conMenu_View_FindType
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Filter * 10#, "检查科室"): cbrControl.ID = conMenu_View_Filter * 10#
        Set cbrControl = .Add(xtpControlButton, conMenu_View_PatInfor, "病人信息(&P)"): cbrControl.IconId = 812
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "快速过滤(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&F)")
    End With


'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题", -1, False)
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "WEB上的中联(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModul, mstrPrivs)
    
'----------------------快键绑定------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print '打印------------------Ctrl+P
        .Add 0, VK_F12, conMenu_File_Parameter      '参数设置--------------F12
        
        .Add 0, VK_F2, conMenu_Manage_Regist       '登记-----------------F2
        .Add 0, VK_F7, conMenu_Manage_CopyCheck    '复制登记-------------F7
        .Add 0, VK_F4, conMenu_Manage_Receive       '报到-----------------F4
        .Add 0, VK_F9, conMenu_Manage_ClearUp       '驳回报告------------F9
        .Add 0, VK_F6, conMenu_Manage_Complete         '审核报告----------F6
        
        
        .Add 0, VK_F1, conMenu_Help_Help              '帮助-------------F1
        .Add 0, VK_F5, conMenu_View_Refresh           '刷新-------------F5
        .Add FCONTROL, Asc("F"), conMenu_View_FindType    '查找方式---------Ctrl+F
        .Add 0, VK_F3, conMenu_View_Filter            '过滤-------------F3
    End With
    
'---------------------设置右上角当前科室----------------------------------
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_Filter * 10#, "检查科室")
            cbrControl.ID = conMenu_View_Filter * 10#: cbrControl.Flags = xtpFlagRightAlign: cbrControl.Category = "Main"
    
'---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.IconId = 102: cbrControl.ToolTipText = "报告预览"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.IconId = 103: cbrControl.ToolTipText = "报告打印"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "登记"): cbrControl.BeginGroup = True: cbrControl.IconId = 211
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "报到"): cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "取消"): cbrControl.IconId = 743: cbrControl.ToolTipText = "取消报到"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "观片"): cbrControl.ToolTipText = "影像观片"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "对比"): cbrControl.IconId = 8112: cbrControl.ToolTipText = "观片对比"
        '如果启用三维重建功能，则创建对应菜单
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "三维"): cbrControl.ID = conMenu_Img_3D: cbrControl.ToolTipText = "三维重建"
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "结果"):  cbrControl.BeginGroup = True: cbrControl.ID = conMenu_Manage_Result: cbrControl.IconId = 3506: cbrControl.ToolTipText = "检查结果阴阳性"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "阳性(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "阴性(&Y)"): cbrPopControl.IconId = 3507
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "质量"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061: cbrControl.ToolTipText = "影像质量"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "甲级(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "乙级(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "完成"): cbrControl.IconId = 225: cbrControl.ToolTipText = "检查最终完成"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
    End With
End Sub
Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    On Error GoTo errH
    
    str来源 = "1,2,3"
    If InStr(mstrPrivs, "所有科室") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    Else
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And instr([1],','||B.服务对象||',')>0  And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    End If
   

    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, "," & str来源 & ",")
    
    If rsTmp.EOF Then
        MsgBoxD Me, "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
        Exit Function
    Else
        str科室IDs = GetUser科室IDs
        Do Until rsTmp.EOF
            mstrCanUse科室 = mstrCanUse科室 & "|" & rsTmp!ID & "_" & rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!ID = UserInfo.部门ID Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '提取默认科室
            If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur科室ID = 0 Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '没有默认科室取第一个所属检查科室
            rsTmp.MoveNext
        Loop
        mstrCanUse科室 = Mid(mstrCanUse科室, 2)
        If InStr(mstrPrivs, "所有科室") > 0 And mlngCur科室ID = 0 Then
            mlngCur科室ID = Split(Split(mstrCanUse科室, "|")(0), "_")(0)
            mstrCur科室 = Split(Split(mstrCanUse科室, "|")(0), "_")(1)
        End If
        
        If mlngCur科室ID = 0 And InStr(mstrPrivs, "所有科室") <= 0 Then '没有所有科室操作权限,而且操作者科室不属于检查类科室
            MsgBoxD Me, "没有发现你所属科室,不能使用医技工作站。", vbInformation, gstrSysName
            Exit Function
        End If
        InitDepts = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 240, 250, DockLeftOf, Nothing)
    Pane1.Title = "检查列表"
    Pane1.Handle = picList.Hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set Pane2 = dkpMain.CreatePane(2, 700, 250, DockRightOf, Nothing)
    Pane2.Title = "子窗体"
'    Pane2.Handle = PicWindow.Hwnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
End Sub
Private Sub InitLocalPars()
Dim TitleFont As New StdFont                                '病人列表表头字体
Dim TextFont As New StdFont                                 '病人列表内容字体
'初始化临时本地参数，以个人设置，注册表参数为主,窗体加载，本地设置等调用
    On Error GoTo err
    mblncmd门诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", 1))
    mblncmd住院 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", 1))
    mblncmd外诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", 1))
    mblncmd体检 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", 1))
    mblncmd已缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", 0))
    mblncmd未缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", 0))
    mblncmd登记 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", 1))
    mblncmd报到 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", 1))
    mblncmd报告 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", 1))
    mblncmd审核 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", 1))
    mblncmd完成 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", 1))
    mstrCurFindtype = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", "标识号")
    
    mbln本次 = (Val(zlDatabase.GetPara("只显示本次住院项目", glngSys, mlngModul, "1")) = 1)
    mbln直接检查 = (Val(zlDatabase.GetPara("登记直接检查", glngSys, mlngModul, 0)) = 1)
    mblnOpenReport = (Val(zlDatabase.GetPara("开始检查自动打开报告", glngSys, mlngModul, 0)) = 1)
    mblnShowImgAtReport = (Val(zlDatabase.GetPara("报告时观片", glngSys, mlngModul, 0)) = 1)
    mblnNoShowCancel = (Val(zlDatabase.GetPara("不显示被取消的登记", glngSys, mlngModul, 0)) = 1)
    mblnPatTrack = (Val(zlDatabase.GetPara("病人跟踪", glngSys, mlngModul, 0)) = 1)
    mstrRoom = zlDatabase.GetPara("执行间范围", glngSys, mlngModul, "")
    If mstrRoom <> "" Then mstrRoom = "'," & Replace(mstrRoom, "|", ",") & ",'"
    
    '读取和设置病人列表的字体
    TitleFont.Name = zlDatabase.GetPara("病人列表表头字体", glngSys, mlngModul, "宋体")
    TitleFont.Size = Val(zlDatabase.GetPara("病人列表表头字号", glngSys, mlngModul, 9))
    TitleFont.Bold = zlDatabase.GetPara("病人列表表头粗体", glngSys, mlngModul, 0) = 1
    TitleFont.Italic = zlDatabase.GetPara("病人列表表头斜体", glngSys, mlngModul, 0) = 1
    
    TextFont.Name = zlDatabase.GetPara("病人列表内容字体", glngSys, mlngModul, "宋体")
    TextFont.Size = Val(zlDatabase.GetPara("病人列表内容字号", glngSys, mlngModul, 9))
    TextFont.Bold = zlDatabase.GetPara("病人列表内容粗体", glngSys, mlngModul, 0) = 1
    TextFont.Italic = zlDatabase.GetPara("病人列表内容斜体", glngSys, mlngModul, 0) = 1
    
    Set rptList.PaintManager.CaptionFont = TitleFont
    Set rptList.PaintManager.TextFont = TextFont
    
    '读取三维重建参数
    mblnUse3D = Val(zlDatabase.GetPara("启用三维重建", glngSys, mlngModul, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D程序路径", glngSys, mlngModul, "")
    mstr3DPara = zlDatabase.GetPara("3D参数", glngSys, mlngModul, "")
    mstr3DFunctions = zlDatabase.GetPara("3D功能", glngSys, mlngModul, "")

    '过滤条件初始
    '-----------------------------------------------------
    mDatType = 1
    mstrFNO = ""
    mlngF科室ID = 0
    mstrF标识号 = 0
    mstrF就诊卡 = ""
    mstrF姓名 = ""
    mdblFChkNO = 0
    mstr标本部位 = ""
    mstr诊断医生 = ""
    mstr审核医生 = ""
    mstr疾病诊断 = ""
    mbln结果阳性 = False
    mstr影像质量 = ""
    mstr检查所见 = ""
    mstr诊断意见 = ""
    mstr建议 = ""
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub InitMvar()
'功能:初始化模块级变量,仅窗体加载时调用一次
    
    mblnIgnoreResult = GetDeptPara(mlngCur科室ID, "忽略结果阴阳性", "0") = "0"     '忽略结果阴阳性
    mintResultInput = GetDeptPara(mlngCur科室ID, "提示阴阳性", "1") = "1"           '提示阴阳性
    mblnFinishCommit = GetDeptPara(mlngCur科室ID, "无报告完成后直接完成", "0") = "0"      '无报告完成后直接完成
    mblnReportWithImage = GetDeptPara(mlngCur科室ID, "有图像才能写报告", "0") = "0"    '有图像才能写报告
    mblnReportWithResult = GetDeptPara(mlngCur科室ID, "无影像诊断为阴性", "0") = "0"   '无影像诊断为阴性
    mblnLocalizerBackward = GetDeptPara(mlngCur科室ID, "定位片后置", "0") = "0"  '定位片后置
    mblnCompleteCommit = GetDeptPara(mlngCur科室ID, "审核后直接完成", "0") = "0"      '审核后直接完成
    mBeforeDays = GetDeptPara(mlngCur科室ID, "默认过滤天数", "2")                  '默认过滤天数
    mblnTechReptSame = GetDeptPara(mlngCur科室ID, "只能填写自己检查的报告", "0") = "0"    '只能填写自己检查的报告
    mblnPacsReport = GetDeptPara(mlngCur科室ID, "报告编辑器", "0") = "0"        '报告编辑器
    mblnPrintCommit = GetDeptPara(mlngCur科室ID, "打印后直接完成", "0") = "0"         '打印后直接完成
    mlngRefreshInterval = GetDeptPara(mlngCur科室ID, "自动刷新间隔", "0")         '自动刷新间隔
    If mlngRefreshInterval > 65 Then
        mlngRefreshInterval = 30
    End If
    If mlngRefreshInterval <> 0 Then
        TimerRefresh.Interval = mlngRefreshInterval * 1000
        TimerRefresh.Enabled = True
    Else
        TimerRefresh.Enabled = False
    End If
    
    mdatFEnd = CDate(0)
    mdatFBegin = CDate(Format(zlDatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
    mblnMoved = MovedByDate(IIf(mdatFBegin = CDate(0), CDate(zlDatabase.Currentdate) - mBeforeDays, mdatFBegin))
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrdock.VisualTheme = xtpThemeOfficeXP
    With Me.cbrdock.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    cbrdock.AddImageList img16 '以VB.ImageList的Tag与ID进行关联
    cbrdock.EnableCustomization False
    cbrdock.ActiveMenuBar.Visible = False
    
    Set objBar = cbrdock.Add("来源", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_门诊, "门诊")
            objControl.ToolTipText = "显示门诊病人"
        Set objControl = .Add(xtpControlButton, ID_住院, "住院")
            objControl.ToolTipText = "显示住院病人"
        Set objControl = .Add(xtpControlButton, ID_外诊, "外诊")
            objControl.ToolTipText = "显示外诊病人"
        Set objControl = .Add(xtpControlButton, ID_体检, "体检")
            objControl.ToolTipText = "显示体检病人"
        Set objControl = .Add(xtpControlButtonPopup, ID_费用, " 费  用")
            objControl.ToolTipText = "显示费用已缴/未缴病人"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_未缴, "未缴")
            cbrPopControl.ToolTipText = "显示费用未缴病人"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_已缴, "已缴")
            cbrPopControl.ToolTipText = "显示费用已缴病人"
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("状态", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_登记, "登记")
            objControl.ToolTipText = "显示已登记病人"
        Set objControl = .Add(xtpControlButton, ID_报到, "报到")
            objControl.ToolTipText = "显示已报到病人"
        Set objControl = .Add(xtpControlButton, ID_报告, "报告")
            objControl.ToolTipText = "显示已报告病人"
        Set objControl = .Add(xtpControlButton, ID_审核, "审核")
            objControl.ToolTipText = "显示已审核病人"
        Set objControl = .Add(xtpControlButton, ID_完成, "完成")
            objControl.ToolTipText = "显示已完成病人"
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbsMain.Add("过滤", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    Set objPopbar = objBar.Controls.Add(xtpControlPopup, ID_过滤方式, "标识号(&D)")
        objPopbar.ID = ID_过滤方式
        objPopbar.Flags = xtpFlagRightAlign
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_过滤值, "标识号")
        objCusControl.Handle = txtFilter.Hwnd
        objCusControl.Flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_开始过滤, "过滤")
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_本次住院, "当次")
    objControl.ToolTipText = "只显示本次住院检查记录"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter
    
    With cbrdock.KeyBindings
        .Add FCONTROL, vbKey0, ID_门诊
        .Add FCONTROL, vbKey1, ID_住院
        .Add FCONTROL, vbKey2, ID_外诊
        .Add FCONTROL, vbKey3, ID_体检
        .Add FCONTROL, vbKey4, ID_费用
        .Add FCONTROL, vbKey5, ID_登记
        .Add FCONTROL, vbKey6, ID_报到
        .Add FCONTROL, vbKey7, ID_报告
        .Add FCONTROL, vbKey8, ID_审核
        .Add FCONTROL, vbKey9, ID_完成
    End With
End Sub
Private Sub Menu_File_Excel_click(ByVal blnNoRecord As Boolean)
Dim bytMode As Byte
    If blnNoRecord Then Exit Sub
    On Error GoTo ErrHandle
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "检查项目清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_File_BatPrint()
Dim cbrControl As CommandBarControl, strReturn As String, l As Integer
Dim objReportPrint As New zlRichEPR.cDockReport
    Set cbrControl = Me.cbrMain(2).FindControl(, conMenu_File_Print)
    If Not cbrControl Is Nothing Then
        cbrControl.ID = conMenu_File_BatPrint
    Else
        Exit Sub
    End If

    '选病人
    strReturn = frmDocPrintPatiList.Showfrm(rptList, Me)
    '循环调用
    For l = 0 To UBound(Split(strReturn, "|"))
        objReportPrint.zlRefresh CLng(Split(strReturn, "|")(l)), mlngCur科室ID
        Call objReportPrint.zlExecuteCommandBars(cbrControl)
        Call AfterPrinted(CLng(Split(strReturn, "|")(l)))
    Next
    cbrControl.ID = conMenu_File_Print
    Unload objReportPrint.zlGetForm
End Sub
Private Sub Menu_RichEPR(ByVal cbrID As Long)
    Dim cbrControl As CommandBarControl, i As Integer
    
    '报告页面不可见时不执行任何操作
    If TabWindow.Selected.Tag <> "报告填写" Then
        For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
            If TabWindow(i).Tag = "报告填写" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
        Next
        If TabWindow.Selected.Tag <> "报告填写" Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    '刷新嵌入页面内容
    If mblnPacsReport = True Then
        Call mfrmPacsReport.zlRefresh(Nvl(rptList.FocusedRow.Record(mcol.医嘱ID).Value, 0), Nvl(rptList.FocusedRow.Record(mcol.发送号).Value, 0), mlngCur科室ID, mstrPrivs, mlngModul, Me, rptList.FocusedRow.Record(mcol.转出).Value = 1)
    Else
        Call mobjReport.zlRefresh(Nvl(rptList.FocusedRow.Record(mcol.医嘱ID).Value, 0), mlngCur科室ID, True)
    End If
    
    '判断按键可用性
    Set cbrControl = Me.cbrMain.FindControl(, IIf(mblnPacsReport, conMenu_PacsReport_Open, cbrID))
    If cbrControl Is Nothing Then Exit Sub
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    Call cbrMain_Execute(cbrControl)
End Sub
Private Sub Menu_File_Parmeter_click()
    With frmTechnicSetup
        .mlngModul = mlngModul
        .mlng科室ID = mlngCur科室ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        If .mblnOK Then
            InitLocalPars
            Call RefreshRptlist
        End If
    End With
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '功能：调用帮助主题
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.Hwnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo Hwnd
End Sub

Private Sub Menu_Manage_取消关联(ByVal intState As Integer)
'取消关联的最后结果是，每次取消关联后，图象全部按照序列被拆散成N条临时记录
Dim strFilter As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    '显示序列选择窗口
    With rptList.FocusedRow
        gstrSQL = "select 0 as 选择,B.序列UID as ID ,B.序列号,B.序列描述,SUM(1) AS 图像数 from 影像检查记录 A ," & _
                "影像检查序列 B, 影像检查图象 C Where a.检查UID = B.检查UID And B.序列UID = C.序列UID" & _
                " And a.医嘱ID = [1] and A.发送号= [2] group by B.序列UID,B.序列号,B.序列描述"
        Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption, CLng(.Record(mcol.医嘱ID).Value), CLng(.Record(mcol.发送号).Value))
        
        frmSelectMuli.ShowSelect rsTmp, "ID,3000,0,1;序列号,800,0,1;序列描述,2000,0,1;图像数,800,0,1", 0, 0, 14000, 10000, "取消关联"
        
        If frmSelectMuli.mblnOK = True Then
            strFilter = frmSelectMuli.strFilter
            rsTmp.Filter = strFilter
            '如果有选中序列，则处理每一个序列的取消
            While Not rsTmp.EOF
                subCancelSeriesRelate CLng(.Record(mcol.医嘱ID).Value), CLng(.Record(mcol.发送号).Value), rsTmp!ID
                rsTmp.MoveNext
            Wend
            
            '设置影像检查状态，如果当前医嘱已经没有图像，而且检查过程为3，则修改为2
            If intState = 3 Then
                gstrSQL = "Select 检查uid From 影像检查记录 Where  医嘱ID=[1] And 发送号=[2]"
                Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption, CLng(.Record(mcol.医嘱ID).Value), CLng(.Record(mcol.发送号).Value))
                If IsNull(rsTmp!检查UID) Then
                    gstrSQL = "Zl_影像检查_State(" & CLng(.Record(mcol.医嘱ID).Value) & "," & CLng(.Record(mcol.发送号).Value) & ",2)"
                    zlDatabase.ExecuteProcedure gstrSQL, "取消关联"
                End If
            End If
            
            mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
            Call RefreshRptlist '真正取消关联点确定才刷新
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_无报告完成()
Dim blnTran As Boolean, arrSQL() As Variant, l As Long
'只有进行中的报告可以操作该菜单,因为此时还没有签名
        On Error GoTo ErrHandle
        arrSQL = Array()
        With rptList.FocusedRow
            If .Record(mcol.报告ID).Value <> 0 Then
                If MsgBoxD(Me, "是否无报告直接完成,直接完成将删除已填写的报告!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            If mblnFinishCommit And InStr(mstrPrivs, "检查完成") > 0 Then '无报告完成后无需再次确认完成,但需要有检查完成的权限
                '此过程,传状态=6,并且报告ID不为空将删除电子病历记录
                If zlDatabase.GetPara(81, glngSys) = 1 And Not bln病人在院(Nvl(.Record(mcol.病人ID).Value), Nvl(.Record(mcol.主页ID).Value)) And bln存在未审划价单(Nvl(.Record(mcol.医嘱ID).Value)) Then '执行后自动审核划价单有效，并且病人已出院，且有未审核的划价单
                    MsgBoxD Me, "该病人已出院，且有未审核的划价单不能完成！", vbExclamation, gstrSysName
                Else
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_影像检查_STATE(" & .Record(mcol.医嘱ID).Value & "," & .Record(mcol.发送号).Value & ",6" & IIf(.Record(mcol.报告ID).Value <> 0, "," & .Record(mcol.报告ID).Value, "") & ")"
                End If
            Else
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_影像检查_STATE(" & .Record(mcol.医嘱ID).Value & "," & .Record(mcol.发送号).Value & ",5" & _
                            IIf(.Record(mcol.报告ID).Value <> 0, "," & .Record(mcol.报告ID).Value, "") & ")"
            End If
        End With
        
        gcnOracle.BeginTrans '--------------------------写入数据
        blnTran = True
        For l = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "写入中心病种数据")
        Next
        gcnOracle.CommitTrans
        
        If mblnPatTrack Then
            If mblnFinishCommit Then
                Call StateCheck(6)
            Else
                Call StateCheck(5)
            End If
        Else
            Call RefreshRptlist
        End If
        Exit Sub
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Edit_无报告回退()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If MsgBoxD(Me, "确认要回退该项检查吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    With rptList.FocusedRow
            
            '如果有图像，则回退到“已检查”，否则回退到“已报到”
            strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有图像", CLng(.Record(mcol.医嘱ID).Value))
            
            gstrSQL = "ZL_影像检查_STATE(" & .Record(mcol.医嘱ID).Value & "," & .Record(mcol.发送号).Value & "," & IIf(IsNull(rsTemp!检查UID) = True, 2, 3) & ")"
            ExecuteProc gstrSQL, Me.Caption
    End With
    If mblnPatTrack Then
        Call StateCheck(2)
    Else
        Call RefreshRptlist
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_检查最终完成(lng医嘱ID As Long, Optional blnRefresh As Boolean = True)
    Dim arrSQL() As Variant, l As Long, blnTran As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If InStr(mstrPrivs, "检查完成") <= 0 Then Exit Sub
    
    strSQL = "Select a.发送号,b.病人ID,b.主页ID From 病人医嘱发送 a,病人医嘱记录 b Where a.医嘱id = [1] And a.医嘱ID=b.Id"
    Set rsTemp = OpenSQLRecord(strSQL, "检查最终完成", lng医嘱ID)
    
    If rsTemp.EOF = True Then Exit Sub
    
    arrSQL = Array()
    If zlDatabase.GetPara(81, glngSys) = 1 And Not bln病人在院(Nvl(rsTemp!病人ID), Nvl(rsTemp!主页ID, 0)) And bln存在未审划价单(Nvl(lng医嘱ID)) Then '执行后自动审核划价单有效，并且病人已出院，且有未审核的划价单
        MsgBoxD Me, "该病人已出院，且有未审核的划价单，不能完成！", vbExclamation, gstrSysName
    Else
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_影像检查_STATE(" & lng医嘱ID & "," & rsTemp!发送号 & ",6)"
    End If

    gcnOracle.BeginTrans '--------------------------写入数据
    blnTran = True
    For l = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "写入中心病种数据")
    Next
    gcnOracle.CommitTrans

    If blnRefresh Then Call StateCheck(6)
    Exit Sub

ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_取消检查完成()
    On Error GoTo ErrHandle
    With rptList.FocusedRow
            If .Record(mcol.转出).Value = 1 Then MsgBox "该病人的本次住院数据已经转出到后备数据库，不允许操作。", vbInformation, gstrSysName: Exit Sub
            gstrSQL = "ZL_影像检查_STATE(" & .Record(mcol.医嘱ID).Value & "," & .Record(mcol.发送号).Value & ",5)"
            ExecuteProc gstrSQL, "取消检查完成"
    End With

    Call StateCheck(5)
    Exit Sub

ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_标记阴阳(ByVal lngID As Long)
    Dim iresult As Integer

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_Negative
            iresult = 1
        Case conMenu_Manage_Positive
            iresult = 0
    End Select
    With rptList.FocusedRow
        gstrSQL = "ZL_影像检查_结果(" & .Record(mcol.医嘱ID).Value & "," & iresult & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "结果阴阳性")
        .Record(mcol.阳性).Value = IIf(iresult = 1, "阳性", "")
        .Record(mcol.阳性).Icon = IIf(iresult = 1, Me.imgList.ListImages("阳性").Index - 1, -1)
        .Selected = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_绿色通道(ByVal lngID As Long)
    Dim intResult As Integer

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    With rptList.FocusedRow
        gstrSQL = "Zl_绿色通道_Update(" & .Record(mcol.医嘱ID).Value & ",'" & intResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "影像质量")
        .Record(mcol.绿色通道).Value = intResult
        .Record(mcol.姓名).Icon = IIf(intResult = 1, Me.imgList.ListImages("绿色通道").Index - 1, -1)
        .Selected = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_影像质量(ByVal lngID As Long)
    Dim strResult As String

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_First
            strResult = "甲"
        Case conMenu_Manage_Second
            strResult = "乙"
    End Select
    With rptList.FocusedRow
        gstrSQL = "Zl_影像质量_Update(" & .Record(mcol.医嘱ID).Value & ",'" & strResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "影像质量")
        .Record(mcol.质量).Value = strResult
        .Selected = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_修改(ByVal blnNoRecord As Boolean, ByVal intState As Integer)
    If blnNoRecord Then Exit Sub
    
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = rptList.FocusedRow.Record(mcol.发送号).Value
        .mlngAdviceID = rptList.FocusedRow.Record(mcol.医嘱ID).Value
        .mintEditMode = IIf(intState > 1, 3, 1) '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .InitMvar
        .RefreshPatiInfor False '刷新病人
        .mblnOK = False
        .Show 1, Me
        If .mblnOK Then RefreshRptlist '成功返回
    End With
End Sub
Private Sub Menu_Manage_复制登记()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .mblnOK = False
        .InitMvar
        .CopyCheck rptList.FocusedRow.Record(mcol.医嘱ID).Value, rptList.FocusedRow.Record(mcol.发送号).Value '刷新病人
        .Show 1, Me
        If .mblnOK Then '成功返回
            If mbln直接检查 Then
                Call StateCheck(2)
            Else
                Call RefreshRptlist
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_登记()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .mblnOK = False
        .InitMvar
        .Show 1, Me
        If .mblnOK Then '成功返回
            If mbln直接检查 Then
                Call StateCheck(2)
            Else
                Call RefreshRptlist
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_取消登记()
    On Error GoTo ErrHandle
    With rptList.FocusedRow
        If MsgBoxD(Me, "确认要取消当前申请吗？" & Chr(10) & Chr(13) & "申请取消后，其对应的医嘱将拒绝执行！", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "ZL_病人医嘱执行_拒绝执行(" & .Record(mcol.医嘱ID).Value & "," & .Record(mcol.发送号).Value & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消登记")
        Call RefreshRptlist
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_召回取消()
'功能：召回被取消的登记
    Dim lng医嘱ID As Long, lng发送号 As Long
    
    On Error GoTo errH
    
    With rptList.SelectedRows(0)
        If MsgBoxD(Me, "确实要召回被取消登记的项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        lng医嘱ID = .Record(mcol.医嘱ID).Value
        lng发送号 = .Record(mcol.发送号).Value
    End With
    
    gstrSQL = "ZL_病人医嘱执行_取消拒绝(" & lng医嘱ID & "," & lng发送号 & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call RefreshRptlist
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub Menu_Manage_报到()
Dim i As Long, cbrControl As CommandBarControl, blnFocusFind As Boolean

    blnFocusFind = (Me.ActiveControl.Name = "Txt标识号")
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = rptList.FocusedRow.Record(mcol.发送号).Value
        .mlngAdviceID = rptList.FocusedRow.Record(mcol.医嘱ID).Value
        .mintEditMode = 2 '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .InitMvar
        .RefreshPatiInfor True '刷新病人
        .mblnOK = False
        .Show 1, Me
        If .mblnOK Then  '成功返回
            Call StateCheck(2)
            If mblnOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '开始检查自动打开报告
        End If
        If blnFocusFind Then Txt标识号.SetFocus '自动定位到定位栏
    End With
End Sub
Private Sub Menu_Manage_取消报到(ByVal intState As Integer)
Dim rsTemp As ADODB.Recordset, lngcur医嘱ID As Long
    If intState <= 1 Then Call Menu_Manage_取消登记: Exit Sub '工具栏调用
    
    On Error GoTo ErrHandle
    With rptList.FocusedRow
        '------------------------------------有签名的需要先回退签名后再撤消
        lngcur医嘱ID = .Record(mcol.医嘱ID).Value
        gstrSQL = "Select Distinct B.完成时间 From 病人医嘱报告 A, 电子病历记录 B Where A.病历ID=B.Id And A.医嘱ID=[1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, "提取是否签名", CLng(.Record(mcol.医嘱ID).Value))
        If Not rsTemp.EOF Then
            If Nvl(rsTemp!完成时间, "") <> "" Then '签名保存
                MsgBoxD Me, "当前病人的检查报告已经签名,若需取消检查,请先回退签名!", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If

        If MsgBoxD(Me, "取消本次检查将删除相应的检查图像和检查报告，是否继续？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        If .Record(mcol.检查UID).Value <> "" And InStr(mstrPrivs, "清除图像") <= 0 Then
            MsgBoxD Me, "您没有清除检查图像权限,不能请除图像,所有不能取消此项检查!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        gstrSQL = "ZL_影像检查_CANCEL(" & .Record(mcol.医嘱ID).Value & "," & .Record(mcol.发送号).Value & "," & Nvl(.Record(mcol.报告ID).Value, 0) & ")"
        ExecuteProc gstrSQL, Me.Caption
        '删除影像文件和目录
        RemoveCheckImages .Record(mcol.医嘱ID).Value, .Record(mcol.发送号).Value

    End With
    
    Call StateCheck(1)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_观片()
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.医嘱ID).Value, rptList.FocusedRow.Record(mcol.发送号).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.转出).Value = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("影像处理")
End Sub
Private Sub Menu_Manage_对比观片()
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.医嘱ID).Value, rptList.FocusedRow.Record(mcol.发送号).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.转出).Value = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("影像对比")
End Sub
            
Private Sub Menu_Manage_图象删除()
Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.医嘱ID).Value, rptList.FocusedRow.Record(mcol.发送号).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.转出).Value = 1)
    End If
    
    gstrSQL = "select 检查UID from 影像检查记录 where 医嘱ID =[1] and  发送号 = [2]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取检查UID", CLng(rptList.FocusedRow.Record(mcol.医嘱ID).Value), CLng(rptList.FocusedRow.Record(mcol.发送号).Value))
    If rsTemp.EOF Then Exit Sub
    
    If MsgBoxD(Me, "是否确认要删除该检查的所有影像？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    '删除影像文件和目录
    RemoveCheckImages CLng(rptList.FocusedRow.Record(mcol.医嘱ID).Value), CLng(rptList.FocusedRow.Record(mcol.发送号).Value)
    gstrSQL = "ZL_影像检查_PhotoDelete(" & CLng(rptList.FocusedRow.Record(mcol.医嘱ID).Value) & "," & CLng(rptList.FocusedRow.Record(mcol.发送号).Value) & ")"
    ExecuteProc gstrSQL, Me.Caption
    
    '设置影像检查状态，如果检查过程为3，则修改为2
    If Val(Mid(rptList.FocusedRow.Record(mcol.检查状态).Value, 1, 1)) = 3 Then
        gstrSQL = "Zl_影像检查_State(" & CLng(rptList.FocusedRow.Record(mcol.医嘱ID).Value) & "," & CLng(rptList.FocusedRow.Record(mcol.发送号).Value) & ",2)"
        zlDatabase.ExecuteProcedure gstrSQL, "删除图像"
    End If
    
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshRptlist
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
        
Private Sub Menu_Manage_获取图像()
Dim strImageDeviceNumber As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.医嘱ID).Value, rptList.FocusedRow.Record(mcol.发送号).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.转出).Value = 1)
    End If
    
    strImageDeviceNumber = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSImageDeviceSetup", "默认影像设备", "")
    
    '没有默认设备时处理
    If strImageDeviceNumber = "" Then
        If MsgBoxD(Me, "没有设置默认影像检查设备！是否现在设置？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            frmPACSImageDeviceSetup.Show vbModal, Me
            strImageDeviceNumber = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSImageDeviceSetup", "默认影像设备", "")
            If strImageDeviceNumber = "" Then Exit Sub
        End If
    End If
    
    gstrSQL = "select 设备号,设备名, IP地址,端口号,本地AE,设备AE from 影像设备目录 where 设备号 = [1] "
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, Mid(strImageDeviceNumber, 2))
    
    '当默认设备被删除后重新设置
    If rsTemp.EOF = True Then
        MsgBoxD Me, "默认设备已被删除，请重新设置！", vbInformation, gstrSysName
        frmPACSImageDeviceSetup.Show vbModal, Me
        Exit Sub
    End If
        
    frmPACSGetDeviceImage.ShowMe Me, rsTemp("IP地址"), rsTemp("端口号"), rsTemp("设备名"), Nvl(rsTemp("本地AE")), Nvl(rsTemp("设备AE")), rptList.FocusedRow.Record(mcol.医嘱ID).Value
    Call RefreshRptlist
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_关联影像(ByVal intState As Integer)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    With rptList.FocusedRow
        Call funRelateSeries(CLng(.Record(mcol.医嘱ID).Value), CLng(.Record(mcol.发送号).Value))
    End With
    
    '设置影像检查状态，如果原来的状态是已报到，则修改成已检查，
    If intState < 3 Then
        '如果病人已经有图像，则修改成已检查
        strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有图像", CLng(rptList.FocusedRow.Record(mcol.医嘱ID).Value))
        
        If Not IsNull(rsTemp!检查UID) Then
            gstrSQL = "Zl_影像检查_State(" & CLng(rptList.FocusedRow.Record(mcol.医嘱ID).Value) & "," & CLng(rptList.FocusedRow.Record(mcol.发送号).Value) & ",3)"
            zlDatabase.ExecuteProcedure gstrSQL, "关联影像"
        End If
    End If
    
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshRptlist
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_View_Find_click()
    Txt标识号.SetFocus
End Sub
Private Sub Menu_View_Find_Type_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    mstrCurFindtype = Split(control.Caption, "(")(0)
    cbrMain.RecalcLayout
    If mstrCurFindtype = "ＩＣ卡" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Else
            Txt标识号.Text = mobjICCard.Read_Card(Me)
        End If
    End If
    Txt标识号.SetFocus
End Sub
Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
    If mlngCur科室ID <> control.DescriptionText Then
        mlngCur科室ID = control.DescriptionText
        mstrCur科室 = Split(control.Caption, "(")(0)
        Call cbrMain.RecalcLayout
        Call InitMvar
        Call InitSubForm
        Call RefreshRptlist
    End If
End Sub
Private Sub Menu_View_病人信息(ByVal blnNoRecord As Boolean, ByVal intState As Integer)
    If blnNoRecord Then Exit Sub
    Call frmDegreeCard.ShowInfo(Me, rptList.FocusedRow.Record(mcol.病人ID).Value)
End Sub
Private Sub Menu_View_Refresh_click()
   Call RefreshRptlist
End Sub

'
Private Sub Menu_Help_Web_Home_click()
    zlHomePage Hwnd
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub cboTimes_Click()
Dim lng医嘱ID As Long, str医嘱附件 As String, rsTemp As ADODB.Recordset, i As Integer, strSQLBak As String, rsAnnex As ADODB.Recordset
    If cboTimes.ListCount <= 1 Then Exit Sub
    
    If lbl个人信息.Caption = "" Then Exit Sub '此时还没加载信息完成，属listindex赋值触发
    On Error GoTo ErrHandle
    lng医嘱ID = cboTimes.ItemData(cboTimes.ListIndex)
    
    If lng医嘱ID = rptList.FocusedRow.Record(mcol.医嘱ID).Value Then '当次与当前选中医嘱ID相同时不由本函数控制
        Call rptList_SelectionChanged
        mblnIsHistory = False
        Exit Sub
    End If
    mblnIsHistory = True
    
    '提取病人历次医嘱及相关信息
    gstrSQL = "Select /*+ Rule */ a.病人来源,A.病人ID,a.Id 医嘱id,a.主页id ,a.病人科室id,a.挂号单,a.医嘱内容,c.检查UID, " & _
                "b.发送号,b.执行状态,b.执行过程,c.身高,c.体重,c.检查号,d.当前病区id 病区id ,d.姓名,d.性别,d.年龄,0 as 转出," & _
                "Decode(a.病人来源, 1, d.门诊号, 2, d.住院号, 4, d.门诊号, Null) As 标识号,d.当前床号,a.开嘱医生,f.名称 as 病人科室 " & _
                " From 病人医嘱记录 a, 病人医嘱发送 b, 影像检查记录 c, 病人信息 d, 影像检查项目 e,部门表 f" & _
                " Where a.Id = [1] And a.相关id Is Null " & _
                    " And a.Id = b.医嘱id And b.医嘱id = c.医嘱id(+) And b.发送号 = c.发送号(+)" & _
                    " And a.病人id = d.病人id And a.诊疗项目id = e.诊疗项目id  And f.id = a.病人科室ID "
    strSQLBak = gstrSQL
    strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
    strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
    strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
    strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
    gstrSQL = gstrSQL & " Union ALL " & strSQLBak
    
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取历次记录", lng医嘱ID)
    If rsTemp.EOF Then
        Select Case TabWindow(TabWindow.Selected.Index).Tag
            Case "影像图象"
                mfrmPACSImg.zlRefresh 0, 0, mstrPrivs, False
            Case "报告填写"
                If mblnPacsReport = True Then
                    mfrmPacsReport.zlRefresh 0, 0, 0, mstrPrivs, mlngModul, Me, False
                Else
                    mobjReport.zlRefresh 0, mlngCur科室ID, False
                End If
            Case "申请费用"
                mobjExpense.zlRefresh mlngCur科室ID, 0, 0, False
            Case "住院医嘱"
                mobjInAdvice.zlRefresh 0, 0, 0, 0, 0, False, 0, 0
            Case "门诊医嘱"
                mobjOutAdvice.zlRefresh 0, "", False, False, 0
            Case "住院病历"
                mobjInEPRs.zlRefresh 0, 0, mlngCur科室ID, False
            Case "门诊病历"
                mobjOutEPRs.zlRefresh 0, 0, mlngCur科室ID, False
        End Select
        Txt基本信息 = ""
        lbl个人信息.Caption = "姓  名:" & Space(12) & "性  别:" & Space(13) & "年  龄:" & Space(10) & "标识号:" & Space(12) & "床  号:" & Space(10)
        lbl检查信息.Caption = "检查号:" & Space(12) & "病人科室:" & Space(11) & "开嘱医生:" & Space(8) & "检查项目:"
        lblCash.Visible = False
        Exit Sub
    End If
    
    Txt基本信息 = ""
    If InStr(Nvl(rsTemp!医嘱内容), ":") > 0 Then
        For i = 0 To UBound(Split(Split(rsTemp!医嘱内容, ":")(1), "),"))
            If i = 0 Then
                Txt基本信息 = "检查部位:" & vbCrLf & Space(2) & "1:" & Split(Split(rsTemp!医嘱内容, ":")(1), "),")(i) & ")"
            Else
                Txt基本信息 = Txt基本信息 & vbCrLf & Space(2) & i + 1 & ":" & Split(Split(rsTemp!医嘱内容, ":")(1), "),")(i) & ")"
            End If
        Next
        If Trim(Txt基本信息) <> "" Then Txt基本信息 = Mid(Txt基本信息, 1, Len(Txt基本信息) - 1)
    Else
        Txt基本信息 = Txt基本信息 & "检查部位:" & Nvl(rsTemp!医嘱内容)
    End If
    
    '显示各种信息
    '提取病人历次病人医嘱附件
    gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
    If rsTemp!转出 = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱附件", "H病人医嘱附件")
    Set rsAnnex = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人医嘱附件", lng医嘱ID)
    Do Until rsAnnex.EOF
        str医嘱附件 = str医嘱附件 & rsAnnex!项目 & ":" & Nvl(rsAnnex!内容) & vbCrLf
        rsAnnex.MoveNext
    Loop
    Txt基本信息 = Txt基本信息 & vbCrLf & vbCrLf & str医嘱附件
    lbl个人信息.Caption = "姓  名:" & Rpad(Nvl(rsTemp!姓名), 12, " ") & "性  别:" & Rpad(Nvl(rsTemp!性别), 13, " ") & _
                                  "年  龄:" & Rpad(Nvl(rsTemp!年龄), 10, " ") & "标识号:" & Rpad(Nvl(rsTemp!标识号), 12, " ") & _
                                  "床  号:" & Rpad(Nvl(rsTemp!当前床号), 10, " ")
    lbl检查信息.Caption = "检查号:" & Rpad(Nvl(rsTemp!检查号), 12, " ") & "病人科室:" & Rpad(Nvl(rsTemp!病人科室), 11, " ") & _
                                  "开嘱医生:" & Rpad(Nvl(rsTemp!开嘱医生), 8, " ") & "检查项目:" & Split(rsTemp!医嘱内容, ":")(0)
                                  
    lblCash.Caption = "历": lblCash.Visible = True
    
    mlngHOrderID = lng医嘱ID
    mlngHSendNo = Nvl(rsTemp!发送号, 0)
    mstrHStudyUID = Nvl(rsTemp!检查UID)
    mblnHMoved = IIf(rsTemp!转出 = 1, True, False)
    
    If Nvl(rsTemp!病人来源, 3) <> 3 Then '根据病人来源控制病历及医嘱选项卡
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "门诊病历", "门诊医嘱"
                    TabWindow(i).Visible = False
                Case "住院病历", "住院医嘱"
                    TabWindow(i).Visible = True
                Case "影像图象"
                    TabWindow(i).Visible = True
                Case "报告填写" '已登记状态不能查看报告页
                    TabWindow(i).Visible = Nvl(rsTemp!执行过程, 0) > 1
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "门诊病历", "门诊医嘱"
                    TabWindow(i).Visible = True
                Case "住院病历", "住院医嘱"
                    TabWindow(i).Visible = False
                Case "影像图象"
                    TabWindow(i).Visible = True
                Case "报告填写"
                    TabWindow(i).Visible = Nvl(rsTemp!执行过程, 0) > 1
            End Select
        Next
    End If
    '刷新当前页信息
    Select Case TabWindow(TabWindow.Selected.Index).Tag
        Case "影像图象"
            mfrmPACSImg.zlRefresh lng医嘱ID, Nvl(rsTemp!发送号, 0), mstrPrivs, rsTemp!转出 = 1
        Case "报告填写"
            If mblnPacsReport = True Then
                mfrmPacsReport.zlRefresh lng医嘱ID, Nvl(rsTemp!发送号, 0), mlngCur科室ID, mstrPrivs, mlngModul, Me, rsTemp!转出 = 1
            Else
                mobjReport.zlRefresh lng医嘱ID, mlngCur科室ID, False
            End If
        Case "申请费用"
            mobjExpense.zlRefresh mlngCur科室ID, lng医嘱ID, Nvl(rsTemp!发送号, 0), rsTemp!转出 = 1
        Case "住院医嘱"
            mobjInAdvice.zlRefresh Nvl(rsTemp!病人ID, 0), Nvl(rsTemp!主页ID, 0), Nvl(rsTemp!病区ID, 0), Nvl(rsTemp!病人科室ID, 0), 0, rsTemp!转出 = 1, lng医嘱ID, Nvl(rsTemp!执行状态, 1)
        Case "门诊医嘱"
            mobjOutAdvice.zlRefresh Nvl(rsTemp!病人ID, 0), Nvl(rsTemp!挂号单, ""), False, rsTemp!转出 = 1, lng医嘱ID
        Case "住院病历"
            mobjInEPRs.zlRefresh Nvl(rsTemp!病人ID, 0), Nvl(rsTemp!主页ID, 0), mlngCur科室ID, False, rsTemp!转出 = 1
        Case "门诊病历"
            mobjOutEPRs.zlRefresh Nvl(rsTemp!病人ID, 0), Nvl(rsTemp!主页ID, 0), mlngCur科室ID, False, rsTemp!转出 = 1
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboTimes_DropDown()
    Call SendMessage(cboTimes.Hwnd, &H160, 500, 0)
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
Select Case control.ID
        Case ID_门诊
            mblncmd门诊 = Not control.Checked
        Case ID_住院
            mblncmd住院 = Not control.Checked
        Case ID_外诊
            mblncmd外诊 = Not control.Checked
        Case ID_体检
            mblncmd体检 = Not control.Checked
        Case ID_已缴
            mblncmd已缴 = Not control.Checked
            If mblncmd已缴 Then mblncmd未缴 = False
        Case ID_未缴
            mblncmd未缴 = Not control.Checked
            If mblncmd未缴 Then mblncmd已缴 = False
        Case ID_登记
            mblncmd登记 = Not control.Checked
        Case ID_报到
            mblncmd报到 = Not control.Checked
        Case ID_报告
            mblncmd报告 = Not control.Checked
        Case ID_审核
            mblncmd审核 = Not control.Checked
        Case ID_完成
            mblncmd完成 = Not control.Checked
    End Select
cbrdock.RecalcLayout
Call RefreshRptlist
End Sub

Private Sub cbrdock_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    If CommandBar.Parent Is Nothing Then Exit Sub
    If CommandBar.Parent.ID = ID_过滤方式 Then
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 0, "标识号(&1)"): objControl.Checked = True
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 1, "就诊卡(&2)")
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 2, "姓名(&3)")
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 3, "单据号(&4)")
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 4, "检查号(&5)")
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 5, "身份证(&6)")
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 6, "ＩＣ卡(&7)")
                Set objControl = .Add(xtpControlButton, ID_过滤方式 * 100# + 7, "病理号(&8)")
            End If
        End With
    End If

End Sub

Private Sub cbrdock_Resize()
Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    rptList.Top = lngTop
    rptList.Width = picList.Width
    rptList.Height = picList.Height - lngTop - Txt基本信息.Height - 100

    Txt基本信息.Top = rptList.Top + rptList.Height + 100
    Txt基本信息.Width = picList.Width - 200
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case ID_门诊
            control.Checked = mblncmd门诊
            control.IconId = IIf(mblncmd门诊, 90001, 90000)
        Case ID_住院
            control.Checked = mblncmd住院
            control.IconId = IIf(mblncmd住院, 90001, 90000)
        Case ID_外诊
            control.Checked = mblncmd外诊
            control.IconId = IIf(mblncmd外诊, 90001, 90000)
        Case ID_体检
            control.Checked = mblncmd体检
            control.IconId = IIf(mblncmd体检, 90001, 90000)
        Case ID_费用
            control.Checked = mblncmd已缴 Xor mblncmd未缴
            control.Caption = IIf(mblncmd已缴 Xor mblncmd未缴, IIf(mblncmd已缴, " 已缴费", " 未缴费"), " 费  用")
        Case ID_已缴
            control.Checked = mblncmd已缴
            control.IconId = IIf(mblncmd已缴, 90001, 90000)
        Case ID_未缴
            control.Checked = mblncmd未缴
            control.IconId = IIf(mblncmd未缴, 90001, 90000)
        Case ID_登记
            control.Checked = mblncmd登记
            control.IconId = IIf(mblncmd登记, 90001, 90000)
        Case ID_报到
            control.Checked = mblncmd报到
            control.IconId = IIf(mblncmd报到, 90001, 90000)
        Case ID_报告
            control.Checked = mblncmd报告
            control.IconId = IIf(mblncmd报告, 90001, 90000)
        Case ID_审核
            control.Checked = mblncmd审核
            control.IconId = IIf(mblncmd审核, 90001, 90000)
        Case ID_完成
            control.Checked = mblncmd完成
            control.IconId = IIf(mblncmd完成, 90001, 90000)
    End Select
End Sub
Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnNoRecord As Boolean  '是否有当前记录
    Dim intState As Integer
    If control.ID <> 0 Then
        If cbrMain.FindControl(, control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    If Not blnNoRecord Then
        intState = Val(Mid(rptList.FocusedRow.Record(mcol.检查状态).Value, 1, 1))
    End If
    
    cbrMain.RecalcLayout
    Select Case control.ID
    
'--------------------------文件------------------
        Case conMenu_File_PrintSet '打印设置
            Call zlPrintSet
           
        Case conMenu_File_Excel '清单打印
            Call Menu_File_Excel_click(blnNoRecord)
            
        Case conMenu_File_BatPrint '批量打印
            Call Menu_File_BatPrint
            
        Case conMenu_File_Parameter '参数设置
            Call Menu_File_Parmeter_click
            
        Case conMenu_Cap_DevSet '影像设备设置
            frmPACSImageDeviceSetup.Show vbModal, Me
            
        Case conMenu_File_SendImg '发送图像
            frmPacsSendImage.ShowMe Me
            
        Case conMenu_File_Exit '退出
            Unload Me
            
'---------------------------检查-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '打印诊疗单据
            Call FuncBillPrint(control)
            
        Case conMenu_Manage_Regist                          '登记
            Call Menu_Manage_登记
            
        Case conMenu_Manage_CopyCheck                       '复制登记
            Call Menu_Manage_复制登记
            
        Case conMenu_Manage_Receive                         '报到
            Call Menu_Manage_报到
            
        Case conMenu_Manage_Redo                            '取消登记
            Call Menu_Manage_取消登记
            
        Case conMenu_Manage_ReGet                           '召回取消
            Call Menu_Manage_召回取消
        
        Case conMenu_Manage_ThingModi                       '修改登记
            Call Menu_Manage_修改(blnNoRecord, intState)
            
        Case conMenu_Manage_Logout                          '取消报到
            Call Menu_Manage_取消报到(intState)
            
        Case conMenu_Img_Look                         '观片
            Call Menu_Manage_观片
        
        Case conMenu_Img_Contrast                     '对比观片
            Call Menu_Manage_对比观片
        
        Case conMenu_Img_3D_MMPR                    '三维重建，MMPR
            Call sub三维重建("MMPR")
        Case conMenu_Img_3D_MPR                     '三维重建，MPR
            Call sub三维重建("MPR")
        Case conMenu_Img_3D_PF                     '三维重建,灌注成像
            Call sub三维重建("PF")
        Case conMenu_Img_3D_SA                     '三维重建，表面重建
            Call sub三维重建("SA")
        Case conMenu_Img_3D_VA                     '三维重建，容积重建
            Call sub三维重建("VA")
        Case conMenu_Img_3D_VE                     '三维重建，虚拟内窥镜
            Call sub三维重建("VE")
            
        Case conMenu_Img_Delete                       '图象删除
            Call Menu_Manage_图象删除
        
        Case conMenu_Img_Query                        '从设备获取图象
            Call Menu_Manage_获取图像
        
        Case conMenu_Manage_Transfer                        '关联影像
            Call Menu_Manage_关联影像(intState)
            
        Case conMenu_Manage_Cancel                          '取消关联
            Call Menu_Manage_取消关联(intState)
        
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '结果阴阳性
            Call Menu_Manage_标记阴阳(control.ID)
            
        Case conMenu_Manage_First, conMenu_Manage_Second
            Call Menu_Manage_影像质量(control.ID)
            
        Case conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
            Call Menu_Manage_绿色通道(control.ID)
            
        Case conMenu_Manage_ClearUp                           '无报告回退
            Call Menu_Edit_无报告回退
                    
        Case conMenu_Manage_Finish                          '无报告直接完成
            Call Menu_Manage_无报告完成
            
        Case conMenu_Manage_Complete                        '检查完成
            If Not rptList.FocusedRow Is Nothing Then
                Call Menu_Manage_检查最终完成(rptList.FocusedRow.Record(mcol.医嘱ID).Value)
            End If
        
        Case conMenu_Manage_Undone                          '取消检查完成
            Call Menu_Manage_取消检查完成
            
        Case conMenu_Manage_ChangeDevice                    '更换检查设备
            Call Menu_Manage_更换检查设备
            
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_FindType * 10# To conMenu_View_FindType * 10# + 6
            Call Menu_View_Find_Type_click(control)
        Case conMenu_View_PatInfor
            Call Menu_View_病人信息(blnNoRecord, intState)
        Case conMenu_View_Filter '过滤
            Call Menu_View_Filter_click
        Case conMenu_View_Refresh '刷新
            Call Menu_View_Refresh_click
            
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            'Case Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse科室, "|"))
            Call Menu_Dept_Select(control)
        Case Else
            If Between(control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And control.Parameter <> "" Then
                 '执行发布到当前模块的报表
                If Not blnNoRecord Then
                    Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, _
                        "NO=" & rptList.FocusedRow.Record(mcol.NO).Value, "性质=" & rptList.FocusedRow.Record(mcol.记录性质).Value, _
                        "医嘱id=" & rptList.FocusedRow.Record(mcol.医嘱ID).Value, 1)
                Else
                    Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, "", 1)
                End If
            Else
                If Not blnNoRecord Then
                    Select Case TabWindow.Selected.Tag
                        Case "报告填写"
                            '没报告不能打印和预览
                            If Nvl(rptList.FocusedRow.Record(mcol.报告ID).Value, 0) = 0 And (control.ID = conMenu_File_Preview Or control.ID = conMenu_File_Print) Then
                                MsgBoxD Me, "当前病人没有检查报告，不能操作，请检查！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            '报告被某人打开后再被允许它人编辑或修订
                            If control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Modify Or control.ID = conMenu_PacsReport_Open Or control.ID = conMenu_Edit_Delete Then
                                If CheckConcurrentReport(rptList.FocusedRow.Record(mcol.医嘱ID).Value) = False Then Exit Sub
                            End If
                            
                            '控制 只能书写自己检查的报告,'不允许书写、修订、删除
                            If mblnTechReptSame = True _
                                And (control.ID = conMenu_Edit_Modify Or control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Delete) _
                                And Nvl(rptList.FocusedRow.Record(mcol.检查技师).Value) <> "" _
                                And Nvl(rptList.FocusedRow.Record(mcol.检查技师).Value) <> mstrUserNameHIS Then
                                MsgBoxD Me, "你不是这个患者的检查技师，无法操作这份报", vbInformation, gstrSysName
                            Else
                                If mblnPacsReport = True Then
                                    If control.ID = conMenu_PacsReport_Open Then   '打开报告窗体
                                        Call Menu_Manage_PACS报告
                                    Else
                                        mfrmPacsReport.zlExecuteCommandBars control
                                    End If
                                Else
                                    mobjReport.zlExecuteCommandBars control
                                End If
                            End If
                        Case "申请费用"
                            mobjExpense.zlExecuteCommandBars control
                        Case "住院医嘱"
                            mobjInAdvice.zlExecuteCommandBars control
                        Case "门诊医嘱"
                            mobjOutAdvice.zlExecuteCommandBars control
                        Case "住院病历"
                            mobjInEPRs.zlExecuteCommandBars control
                        Case "门诊病历"
                            mobjOutEPRs.zlExecuteCommandBars control
                    End Select
                End If
            End If
    End Select
End Sub

Private Sub Menu_View_Filter_click()
    On Error GoTo ErrHandle
    
    With frmPACSFilter
        .mlngModul = mlngModul
        .mBeforeDays = mBeforeDays
        .mDept = mlngCur科室ID '当前科室
        .Show 1, Me
        If Not .mblnOK Then Exit Sub '没有返回条件
        
        mdatFBegin = Format(.dtpBegin.Value, "yyyy-MM-dd HH:mm:00")
        If Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            mdatFEnd = CDate(0) '表示取当前时间
        Else
            mdatFEnd = Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm:59")
        End If
        
        mblnMoved = MovedByDate(IIf(mdatFBegin = CDate(0), CDate(zlDatabase.Currentdate) - mBeforeDays, mdatFBegin))
        
        '是否本次住院
        mbln本次 = (.chk本次住院.Value = 1)
        
        '时间查找方式：1＝按检查时间、2＝按开嘱时间
        If .optFindType(0).Value = True Then
            mDatType = 1
        Else
            mDatType = 2
        End If
        
        '单据号
        If .txtNO.Text <> "" Then
            mstrFNO = .txtNO.Text
        Else
        mstrFNO = ""
        End If
        
        '检查标本部位
        If .cboPart.ListIndex <> 0 Then
            mstr标本部位 = .cboPart.Text
        Else
            mstr标本部位 = ""
        End If
        
        '病人科室
        If .cboDept.ListIndex <> 0 Then
            mlngF科室ID = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            mlngF科室ID = 0
        End If
        
        '病人标识
        If .Txt标识号.Text <> "" Then
            mstrF标识号 = Trim(.Txt标识号.Text)
        Else
            mstrF标识号 = 0
        End If
        '就诊卡
        If .txt就诊卡.Text <> "" Then
            mstrF就诊卡 = .txt就诊卡.Text
        Else
            mstrF就诊卡 = ""
        End If
        '姓名
        If .txt姓名.Text <> "" Then
            mstrF姓名 = .txt姓名.Text
        Else
            mstrF姓名 = ""
        End If
        '检查号
        If .txtChkNO.Text <> "" Then
            mdblFChkNO = Val(.txtChkNO.Text)
        Else
            mdblFChkNO = 0
        End If

        '诊断医生
        If .cbodiagdoc.ListIndex <> 0 Then
            mstr诊断医生 = NeedName(.cbodiagdoc.Text)
        Else
            mstr诊断医生 = ""
        End If
        '审核医生
        If .cboAuditing.ListIndex <> 0 Then
            mstr审核医生 = NeedName(.cboAuditing.Text)
        Else
            mstr审核医生 = ""
        End If
        
        '检查过程
        If .cboCheckStep.ListIndex <> 0 Then
            mstr检查过程 = .cboCheckStep.Text
        Else
            mstr检查过程 = ""
        End If
        
        '影像类别
        If .cboModality.ListIndex <> 0 Then
            mstr影像类别 = Split(.cboModality.Text, "--")(1)
        Else
            mstr影像类别 = ""
        End If
        
        '影像诊断
        If Trim(.Txt影像诊断) <> "" Then
            mstr疾病诊断 = Trim(.Txt影像诊断)
        Else
            mstr疾病诊断 = ""
        End If
        
        If .chk结果阳性.Value = 1 Then
            mbln结果阳性 = True
        Else
            mbln结果阳性 = False
        End If
        
        If .cbo质量.ListIndex = 0 Then
            mstr影像质量 = ""
        Else
            mstr影像质量 = NeedName(.cbo质量.Text)
        End If
        
        If .cbo检查技师.ListIndex = 0 Then
            mstr检查技师 = ""
        Else
            mstr检查技师 = NeedName(.cbo检查技师.Text)
        End If
        
        'PACS报告检索
        If Trim(.txtPacsRpt(0)) <> "" Then
            mstr检查所见 = Trim(.txtPacsRpt(0))
        Else
            mstr检查所见 = ""
        End If
        
        If Trim(.txtPacsRpt(1)) <> "" Then
            mstr诊断意见 = Trim(.txtPacsRpt(1))
        Else
            mstr诊断意见 = ""
        End If
        
        If Trim(.txtPacsRpt(2)) <> "" Then
            mstr建议 = Trim(.txtPacsRpt(2))
        Else
            mstr建议 = ""
        End If
        
        '调用刷新
        Call RefreshRptlist
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_View_FindType
            With CommandBar.Controls
                If .Count = 0 Then '动态子菜单,扩1位
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10#, "标识号(&1)"): objControl.Category = "Main": objControl.Checked = True
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 1, "就诊卡(&2)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 2, "姓名(&3)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 3, "单据号(&4)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 4, "检查号(&5)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 5, "身份证(&6)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 6, "ＩＣ卡(&7)"): objControl.Category = "Main"
                End If
            End With
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    For i = 0 To UBound(Split(mstrCanUse科室, "|")) 'mstrCanUse科室=id_编码-名称|id_编码-名称
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i, Split(Split(mstrCanUse科室, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse科室, "|")(i), "_")(0)
                        If mlngCur科室ID = objControl.DescriptionText Then objControl.Checked = True
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.Tag
                Case "住院医嘱"
                    mobjInAdvice.zlPopupCommandBars CommandBar
                Case "门诊医嘱" '门诊
                    mobjOutAdvice.zlPopupCommandBars CommandBar
                Case "申请费用"
                    mobjExpense.zlPopupCommandBars CommandBar
            End Select
    End Select
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim blnNoRecord As Boolean, intState As Integer, intStep As Integer
    If Not mblnInitOK Then Exit Sub
    
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    control.Style = cbrMain(2).Controls(3).Style
    
    If Not blnNoRecord Then
        intState = Nvl(rptList.FocusedRow.Record(mcol.检查状态).Value, 0)
        intStep = Nvl(rptList.FocusedRow.Record(mcol.执行状态).Value, 0)
    End If
    
    Select Case control.ID
        Case conMenu_View_FindType
            control.Caption = "按" & mstrCurFindtype & "查找(&G)"
        Case conMenu_View_FindType * 10# To conMenu_View_FindType * 10# + 6
            control.Checked = (InStr(control.Caption, mstrCurFindtype) > 0)
        Case conMenu_View_Filter * 10#
            control.Caption = "当前科室:" & mstrCur科室
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse科室, "|"))
            control.Checked = (control.DescriptionText = mlngCur科室ID)
        Case conMenu_View_ToolBar_Button '工具栏
            If cbrMain.Count >= 2 Then
                control.Checked = Me.cbrMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbrMain.Count >= 2 Then
                control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            control.Checked = Me.cbrMain.Options.LargeIcons
        Case conMenu_View_StatusBar '状态栏
            control.Checked = Me.stbThis.Visible
        Case conMenu_View_PatInfor '病人信息
            control.Enabled = Not blnNoRecord
        Case conMenu_View_Filter   '过滤
        
        Case conMenu_View_Refresh  '刷新
        
        Case conMenu_Manage_RequestPrint
            control.Enabled = control.CommandBar.Controls.Count > 0 And Not blnNoRecord
                
        Case conMenu_Manage_Regist   '检查登记(&I)
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            End If
        Case conMenu_Manage_CopyCheck '再次登记
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Redo   '取消登记(&R)
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And intStep <> 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ReGet   '召回取消
            If Not blnNoRecord Then
                control.Enabled = intStep = 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ThingModi   '修改信息(&M)
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 3 And intStep <> 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Receive   '检查报到(&L)
            If InStr(mstrPrivs, "检查报到") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And intStep <> 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Logout   '取消报到(&D)
            If blnNoRecord Then
                control.Enabled = False
            ElseIf control.Parent.Type = xtpControlPopup Then
                If InStr(mstrPrivs, "取消报到") <= 0 Then
                    control.Visible = False
                Else
                    control.Visible = True
                    control.ToolTipText = "取消报到"
                    control.Caption = "取消报到(&D)"
                    control.Enabled = (intState = 2 Or intState = 3)
                End If
            Else ' 工具栏中的用取消检查代替取消登记,同一按键完成取消登记和取消检查功能
                control.Visible = IIf(intState <= 1, InStr(mstrPrivs, "检查登记") > 0, InStr(mstrPrivs, "取消报到") > 0)
                control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And intStep <> 2) '被拒绝的不能被再次拒绝
                control.ToolTipText = IIf(intState <= 1, "取消登记", "取消报到")
                control.Caption = "取消"
            End If
        Case conMenu_Manage_Transfer   '关联影像(&C)
            If InStr(mstrPrivs, "图像关联") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '在2---5之间可用
            End If
        Case conMenu_Manage_Cancel   '取消关联(&B)
            If InStr(mstrPrivs, "图像关联") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_First, conMenu_Manage_Second, conMenu_Manage_Quality
            If InStr(mstrPrivs, "影像质控") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive '结果阴阳性(&X)
            If (InStr(GetInsidePrivs(p诊疗报告管理), "报告书写") <= 0 And InStr(GetInsidePrivs(p诊疗报告管理), "报告修订") <= 0) Or _
                mblnIgnoreResult Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '在2---5之间可用
            End If
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '绿色通道标记/取消
            If InStr(mstrPrivs, "绿色通道") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '在2---5之间可用
            End If
        Case conMenu_Manage_Finish   '无报告完成(&F)
            If InStr(mstrPrivs, "无报告完成") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 2 Or intState = 3
            End If
        Case conMenu_Manage_ClearUp   '无报告回退(&U)
            If InStr(mstrPrivs, "无报告完成") <= 0 Then
                control.Visible = False
            ElseIf intState = 5 Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.报告ID).Value) = 0
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Complete   '检查完成(&E)
            If InStr(mstrPrivs, "检查完成") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState = 4 Or intState = 5)
            End If
        Case conMenu_Manage_Undone   '取消完成(&U)
            If InStr(mstrPrivs, "取消检查完成") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 6
            End If
        Case conMenu_File_SendImg  '发送图像
            If InStr(mstrPrivs, "文件发送") <= 0 Then control.Visible = False
        Case conMenu_Img_Contrast, conMenu_Img_Look     '影像对比,影像观片
            If blnNoRecord Then control.Enabled = False: Exit Sub
            If mblnIsHistory = True Then
                control.Enabled = mstrHStudyUID <> ""
            Else
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
            End If
            If control.Parent.Type <> xtpControlPopup Then control.Visible = control.Enabled
'        Case conMenu_Img_Look      '影像观片
'            If blnNoRecord Then control.Enabled = False: Exit Sub
'
'            If control.Parent.Type <> xtpControlPopup Then
'                control.Visible = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
'                control.Enabled = control.Visible
'            Else
'                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
'            End If
        Case conMenu_Img_3D     '三维重建
            If InStr(mstrPrivs, "三维重建操作") <> 0 And mblnUse3D = True Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                If blnNoRecord Then control.Enabled = False: Exit Sub
                If control.Parent.Type <> xtpControlPopup Then
                    control.Visible = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
                    control.Enabled = control.Visible
                Else
                    control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
                End If
            End If
        Case conMenu_Img_Delete '清除图像
            If InStr(mstrPrivs, "清除图像") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Img_Query ',获取图像
            If InStr(mstrPrivs, "清除图像") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState > 1
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ChangeDevice    '更改影像设备
            If blnNoRecord = True Then
                control.Enabled = False
            Else
                If UCase(Nvl(rptList.FocusedRow.Record(mcol.影像类别).Value)) = "CR" Or _
                    UCase(Nvl(rptList.FocusedRow.Record(mcol.影像类别).Value)) = "DR" Or _
                    UCase(Nvl(rptList.FocusedRow.Record(mcol.影像类别).Value)) = "DX" Or _
                    UCase(Nvl(rptList.FocusedRow.Record(mcol.影像类别).Value)) = "RF" Then
                    control.Enabled = True
                Else
                    control.Enabled = False
                End If
            End If
        Case conMenu_File_PrintSet     '打印设置(&S)
        Case conMenu_File_Preview, conMenu_File_Print '报告预览(&V) 报告打印(&P)
            control.Enabled = rptList.Records.Count > 0
        Case conMenu_File_Excel         '清单打印(&L)
            control.Enabled = rptList.Records.Count > 0
        Case conMenu_File_BatPrint    ' 批量打印(&B)
            control.Enabled = rptList.Records.Count > 0
        Case conMenu_File_Parameter     '参数设置(&O)
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '报表
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_File_Exit
        Case conMenu_View_ToolBar
        Case Else
            If control.Category <> "Main" Then
                Select Case TabWindow.Selected.Tag
                    Case "报告填写"
                        If mblnPacsReport = True Then
                            mfrmPacsReport.zlUpdateCommandBars control
                        Else
                            mobjReport.zlUpdateCommandBars control
                        End If
                    Case "申请费用"
                        mobjExpense.zlUpdateCommandBars control
                    Case "住院医嘱"
                        mobjInAdvice.zlUpdateCommandBars control
                    Case "门诊医嘱"
                        mobjOutAdvice.zlUpdateCommandBars control
                    Case "住院病历"
                        mobjInEPRs.zlUpdateCommandBars control
                    Case "门诊病历"
                        mobjOutEPRs.zlUpdateCommandBars control
                End Select

                If Not blnNoRecord Then
                    '删除只能在已报告和进行中可用
                    If control.ID = conMenu_Edit_Delete And rptList.FocusedRow.Record(mcol.检查状态).Value >= 4 Then
                        control.Enabled = False
                    End If
                    '当前查看的是历次记录则菜单均不可用
                    If cboTimes.ListIndex <> -1 Then
                        If rptList.FocusedRow.Record(mcol.医嘱ID).Value <> cboTimes.ItemData(cboTimes.ListIndex) Then control.Enabled = False
                    End If
                    '已完成除查阅,以及医嘱中报告查看打印，观片菜单外均不可用
                    If rptList.FocusedRow.Record(mcol.检查状态).Value = 6 Then
                        Select Case control.ID
                            Case conMenu_Edit_MarkMap, conMenu_File_Open, conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                                control.Enabled = True
                            Case Else
                                control.Enabled = False
                        End Select
                    End If
                End If
            End If
    End Select
End Sub

Private Sub chkSource_Click(Index As Integer)
    If Not mblnInitOK Then Exit Sub
    Call RefreshRptlist
End Sub



Private Sub Menu_Manage_PACS报告()
    Dim i As Integer
    
    If Not rptList.FocusedRow Is Nothing Then
        If Not mfrmPacsReportDock Is Nothing Then
            '先判断当前窗体是否是需要打开的窗体，如果不是，则查找窗口数组
            If rptList.FocusedRow.Record(mcol.医嘱ID).Value = mfrmPacsReportDock.mlngAdviceID Then
                '当前mfrmPacsReportDock指向的窗体，就是需要打开的窗体
                mfrmPacsReportDock.WindowState = 0  'normal
                mfrmPacsReportDock.ZOrder
                Exit Sub
            End If
        End If
        
        '查找窗口数组,找到需要打开的窗体，则通过Zorder把窗体显示到最前面
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = 1 To UBound(mobjPacsReportArry)
                If rptList.FocusedRow.Record(mcol.医嘱ID).Value = mobjPacsReportArry(i).mlngAdviceID Then
                    Set mfrmPacsReportDock = mobjPacsReportArry(i)
                    mfrmPacsReportDock.WindowState = 0  'normal
                    mfrmPacsReportDock.ZOrder
                    Exit Sub
                End If
            Next i
        End If
        
        '没有找到需要打开的窗体，且打开新窗体,并记录当前窗体
        Set mfrmPacsReportDock = New frmReport
        mfrmPacsReportDock.zlEditReport rptList.FocusedRow.Record(mcol.医嘱ID).Value, rptList.FocusedRow.Record(mcol.发送号).Value, mlngCur科室ID, Me, mstrPrivs, mlngModul, rptList.FocusedRow.Record(mcol.转出).Value = 1
        
        If SafeArrayGetDim(mobjPacsReportArry) = 0 Then
            ReDim mobjPacsReportArry(1) As frmReport
        Else
            ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) + 1) As frmReport
        End If
        Set mobjPacsReportArry(UBound(mobjPacsReportArry)) = mfrmPacsReportDock
    End If
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picList.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = PicWindow.Hwnd
    End If
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs           '权限
    mlngModul = glngModul           '模块号
    mlngCur科室ID = 0
    mstrCur科室 = ""
    mstrCanUse科室 = ""
    mstrCurFindtype = "就诊卡"
    mblnInitOK = False  '初始数据,初始化完成之前不进行数据的提取
    Call InitLocalPars '本地注册表参数
    If Not InitDepts Then Unload Me: Exit Sub '初始化医技科室
    Call InitMvar '初始化模块级变量
    '初始子窗体
    Set mfrmPACSImg = New frmPACSImg
    Set mfrmPacsReport = New frmReport  'PACS报告
    Set mobjReport = New zlRichEPR.cDockReport
    Set mobjExpense = New zlCISKernel.clsDockExpense
    Set mobjInAdvice = New zlCISKernel.clsDockInAdvices
    Set mobjOutAdvice = New zlCISKernel.clsDockOutAdvices
    Set mobjInEPRs = New zlRichEPR.cDockInEPRs
    Set mobjOutEPRs = New zlRichEPR.cDockOutEPRs
    Set mobjPacsCore = New zl9PacsCore.clsViewer
    
    Call InitFilterCmd
    Call InitCommandBars
    Call InitSubForm
    Call InitFaceScheme
    Call InitRptList

    Set mfrmPACSImg.pobjPacsCore = mobjPacsCore
    '去掉PACS报告窗体的控制框
    FormSetCaption mfrmPacsReport, False, False
    mblnInitOK = True '初始化完成
    Call RefreshRptlist
    
    Call RestoreWinState(Me, App.ProductName)
    
    ClearCacheFolder App.Path & "\TmpImage\"    '若临时目录满了，则清空该目录
    
    mstrUserNameHIS = UserInfo.姓名
    Me.stbThis.Panels(3).Text = "报告医生：" & mstrUserNameHIS
    ReDim mobjPacsReportArry(0) As frmReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", IIf(mblncmd门诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", IIf(mblncmd住院, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", IIf(mblncmd外诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", IIf(mblncmd体检, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", IIf(mblncmd已缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", IIf(mblncmd未缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", IIf(mblncmd登记, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", IIf(mblncmd报到, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", IIf(mblncmd报告, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", IIf(mblncmd审核, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", IIf(mblncmd完成, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", mstrCurFindtype
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    
    '判断嵌入式报告编辑器中的报告是否没有保存
    If mblnPacsReport = True Then    '使用PACS报告编辑器
        Call mfrmPacsReport.PromptModify
    End If
    
    Unload mfrmPACSImg
    Unload mfrmPacsReport
    Unload mobjReport.zlGetForm
    Unload mobjExpense.zlGetForm
    Unload mobjInAdvice.zlGetForm
    Unload mobjOutAdvice.zlGetForm
    Unload mobjInEPRs.zlGetForm
    Unload mobjOutEPRs.zlGetForm
    If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    
    Set mobjIDCard = Nothing
    Set mfrmPacsReport = Nothing
    Set mobjReport = Nothing
    Set mobjExpense = Nothing
    Set mobjInAdvice = Nothing
    Set mobjOutAdvice = Nothing
    Set mobjInEPRs = Nothing
    Set mobjOutEPRs = Nothing
    Set mobjPacsCore = Nothing
    
    '如果有三维重建，关闭三维重建的窗体
    If mblnUse3D = True Then
        On Error Resume Next
        Call sub3DProcess("EXIT")
    End If
End Sub

Private Sub mfrmPacsReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub mfrmPacsReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID)
End Sub

Private Sub mfrmPacsReport_BeforeEdit()
Dim lngOrderID As Long

    On Error GoTo ErrHandle
    lngOrderID = rptList.FocusedRow.Record(mcol.医嘱ID).Value
    If CheckConcurrentReport(lngOrderID) Then '检查是否有人正在操作报告
        Call UpdateReporter(lngOrderID, UserInfo.姓名)
    Else
        Call mfrmPacsReport.PromptModify(True)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPacsReportDock_AfterOpen()
    Call AfterReportOpen
End Sub

Private Sub mfrmPacsReportDock_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If Txt标识号.Text = "" And Me.ActiveControl Is Txt标识号 Then
        IDKind.IDKind = IDKinds.C2身份证号
        mstrCurFindtype = "身份证"
        Txt标识号 = strID
        Call Txt标识号_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub mobjInAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lng医嘱ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select 医嘱ID FROM 病人医嘱报告 where 病历ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取医嘱ID", CLng(报告ID))
    If rsTemp.EOF Then Exit Sub
    
    lng医嘱ID = Nvl(rsTemp!医嘱ID, 0)
    mobjReport.zlRefresh lng医嘱ID, mlngCur科室ID, False '以不可Edit方式刷新对像
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '调用查阅报告
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjInAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    '超过100张图像的序列，默认每隔5张传一张
    Call OpenViewer(mobjPacsCore, 医嘱ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjOutAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lng医嘱ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select 医嘱ID FROM 病人医嘱报告 where 病历ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取医嘱ID", CLng(报告ID))
    If rsTemp.EOF Then Exit Sub
    
    lng医嘱ID = Nvl(rsTemp!医嘱ID, 0)
    mobjReport.zlRefresh lng医嘱ID, mlngCur科室ID, False '以不可Edit方式刷新对像
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '调用查阅报告
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjOutAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    '超过100张图像的序列，默认每隔5张传一张
    Call OpenViewer(mobjPacsCore, 医嘱ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    If mblnPacsReport = True Then
        mfrmPacsReport.RefPacsPic '刷新图片
        If Not mfrmPacsReportDock Is Nothing Then
            mfrmPacsReportDock.RefPacsPic '刷新图片
        End If
    Else
        mobjReport.RefPacsPic '刷新图片
    End If
End Sub

Private Sub mobjReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub
Public Sub EditorClosed(ByVal lngOrderID As Long)
    Dim i As Integer
    Dim j As Integer
    
    Call UpdateReporter(lngOrderID, "")
    '处理PACS报告编辑器的窗口数组
    On Error Resume Next
    If mblnPacsReport = True Then
        '查找窗口数组，找到对应的窗口并删除
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = 1 To UBound(mobjPacsReportArry)
                If mobjPacsReportArry(i).mlngAdviceID = lngOrderID Then
                    '从数组中删除
                    For j = i To UBound(mobjPacsReportArry)
                        Set mobjPacsReportArry(j) = mobjPacsReportArry(j + 1)
                    Next j
                    ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) - 1) As frmReport
                    Exit For
                End If
            Next i
        End If
        
        If Not mfrmPacsReportDock Is Nothing Then
            If lngOrderID = mfrmPacsReportDock.mlngAdviceID Then
                '关闭当前报告窗口，将当前窗口设置成空
                Set mfrmPacsReportDock = Nothing
            End If
        End If
    End If
End Sub

Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
    On Error GoTo ErrHandle
    gstrSQL = "ZL_影像报告标记_Clear(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "清空标记"
    Call RefreshRptlist
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterOpen(ByVal intEditType As zlRichEPR.EditTypeEnum)
    Call AfterReportOpen
End Sub

Private Sub AfterReportOpen()
Dim lngOrderID As Long
    On Error GoTo ErrHandle
    lngOrderID = rptList.FocusedRow.Record(mcol.医嘱ID).Value
    
    Call UpdateReporter(lngOrderID, UserInfo.姓名)
    
    If mblnShowImgAtReport And Nvl(rptList.FocusedRow.Record(mcol.检查UID).Value) <> "" Then
        Dim intImageInverval As Integer
        
        intImageInverval = Val(mfrmPACSImg.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        Call OpenViewer(mobjPacsCore, lngOrderID, False, Me, , , mblnLocalizerBackward, intImageInverval)
    End If
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub
Public Sub AfterPrinted(lngOrderID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "ZL_影像报告打印_Update(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "更新打印标记"
    
    If Not mblnIgnoreResult And mintResultInput = 2 Then
        strSQL = "Select 结果阳性  From  病人医嘱发送 Where 医嘱id= [1]"
        Set rsTemp = OpenSQLRecord(strSQL, "提取结果阳性", lngOrderID)
        
        If IsNull(rsTemp!结果阳性) Then  '在报告时提示结果阴阳性
            Call PromptResult(lngOrderID, mlngModul, Me)
        End If
    End If
    
    If mblnPrintCommit = True Then
        Call Menu_Manage_检查最终完成(lngOrderID, False)
    End If
    
    Call RefreshRptlist
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID)
End Sub

Public Sub AfterReportSaved(lngOrderID As Long)
    Dim rsTemp As ADODB.Recordset, i As Integer, intState As Integer, lngSendId As Long
    If mblnPacsReport = True Then
'        mfrmPacsReport.zlRefresh 0, 0, 0
    Else
        mobjReport.zlRefresh 0, mlngCur科室ID, False
    End If

    gstrSQL = "Select Distinct A.医嘱id, B.ID,B.创建人,B.保存人,B.签名级别, B.完成时间, B.最后版本, C.发送号,C.结果阳性, D.检查UID " & vbNewLine & _
                "From 病人医嘱报告 A, 电子病历记录 B, 病人医嘱发送 C,影像检查记录 D " & vbNewLine & _
                "Where A.医嘱id =[1] And A.病历id = B.ID And A.医嘱id = C.医嘱id AND D.医嘱id = C.医嘱id"
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取是否签名", CLng(lngOrderID))
    If rsTemp.EOF Then Exit Sub
    lngSendId = rsTemp!发送号
    
    If Nvl(rsTemp!完成时间, "") = "" And rsTemp!最后版本 = 1 Then '未签名保存 或最后一次医师退签
        gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & "," & IIf(Nvl(rsTemp!检查UID) = "", 2, 3) & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "改为进行时"
        gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & Nvl(rsTemp!保存人, rsTemp!创建人) & "','')"
        zlDatabase.ExecuteProcedure gstrSQL, "保存报告人"
        intState = IIf(Nvl(rsTemp!检查UID) = "", 2, 3)
    Else
        If rsTemp!签名级别 < 2 Then '最后一次签名为医师,有可能的情况 1-医师第N次签名 2-主任级别最后一次退签 3-修订模式下保存(签名级别=0)
            gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & ",4)"
            zlDatabase.ExecuteProcedure gstrSQL, "改为报告时"
            
            intState = 4
        Else                        '主任及以上级别签名
            gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & ",5)"
            zlDatabase.ExecuteProcedure gstrSQL, "改为审核时"

            intState = 5
            If mblnCompleteCommit Then
                intState = 6
                Call Menu_Manage_检查最终完成(lngOrderID, False)
            End If
        End If
        
        gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & IIf(rsTemp!签名级别 = 1, Nvl(rsTemp!保存人), IIf(rsTemp!最后版本 = 1, Nvl(rsTemp!保存人), "")) & "','" & IIf(rsTemp!签名级别 = 1, "", Nvl(rsTemp!保存人)) & "')"
        zlDatabase.ExecuteProcedure gstrSQL, "保存复核人" '签名级别＝１表示是医生签名,无论是第N次，此时，报告人需要保存，复制人需要清空;其它情况报告人传空（过程中有处理），复核人填值
    
        If Not mblnIgnoreResult And IsNull(rsTemp!结果阳性) Then  '在报告时提示结果阴阳性
            If mblnReportWithResult Then '无影像诊断为阴性  -无提示自动标记
                gstrSQL = "ZL_影像检查_结果(" & lngOrderID & ",0)"
                zlDatabase.ExecuteProcedure gstrSQL, "标记阴阳性"
            ElseIf mintResultInput = 1 Then
                Call PromptResult(lngOrderID, mlngModul, Me)
            End If
        End If
    End If

    Call StateCheck(intState)
End Sub

Private Sub StateCheck(ByVal intState As Integer)
Dim cbrControl As CommandBarControl
    Select Case intState '跟据病人新状态确定新状态过滤是否选中
        Case 0, 1
            If Not mblncmd登记 Then Set cbrControl = Me.cbrdock.FindControl(, ID_登记)
        Case 2, 3
            If Not mblncmd报到 Then Set cbrControl = Me.cbrdock.FindControl(, ID_报到)
        Case 4
            If Not mblncmd报告 Then Set cbrControl = Me.cbrdock.FindControl(, ID_报告)
        Case 5
            If Not mblncmd审核 Then Set cbrControl = Me.cbrdock.FindControl(, ID_审核)
        Case 6
            If Not mblncmd完成 Then Set cbrControl = Me.cbrdock.FindControl(, ID_完成)
    End Select
    If mblnPatTrack Then
        If Not cbrControl Is Nothing Then '触发选中,选中触发列表刷新同时实现跟踪
            cbrdock_Execute cbrControl
        Else
            Call RefreshRptlist
        End If
    Else '不跟踪只刷新列表
        Call RefreshRptlist
    End If
End Sub
Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'功能：显示当前执行医嘱可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
        
    On Error GoTo errH
    
    objPopup.CommandBar.Controls.DeleteAll
    With rptList.FocusedRow
        gstrSQL = "Select Distinct C.编号,C.名称,C.说明" & _
            " From 病人医嘱记录 A,病历单据应用 B,病历文件列表 C" & _
            " Where A.ID=[1] And A.相关ID IS NULL" & _
            " And A.诊疗项目ID=B.诊疗项目ID" & _
            " And B.应用场合=[2] And B.病历文件ID=C.ID And C.种类=7" & _
            " Order by C.编号"
        If .Record(mcol.转出).Value = 1 Then
            gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
            gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(.Record(mcol.医嘱ID).Value), CLng(Decode(Nvl(.Record(mcol.来源).Value, "门"), "门", 1, "住院", 2, "外", 3, 4)))
    End With
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!名称 & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub FuncBillPrint(objControl As CommandBarControl)
'功能：打印诊疗单据
    On Error GoTo errH
    If objControl.Parameter = "" Then '奇怪，直接按F10时，是一个空的Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.Parameter = "" Then Exit Sub
    
    With rptList.FocusedRow
        If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
            Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & .Record(mcol.NO).Value, "性质=" & .Record(mcol.记录性质).Value, 1)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RefreshRptlist()
Dim i As Integer, lngcur医嘱ID As Long
    If rptList.Records.Count >= 1 Then
        lngcur医嘱ID = rptList.FocusedRow.Record(mcol.医嘱ID).Value
    End If
    Call LoadPatiList
    If lngcur医嘱ID = 0 Then
        If rptList.Records.Count >= 1 And rptList.Rows.Count >= 1 Then
            rptList.FocusedRow = rptList.Rows(0)
        End If
        Exit Sub
    End If
    
    
    '有记录时要重新定位回之前记录
    For i = 0 To rptList.Records.Count - 1
        If lngcur医嘱ID = rptList.Rows(i).Record.Item(mcol.医嘱ID).Value Then
            rptList.FocusedRow = rptList.Rows(i)
            Exit Sub
        End If
    Next
    '没能定位之前的记录，则定位到第0条
    If rptList.Records.Count >= 1 Then
        rptList.FocusedRow = rptList.Rows(0)
    End If
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraRegist.Left = 0
    fraRegist.Top = -75
    fraInfo.Top = -75
    fraInfo.Left = fraRegist.Left + fraRegist.Width
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left
    
    lblCash.Top = (picInfo.ScaleHeight - lblCash.Height) / 2 - fraInfo.Top
    lblCash.Left = fraInfo.Width - lblCash.Width - 100

    lbl个人信息.Width = lblCash.Left
    lbl检查信息.Width = lblCash.Left
End Sub

Private Sub LoadPatiList()
'功能：读取当前医技科室的执行医嘱(病人)清单
'blnLocate：是否刷新当前状态卡数据
Dim strSQL As String, strSQLBak As String, i As Long, rsPatList As ADODB.Recordset
Dim str来源 As String
Dim strFilter As String

    
    If Not mblnInitOK Then Exit Sub      '初始化未完成
    
    
    On Error GoTo ErrHandle
        
    '病人来源权限:(1-门诊,2-住院,3-外来,4-体检)
    If mblncmd门诊 Then str来源 = "1,"
    If mblncmd住院 Then str来源 = str来源 & "2,"
    If mblncmd外诊 Then str来源 = str来源 & "3,"
    If mblncmd体检 Then str来源 = str来源 & "4,"
    If InStr(Len(str来源), str来源, ",") > 0 Then str来源 = Mid(str来源, 1, Len(str来源) - 1) '去掉最后的逗号,两边逗号在参数中增加
        
    '发送时间
    If mdatFEnd <> CDate(0) Then
        strFilter = " And " & IIf(mblncmd登记, "A.发送时间", IIf(mDatType = 2, "A.发送时间", "A.首次时间")) & " Between [1] and [2] "
    Else '缺省查询条件
        strFilter = " And " & IIf(mblncmd登记, "A.发送时间", IIf(mDatType = 2, "A.发送时间", "A.首次时间")) & " Between [1] and Sysdate+1/(24*3600) "
    End If
    '单据号
    If mstrFNO <> "" Then
        strFilter = strFilter & " And A.NO= [3] "
    End If

    '病人科室
    If mlngF科室ID <> 0 Then
        strFilter = strFilter & " And B.病人科室ID+0= [4] "
    End If

    '病人标识

    If mstrF标识号 <> 0 Then
        strFilter = strFilter & " And Decode(B.病人来源,2,D.住院号,D.门诊号)= [5] "
    End If

    If mstrF就诊卡 <> "" Then
        strFilter = strFilter & " And D.就诊卡号 = [6] "
    End If

    If mstrF姓名 <> "" Then
        strFilter = strFilter & " And Instr(D.姓名 , [7])>0 "
    End If

    If mstr标本部位 <> "" Then
        strFilter = strFilter & " And instr(B.医嘱内容,[12])>0"
    End If

    If mdblFChkNO <> 0 Then
        strFilter = strFilter & " And H.检查号=[11] "
    End If
    
    If mbln结果阳性 Then
        strFilter = strFilter & " And Nvl(a.结果阳性, 0)=1"
    End If
    
    If mstr诊断医生 <> "" Then
        strFilter = strFilter & " And h.报告人=[13] "
    End If
    
    If mstr审核医生 <> "" Then
        strFilter = strFilter & " And h.复核人=[14] "
    End If
    
    '检查过程
    If mstr检查过程 <> "" Then
        If mstr检查过程 = "全部" Then
        
        ElseIf mstr检查过程 = "已登记" Then
            strFilter = strFilter & " And (a.执行过程 =0 or a.执行过程=1 Or a.执行过程 Is Null) "
        ElseIf mstr检查过程 = "已报到" Then
            strFilter = strFilter & " And (a.执行过程 = 2 and h.报告人 is null) "
        ElseIf mstr检查过程 = "已检查" Then
            strFilter = strFilter & " And (a.执行过程 = 3 and h.报告人 is null) "
        ElseIf mstr检查过程 = "处理中" Then
            strFilter = strFilter & " And (not h.报告操作 is null) "
        ElseIf mstr检查过程 = "报告中" Then
            strFilter = strFilter & " And ((a.执行过程 =2 or a.执行过程=3) and not h.报告人 is null and h.报告操作 is null) "
        ElseIf mstr检查过程 = "已报告" Then
            strFilter = strFilter & " And (a.执行过程=4 and h.复核人 is null) "
        ElseIf mstr检查过程 = "审核中" Then
            strFilter = strFilter & " And (a.执行过程=4 and not h.复核人 is null) "
        ElseIf mstr检查过程 = "已审核" Then
            strFilter = strFilter & " And a.执行过程=5 "
        ElseIf mstr检查过程 = "已完成" Then
            strFilter = strFilter & " And a.执行过程=6 "
        End If
    End If
    
    If mstr疾病诊断 <> "" Then
        strFilter = strFilter & " And F.病历ID IN(Select Distinct A.Id From 电子病历记录 A,电子病历内容 B Where A.创建时间>[1] AND A.Id=B.文件ID And instr(B.内容文本,[15])>0)"
    End If
    
    If mstr影像质量 <> "" Then
        strFilter = strFilter & " And h.影像质量=[16]"
    End If
    
    If mstr检查技师 <> "" Then
        strFilter = strFilter & " And h.检查技师=[17]"
    End If
    
    If mstrRoom <> "" Then
        If Not mblncmd登记 Then
            strFilter = strFilter & " And Instr([10],','|| A.执行间 || ',' )>0"
        Else
            strFilter = strFilter & " And Instr([10],','|| A.执行间 || ',' )>0"
        End If
    End If
    
    If mblnNoShowCancel Then '不显示取消登记的检查
        strFilter = strFilter & " And A.执行状态<>2 "
    End If

    '影像类别
    If mstr影像类别 <> "" Then
        strFilter = strFilter & " And h.影像类别=[18] "
    End If
    
    '增加PACS报告检索条件
    If mstr检查所见 <> "" Or mstr诊断意见 <> "" Or mstr建议 <> "" Then
        Dim strSubFilter As String
        If mstr检查所见 <> "" Then
            strSubFilter = " (b.内容文本 ='检查所见' And Instr(c.内容文本, [19]) > 0)"
        End If
        
        If mstr诊断意见 <> "" Then
            If strSubFilter = "" Then
                strSubFilter = " (b.内容文本 ='诊断意见' And Instr(c.内容文本, [20]) > 0)"
            Else
                strSubFilter = strSubFilter & " or (b.内容文本 ='诊断意见' And Instr(c.内容文本, [20]) > 0)"
            End If
        End If
        
        If mstr建议 <> "" Then
            If strSubFilter = "" Then
                strSubFilter = " (b.内容文本 ='建议' And Instr(c.内容文本, [21]) > 0)"
            Else
                strSubFilter = strSubFilter & " or (b.内容文本 ='建议' And Instr(c.内容文本, [21]) > 0)"
            End If
        End If
        
        strSubFilter = " (" & strSubFilter & ")"
        
        
        strFilter = strFilter & " And F.病历ID IN(Select Distinct a.Id From 电子病历记录 a, 电子病历内容 b,电子病历内容 c " _
            & " Where a.创建时间 > [1] And a.Id = b.文件id And b.Id = C.父ID And b.对象类型 = 3 And c.对象类型 = 2 And c.终止版 = 0 and " _
            & strSubFilter & ")"
    End If
    

    
    strSQL = "Select /*+ Rule*/" & vbNewLine & _
                "Distinct a.医嘱id, a.发送号, a.首次时间 As 检查时间, a.发送时间 As 开嘱时间, a.No, a.记录性质, a.执行状态," & vbNewLine & _
                "         Nvl(a.执行过程, 0) As 检查状态, a.执行间, a.结果阳性 As 阳性, b.诊疗项目id, b.病人id, b.主页id," & vbNewLine & _
                "         b.挂号单 As 挂号单, b.病人科室id, Decode(b.病人来源, 1, '门', 2, '住院', 3, '外', 4, '体') As 来源, b.医嘱内容," & vbNewLine & _
                "         b.标本部位, Nvl(b.紧急标志, 0) 紧急标志, Nvl(b.婴儿, 0) 婴儿, c.名称 As 内容, d.姓名, d.性别, d.年龄," & vbNewLine & _
                "         d.身份证号, Decode(b.病人来源, 1, d.门诊号, 2, d.住院号, 4, d.门诊号, Null) As 标识号," & vbNewLine & _
                "         Nvl(d.费别, '普通') As 费别, d.当前病区id As 病区id, d.就诊卡号, e.名称 As 科室, Nvl(f.病历id, 0) As 报告id," & vbNewLine & _
                "         Nvl(h.身高, '') 身高, Nvl(h.体重, '') 体重, Nvl(h.检查号, '') As 检查号, Nvl(h.检查uid, '') As 检查uid,H.影像质量,h.检查技师, " & vbNewLine & _
                "         H.是否打印,H.报告操作,0 as 转出,h.影像类别,H.绿色通道,H.报告打印,H.报告人,H.复核人,a.发送人 as 登记人,h.报到人,h.完成人,d.当前床号,b.开嘱医生,h.接收日期 as 采图时间  " & vbNewLine & _
                " From 病人医嘱发送 a, 病人医嘱记录 b, 诊疗项目目录 c, 病人信息 d, 部门表 e, 病人医嘱报告 f, 影像检查记录 h,影像检查项目 G" & vbNewLine & _
                " Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And b.病人id = d.病人id And b.病人科室id = e.Id And" & vbNewLine & _
                "      a.医嘱id = h.医嘱id(+) And a.发送号 = h.发送号(+) And a.医嘱id = f.医嘱id(+) And B.诊疗项目ID=G.诊疗项目ID AND" & vbNewLine & _
                "      Instr([8],','||B.病人来源||',')> 0 And A.执行部门ID+0= [9] And" & _
                IIf(mbln本次, " (B.病人来源=2 And b.主页ID=d.住院次数 Or Nvl(B.病人来源,0)<>2) and ", "") & _
                "      B.相关ID is NULL " & strFilter
    '如果有数据转出则还要检索后备表
    If mblnMoved Then
        strSQLBak = strSQL
        strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
        strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
        strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
        strSQLBak = Replace(strSQLBak, "病人医嘱报告", "H病人医嘱报告")
        strSQLBak = Replace(strSQLBak, "电子病历记录", "H电子病历记录")
        strSQLBak = Replace(strSQLBak, "电子病历内容", "H电子病历内容")
        strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
        strSQL = strSQL & " Union ALL " & strSQLBak
    End If
    strSQL = "Select * From (" & strSQL & ") Order by 检查状态,检查时间,开嘱时间"
    
    Set rsPatList = OpenSQLRecord(strSQL, Me.Caption, CDate(Format(mdatFBegin, "yyyy-MM-dd HH:mm:00")), CDate(Format(mdatFEnd, "yyyy-MM-dd HH:mm:59")), _
                                mstrFNO, mlngF科室ID, mstrF标识号, mstrF就诊卡, mstrF姓名, "," & str来源 & ",", mlngCur科室ID, _
                                mstrRoom, mdblFChkNO, mstr标本部位, mstr诊断医生, mstr审核医生, mstr疾病诊断, mstr影像质量, mstr检查技师, mstr影像类别, mstr检查所见, mstr诊断意见, mstr建议)
    strFilter = ""
    If mblncmd登记 Then strFilter = "检查状态=0 or 检查状态=1 or "
    If mblncmd报到 Then strFilter = IIf(strFilter <> "", strFilter & "检查状态=2 or 检查状态=3 or ", "检查状态=2 or 检查状态=3 or ")
    If mblncmd报告 Then strFilter = IIf(strFilter <> "", strFilter & "检查状态=4 or ", "检查状态=4 or ")
    If mblncmd审核 Then strFilter = IIf(strFilter <> "", strFilter & "检查状态=5 or ", "检查状态=5 or ")
    If mblncmd完成 Then strFilter = IIf(strFilter <> "", strFilter & "检查状态=6 or ", "检查状态=6 or ")
    
    If strFilter = "" Then
        strFilter = "检查状态<0"
    Else
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
    End If
    rsPatList.Filter = strFilter
    Call RefreshPatList(rsPatList)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshPatList(ByVal rsPatList As ADODB.Recordset)
Dim rptRecord As ReportRecord, i As Long, j As Long, blnRowShow As Boolean
Dim rsTemp As New ADODB.Recordset, StrAdviceIds As String, strSQL As String
        
    On Error GoTo ErrHandle
    If Not mblnInitOK Then Exit Sub
    
    rptList.Records.DeleteAll
    If rsPatList.EOF Then
        rptList.Populate
        rptList_SelectionChanged
        Exit Sub
    End If
    
    If mblncmd已缴 Xor mblncmd未缴 Then '互斥选择
        gstrSQL = "Select /*+ RULE */" & vbNewLine & _
                    "Distinct A.医嘱id" & vbNewLine & _
                    "From 病人医嘱发送 A, 病人费用记录 B, Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C" & vbNewLine & _
                    "Where A.医嘱id = C.Column_Value And A.NO = B.NO And A.记录性质=B.记录性质 And B.记录状态 = 1"
        Do Until rsPatList.EOF
            StrAdviceIds = StrAdviceIds & "," & rsPatList!医嘱ID
            If Len(StrAdviceIds) > 3880 Or rsPatList.AbsolutePosition = rsPatList.RecordCount Then 'VARCHAR2最大长度4000
                StrAdviceIds = Mid(StrAdviceIds, 2)
                strSQL = strSQL & " Union " & Replace(gstrSQL, "[1]", "'" & StrAdviceIds & "'")
                StrAdviceIds = ""
            End If
            rsPatList.MoveNext
        Loop
        strSQL = Mid(strSQL, 8)
        Call zlDatabase.OpenRecordset(rsTemp, strSQL, "提取是否收费")
    End If
    
    rsPatList.MoveFirst
    Do Until rsPatList.EOF
        blnRowShow = False
        If mblncmd已缴 Xor mblncmd未缴 Then '互斥选择
            If Not rsTemp Is Nothing Then
                rsTemp.Filter = ""
                rsTemp.Filter = "医嘱ID=" & rsPatList!医嘱ID
                If mblncmd已缴 Then '仅过滤已缴
                    If Not rsTemp.EOF Then blnRowShow = True
                Else                '仅过滤未缴
                    If rsTemp.EOF Then blnRowShow = True
                End If
            End If
        Else
            blnRowShow = True
        End If
    
        If blnRowShow Then
            Set rptRecord = rptList.Records.Add
            For j = 0 To Me.rptList.Columns.Count + 1
                rptRecord.AddItem ""
                If rsPatList!检查状态 = 6 Then rptRecord.Item(j).BackColor = &HFF00&
                If rsPatList!执行状态 = 2 Then rptRecord.Item(j).BackColor = &HFFFF&
            Next

            rptRecord.Item(mcol.紧急).Value = IIf(rsPatList("紧急标志") = 0, "", "紧急")
                rptRecord.Item(mcol.紧急).Icon = IIf(rsPatList("紧急标志") = 0, -1, Me.imgList.ListImages("紧急").Index - 1)
            rptRecord.Item(mcol.来源).Value = IIf(rsPatList("来源") = "住院", "住院", rsPatList!来源)
                rptRecord.Item(mcol.来源).Icon = IIf(rsPatList("来源") = "住院", Me.imgList.ListImages("住院").Index - 1, -1)
            rptRecord.Item(mcol.阳性).Value = Decode(rsPatList!阳性, 0, "", 1, "阳性", rsPatList!阳性)
                rptRecord.Item(mcol.阳性).Icon = IIf(Nvl(rsPatList!阳性, 0) = 0, -1, Me.imgList.ListImages("阳性").Index - 1)
            rptRecord.Item(mcol.质量).Value = Nvl(rsPatList!影像质量)
            rptRecord.Item(mcol.绿色通道).Value = Nvl(rsPatList!绿色通道, 0)
            rptRecord.Item(mcol.姓名).Value = rsPatList("姓名")
                rptRecord.Item(mcol.姓名).Icon = IIf(rptRecord.Item(mcol.绿色通道).Value = 1, Me.imgList.ListImages("绿色通道").Index - 1, -1)
            rptRecord.Item(mcol.检查号).Value = Nvl(rsPatList!检查号)
                rptRecord.Item(mcol.检查号).Icon = IIf(Len(rsPatList("检查UID")) > 0, Me.imgList.ListImages("影像").Index - 1, -1)
            rptRecord.Item(mcol.标识号).Value = Nvl(rsPatList("标识号"))
            rptRecord.Item(mcol.性别).Value = Nvl(rsPatList("性别"))
            rptRecord.Item(mcol.年龄).Value = Nvl(rsPatList("年龄"))
            rptRecord.Item(mcol.检查过程).Value = IIf(rsPatList!执行状态 = 2, "已拒绝", Decode(Nvl(rsPatList!检查状态, 0), 0, "已登记", 1, "已登记", 2, IIf(Nvl(rsPatList!报告操作) <> "", "处理中", IIf(Nvl(rsPatList!报告人) = "", "已报到", "报告中")), 3, IIf(Nvl(rsPatList!报告操作) <> "", "处理中", IIf(Nvl(rsPatList!报告人) = "", "已检查", "报告中")), 4, IIf(Nvl(rsPatList!报告操作) <> "", "处理中", IIf(Nvl(rsPatList!复核人) <> "", "审核中", "已报告")), 5, "已审核", "已完成"))
             rptRecord.Item(mcol.内容).Value = rsPatList("内容")
            If InStr(Nvl(rsPatList!医嘱内容), ":") > 0 Then '新的模式保存在医嘱内容中信息是 名称,执行标记:部位(方法,方法),部位---
'                rptRecord.Item(mcol.内容).Value = Split(Split(rsPatList!医嘱内容, ":")(0), ",")(0)
                rptRecord.Item(mcol.部位).Value = Split(rsPatList!医嘱内容, ":")(1)
            Else
'                rptRecord.Item(mcol.内容).Value = Nvl(rsPatList!医嘱内容)
                rptRecord.Item(mcol.部位).Value = Nvl(rsPatList!标本部位)
            End If
            rptRecord.Item(mcol.执行间).Value = Nvl(rsPatList("执行间"))
            rptRecord.Item(mcol.检查时间).Value = Nvl(rsPatList("检查时间"))
            rptRecord.Item(mcol.开嘱时间).Value = Nvl(rsPatList("开嘱时间"))
            rptRecord.Item(mcol.采图时间).Value = Nvl(rsPatList("采图时间"))
            
            rptRecord.Item(mcol.费别).Value = Nvl(rsPatList("费别"))
            rptRecord.Item(mcol.病人科室).Value = rsPatList("科室")
            rptRecord.Item(mcol.就诊卡号).Value = Nvl(rsPatList!就诊卡号)
            rptRecord.Item(mcol.身份证号).Value = Nvl(rsPatList!身份证号)
            rptRecord.Item(mcol.IC卡).Value = Nvl(rsPatList!就诊卡号)
            
            rptRecord.Item(mcol.病人ID).Value = Nvl(rsPatList("病人ID"), 0)
            rptRecord.Item(mcol.主页ID).Value = Nvl(rsPatList!主页ID, 0)
            rptRecord.Item(mcol.病人科室ID).Value = Nvl(rsPatList("病人科室ID"), 0)
            rptRecord.Item(mcol.病区ID).Value = Nvl(rsPatList("病区ID"), 0)
            rptRecord.Item(mcol.挂号单).Value = Nvl(rsPatList("挂号单"))
            rptRecord.Item(mcol.医嘱ID).Value = rsPatList("医嘱ID")
            rptRecord.Item(mcol.发送号).Value = rsPatList("发送号")
            rptRecord.Item(mcol.诊疗项目ID).Value = rsPatList("诊疗项目ID")
            rptRecord.Item(mcol.床号).Value = Nvl(rsPatList("当前床号"))
            rptRecord.Item(mcol.开嘱医生).Value = Nvl(rsPatList("开嘱医生"))
            rptRecord.Item(mcol.打印胶片).Value = IIf(Nvl(rsPatList!是否打印, 0) = 0, "未打印", "已打印")
            rptRecord.Item(mcol.报告操作).Value = Nvl(rsPatList!报告操作)
            rptRecord.Item(mcol.报告打印).Value = IIf(Nvl(rsPatList!报告打印, 0) = 0, "未打印", "已打印")
            rptRecord.Item(mcol.报告人).Value = Nvl(rsPatList!报告人)
            rptRecord.Item(mcol.复核人).Value = Nvl(rsPatList!复核人)
            
            
            rptRecord.Item(mcol.NO).Value = rsPatList("NO")
            rptRecord.Item(mcol.记录性质).Value = Nvl(rsPatList!记录性质, 0)
            rptRecord.Item(mcol.医嘱内容).Value = Nvl(rsPatList("医嘱内容"))
            rptRecord.Item(mcol.检查UID).Value = rsPatList("检查UID")
            rptRecord.Item(mcol.检查状态).Value = rsPatList("检查状态")
            rptRecord.Item(mcol.婴儿).Value = rsPatList!婴儿
            rptRecord.Item(mcol.报告ID).Value = Nvl(rsPatList("报告ID"))
            rptRecord.Item(mcol.医嘱附件).Value = ""
            rptRecord.Item(mcol.执行状态).Value = rsPatList!执行状态
            rptRecord.Item(mcol.转出).Value = rsPatList!转出
            rptRecord.Item(mcol.身高).Value = Nvl(rsPatList!身高)
            rptRecord.Item(mcol.体重).Value = Nvl(rsPatList!体重)
            rptRecord.Item(mcol.检查技师).Value = Nvl(rsPatList!检查技师)
            rptRecord.Item(mcol.影像类别).Value = Nvl(rsPatList!影像类别)
            
            rptRecord.Item(mcol.登记人).Value = Nvl(rsPatList!登记人)
            rptRecord.Item(mcol.报到人).Value = Nvl(rsPatList!报到人)
            rptRecord.Item(mcol.完成人).Value = Nvl(rsPatList!完成人)
        End If
        rsPatList.MoveNext
    Loop
    rptList.Populate
'    If rptList.Records.Count > 0 Then
'            rptList.FocusedRow = rptList.Rows(0)
'    End If
    stbThis.Panels(2).Text = "共 " & rptList.Records.Count & " 条记录": stbThis.Panels(2).Alignment = sbrCenter
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub PicWindow_Resize()
    On Error Resume Next
    With picInfo
        .Top = 0
        .Left = 0
        .Width = PicWindow.ScaleWidth
    End With
        
    With TabWindow
        .Top = picInfo.ScaleHeight
        .Left = 0
        .Width = PicWindow.ScaleWidth
        .Height = PicWindow.ScaleHeight - picInfo.ScaleHeight
    End With
End Sub


Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Set Popup = cbrMain.Add("右键菜单", xtpBarPopup)
        For Each Menucontrol In cbrMain.ActiveMenuBar.Controls
            If (Menucontrol.ID <> conMenu_FilePopup And Menucontrol.ID <> conMenu_ToolPopup _
                And Menucontrol.ID <> conMenu_ViewPopup And Menucontrol.ID <> conMenu_HelpPopup _
                And Menucontrol.ID <> conMenu_View_Filter * 10# And Menucontrol.ID <> conMenu_View_FindType) And Menucontrol.Type = xtpControlPopup Then
                For Each control In Menucontrol.CommandBar.Controls
                    control.Copy Popup
                Next
            End If
        Next
        Popup.ShowPopup
    End If
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim blnNoRecord As Boolean

    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If

    If Not blnNoRecord Then
            Select Case rptList.FocusedRow.Record(mcol.检查状态).Value
                Case 1, 0
                    Call Menu_Manage_报到
                Case 2, 3               '双击打开书写报告,报告打开时跟据设定是否打开观片站
                    Call Menu_RichEPR(conMenu_Edit_Modify)
                Case 4, 5               '双击修订报告,报告打开时跟据设定是否打开观片站
                    Call Menu_RichEPR(conMenu_Edit_Audit)
                Case 6                  '查阅
                    Call Menu_RichEPR(conMenu_File_Open)
            End Select
    End If
End Sub

Private Sub rptList_SelectionChanged()
Dim blnNoRecord As Boolean, rsTemp As ADODB.Recordset, rptRecord As ReportRecord, str医嘱附件 As String, i As Integer
Dim blnShowReport As Boolean, strTemp As String
    
    mblnIsHistory = False
    
    blnShowReport = True
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    
    If Not blnNoRecord Then
        '判断 无图像不许写报告
        If mblnReportWithImage = True Then
            If rptList.FocusedRow.Record(mcol.检查UID).Value = "" Or IsNull(rptList.FocusedRow.Record(mcol.检查UID).Value) Then blnShowReport = False
        End If
        
        With rptList.FocusedRow
            lbl个人信息.Caption = "" 'cbotime中会用到，用于区别是listindex时触发还是点击cbotimes触发
            gstrSQL = "Select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                       " From 病人医嘱记录 A,病人医嘱发送 B,影像检查项目 C" & _
                       " Where A.病人id = [1] And A.相关id Is Null And A.执行科室id+0 =[2] And B.医嘱ID=A.ID " & _
                       "" & IIf(.Record(mcol.检查过程).Value = "已拒绝", "", " And B.执行状态<>2 ") & _
                       " AND A.诊疗项目ID=C.诊疗项目ID"
            strTemp = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
            strTemp = Replace(strTemp, "病人医嘱发送", "H病人医嘱发送")
            gstrSQL = gstrSQL & " Union ALL " & strTemp
            gstrSQL = "select * from (" & gstrSQL & ") Order By 开嘱时间 Asc"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(.Record(mcol.病人ID).Value), mlngCur科室ID)
            cboTimes.Clear
            Do Until rsTemp.EOF
               cboTimes.AddItem "第" & rsTemp.AbsolutePosition & "次(" & Format(rsTemp!开嘱时间, "yyyy-mm-dd") & ")  " & Trim(rsTemp!医嘱内容)
               cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!医嘱ID
               If rsTemp!医嘱ID = rptList.FocusedRow.Record(mcol.医嘱ID).Value Then cboTimes.ListIndex = cboTimes.NewIndex
               rsTemp.MoveNext
            Loop
            
            '判断嵌入式报告编辑器中的报告是否没有保存
            If mblnPacsReport = True Then    '使用PACS报告编辑器
                Call mfrmPacsReport.PromptModify
            End If
                
            '显示基本信息
            lbl个人信息.Caption = "姓  名:" & Rpad(.Record(mcol.姓名).Value, 12, " ") & "性  别:" & Rpad(.Record(mcol.性别).Value, 13, " ") & _
                                  "年  龄:" & Rpad(.Record(mcol.年龄).Value, 10, " ") & "标识号:" & Rpad(.Record(mcol.标识号).Value, 12, " ") & _
                                  "床  号:" & Rpad(.Record(mcol.床号).Value & "", 10, " ")
            lbl检查信息.Caption = "检查号:" & Rpad(Nvl(.Record(mcol.检查号).Value), 12, " ") & "病人科室:" & Rpad(Nvl(.Record(mcol.病人科室).Value), 11, " ") & _
                                  "开嘱医生:" & Rpad(Nvl(.Record(mcol.开嘱医生).Value), 8, " ") & "检查项目:" & .Record(mcol.内容).Value
            lblCash.Caption = "收": lblCash.Visible = False
            
            If Nvl(.Record(mcol.医嘱附件).Value) = "" Then
                gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
                If .Record(mcol.转出).Value = 1 Then
                    gstrSQL = Replace(gstrSQL, "病人医嘱附件", "H病人医嘱附件")
                End If
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人附件", CLng(.Record(mcol.医嘱ID).Value))
                Do Until rsTemp.EOF
                    str医嘱附件 = str医嘱附件 & rsTemp!项目 & ":" & Nvl(rsTemp!内容) & vbCrLf
                    rsTemp.MoveNext
                Loop
                rptList.FocusedRow.Record(mcol.医嘱附件).Value = str医嘱附件
            End If
            Txt基本信息 = ""
            If InStr(Nvl(.Record(mcol.医嘱内容).Value), ":") > 0 Then
                For i = 0 To UBound(Split(Split(.Record(mcol.医嘱内容).Value, ":")(1), "),"))
                    If i = 0 Then
                        Txt基本信息 = "检查部位:" & vbCrLf & Space(2) & "1:" & Split(Split(.Record(mcol.医嘱内容).Value, ":")(1), "),")(i) & ")"
                    Else
                        Txt基本信息 = Txt基本信息 & vbCrLf & Space(2) & i + 1 & ":" & Split(Split(.Record(mcol.医嘱内容).Value, ":")(1), "),")(i) & ")"
                    End If
                Next
                If Trim(Txt基本信息) <> "" Then Txt基本信息 = Mid(Txt基本信息, 1, Len(Txt基本信息) - 1)
            Else
                Txt基本信息 = "检查部位:" & Nvl(.Record(mcol.医嘱内容).Value)
            End If
            
            Txt基本信息 = Txt基本信息 & vbCrLf & vbCrLf & Nvl(.Record(mcol.医嘱附件).Value)
            lblCash.Visible = CheckChargeState(.Record(mcol.医嘱ID).Value) = 1
            
            '有记录时根据不同病人状态提供选项卡
            If .Record(mcol.来源).Value = "住院" Then '根据病人来源控制病历及医嘱选项卡
                For i = 0 To TabWindow.ItemCount - 1
                    Select Case TabWindow(i).Tag
                        Case "门诊病历", "门诊医嘱"
                            TabWindow(i).Visible = False
                        Case "住院病历", "住院医嘱"
                            TabWindow(i).Visible = True
                        Case "影像图象"
                            TabWindow(i).Visible = True
                        Case "报告填写"
                            TabWindow(i).Visible = .Record(mcol.检查状态).Value > 1 And blnShowReport
                    End Select
                Next
            Else
                For i = 0 To TabWindow.ItemCount - 1
                    Select Case TabWindow(i).Tag
                        Case "门诊病历", "门诊医嘱"
                            TabWindow(i).Visible = True
                        Case "住院病历", "住院医嘱"
                            TabWindow(i).Visible = False
                        Case "影像图象"
                            TabWindow(i).Visible = True
                        Case "报告填写"
                            TabWindow(i).Visible = .Record(mcol.检查状态).Value > 1 And blnShowReport
                    End Select
                Next
            End If
            
            If mstrFirstTab <> "" Then '不为空表示按定制首页显示
                For i = 0 To TabWindow.ItemCount - 1
                    If InStr(TabWindow.Item(i).Tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                        If TabWindow.Item(i).Selected Then
                            Call TabWindow_SelectedChanged(TabWindow.Item(i))
                        Else
                            TabWindow.Item(i).Selected = True
                        End If
                        Exit Sub
                    End If
                Next
                TabWindow(0).Selected = True '没循环到了触发第0个tab
            Else
                If TabWindow.Selected.Visible Then
                    Call TabWindow_SelectedChanged(TabWindow(TabWindow.Selected.Index))
                Else
                    Select Case TabWindow.Selected.Tag
                        Case "门诊病历"
                            For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                                If TabWindow(i).Tag = "住院病历" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                        Case "门诊医嘱"
                            For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                                If TabWindow(i).Tag = "住院医嘱" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                        Case "住院病历"
                            For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                                If TabWindow(i).Tag = "门诊病历" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                        Case "住院医嘱"
                            For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                                If TabWindow(i).Tag = "门诊医嘱" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                    End Select
                    TabWindow(0).Selected = True '没循环到了触发第0个tab
                End If
            End If
        End With
        '显示可打印的诊疗单据:之所以即时加载,是为了使用F2热键
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    Else
        If TabWindow(TabWindow.Selected.Index).Visible Then
            Call TabWindow_SelectedChanged(TabWindow(TabWindow.Selected.Index))
        Else
             TabWindow(0).Selected = True
        End If
        cboTimes.Clear
        Txt基本信息 = ""
        
        lbl个人信息.Caption = "姓  名:" & Space(12) & "性  别:" & Space(13) & "年  龄:" & Space(10) & "标识号:" & Space(12) & _
                                  "床  号:" & Space(10)
        lbl检查信息.Caption = "检查号:" & Space(12) & "病人科室:" & Space(11) & "开嘱医生:" & Space(8) & "检查项目:"
                                  
        lblCash.Visible = False
    End If
    
    On Error Resume Next
    If rptList.Visible = True Then rptList.SetFocus
    err.Clear
End Sub

Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim blnNoRecord As Boolean
Dim Menucontrol As CommandBarPopup, cbrControl As CommandBarButton, bcontrol As CommandBarControl, i As Integer
    
    If Not mblnInitOK Then Exit Sub
    
    On Error Resume Next
    '判断是否有记录
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    
    '以下指定的菜单在循环时循环不到,所以指定删除
    
    For Each Menucontrol In cbrMain.ActiveMenuBar.Controls '删除非主界面菜单,但是保留“报表”菜单
        If Menucontrol.ID <> conMenu_ReportPopup Then
            For Each bcontrol In Menucontrol.CommandBar.Controls
                If bcontrol.Category <> "Main" Then bcontrol.Delete
            Next
            If Menucontrol.Category <> "Main" Then Menucontrol.Delete
        End If
    Next
    
    Set Menucontrol = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    If Not Menucontrol Is Nothing Then
        Set cbrControl = Menucontrol.CommandBar.Controls.Find(, conMenu_View_Append)
        If Not cbrControl Is Nothing Then cbrControl.Delete
    End If
    
    Set bcontrol = cbrMain(2).Controls.Find(, conMenu_Edit_NewItem)
    If Not bcontrol Is Nothing Then bcontrol.Delete
    Set bcontrol = cbrMain(2).Controls.Find(, conMenu_Edit_Delete)
    If Not bcontrol Is Nothing Then bcontrol.Delete
    Set bcontrol = cbrMain(2).Controls.Find(, conMenu_Edit_Modify)
    If Not bcontrol Is Nothing Then bcontrol.Delete
    For Each bcontrol In cbrMain(2).Controls '删除主界面工具条
        If bcontrol.Category <> "Main" Then
            bcontrol.Delete
        End If
    Next
    
    err.Clear


    On Error GoTo ErrHandle
    If blnNoRecord Then
        Select Case Item.Tag
            Case "影像图象"
                Call mfrmPACSImg.zlRefresh(0, 0, mstrPrivs, False)
            Case "报告填写"
                If mblnPacsReport = True Then    '使用PACS报告编辑器
                    mfrmPacsReport.zlDefCommandBars Me.cbrMain
                    mfrmPacsReport.zlRefresh 0, 0, 0, mstrPrivs, mlngModul, Me, False
                Else
                    mobjReport.zlDefCommandBars Me.cbrMain
                    mobjReport.zlRefresh 0, mlngCur科室ID, False
                End If
            Case "申请费用"
                mobjExpense.zlDefCommandBars Me, Me.cbrMain
                mobjExpense.zlRefresh 0, 0, 0
            Case "住院医嘱"
                mobjInAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                mobjInAdvice.zlRefresh 0, 0, 0, 0, 0, False, 0, 0
            Case "门诊医嘱"
                mobjOutAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                mobjOutAdvice.zlRefresh 0, "", False
            Case "住院病历"
                mobjInEPRs.zlDefCommandBars cbrMain
                mobjInEPRs.zlRefresh 0, 0, 0, False
            Case "门诊病历"
                mobjOutEPRs.zlDefCommandBars cbrMain
                mobjOutEPRs.zlRefresh 0, 0, 0, False
        End Select
        Exit Sub
    End If
    
    
    With rptList.FocusedRow

        If cboTimes.ListIndex <> -1 Then
            If .Record(mcol.医嘱ID).Value <> cboTimes.ItemData(cboTimes.ListIndex) Then '当前医嘱ID与当前历次记录中医嘱ID不同时由CboTimes控制
                Call cboTimes_Click
                Exit Sub
            End If
        End If
        
        Select Case Item.Tag
            Case "影像图象"
                '不管当前记录没有图像，都对mfrmPACSImg进行强制刷新
                Call mfrmPACSImg.zlRefresh(.Record(mcol.医嘱ID).Value, .Record(mcol.发送号).Value, mstrPrivs, .Record(mcol.转出).Value = 1, True)
                '如果刷新出来有记录，则刷新病人列表
                If (IsNull(.Record(mcol.检查UID).Value) Or .Record(mcol.检查UID).Value = "") And mfrmPACSImg.lvwSeq.ListItems.Count > 0 Then
                    Call RefreshRptlist
                End If
            Case "报告填写"
                If mblnPacsReport = True Then
                    mfrmPacsReport.zlDefCommandBars Me.cbrMain
                    Call mfrmPacsReport.zlRefresh(.Record(mcol.医嘱ID).Value, .Record(mcol.发送号).Value, mlngCur科室ID, mstrPrivs, mlngModul, Me, .Record(mcol.转出).Value = 1)
                Else
                    mobjReport.zlDefCommandBars Me.cbrMain
                    Call mobjReport.zlRefresh(Nvl(.Record(mcol.医嘱ID).Value, 0), mlngCur科室ID, True)
                End If
            Case "申请费用"
                mobjExpense.zlDefCommandBars Me, Me.cbrMain
                mobjExpense.zlRefresh mlngCur科室ID, .Record(mcol.医嘱ID).Value, .Record(mcol.发送号).Value, .Record(mcol.转出).Value = 1
            Case "住院医嘱"
                mobjInAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                mobjInAdvice.zlRefresh .Record(mcol.病人ID).Value, Val(.Record(mcol.主页ID).Value), _
                    .Record(mcol.病区ID).Value, .Record(mcol.病人科室ID).Value, 0, .Record(mcol.转出).Value = 1, _
                    Nvl(.Record(mcol.医嘱ID).Value, 0), .Record(mcol.执行状态).Value
            Case "门诊医嘱"
                mobjOutAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                If .Record(mcol.挂号单).Value = "" Then
                    mobjOutAdvice.zlRefresh 0, "", False
                Else
                    mobjOutAdvice.zlRefresh .Record(mcol.病人ID).Value, .Record(mcol.挂号单).Value, True, .Record(mcol.转出).Value = 1, Nvl(.Record(mcol.医嘱ID).Value, 0)
                End If
            Case "住院病历"
                mobjInEPRs.zlDefCommandBars cbrMain
                mobjInEPRs.zlRefresh .Record(mcol.病人ID).Value, Val(.Record(mcol.主页ID).Value), .Record(mcol.病人科室ID).Value, False, .Record(mcol.转出).Value = 1
            Case "门诊病历"
                mobjOutEPRs.zlDefCommandBars cbrMain
                mobjOutEPRs.zlRefresh .Record(mcol.病人ID).Value, Val(.Record(mcol.主页ID).Value), .Record(mcol.病人科室ID).Value, False, .Record(mcol.转出).Value = 1
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub TimerRefresh_Timer()
    '刷新病人列表
    Call Menu_View_Refresh_click
End Sub

Private Sub txt标识号_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Txt标识号.Text = "" And Me.ActiveControl Is Txt标识号)
    If Txt标识号.Text = "" Then Txt标识号.Tag = ""
End Sub

Private Sub txt标识号_GotFocus()
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard         '身份证识别对象
    
    If Txt标识号.Text <> "" Then Call zlControl.TxtSelAll(Txt标识号)
    If mstrCurFindtype = "姓名" Then
        Call zlCommFun.OpenIme(True)
    End If
    If Not mobjIDCard Is Nothing And Txt标识号.Text = "" Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub Txt标识号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txt标识号_Validate(False)
        Call zlControl.TxtSelAll(Txt标识号)
        Call SeekNextPati(Txt标识号.Tag <> Txt标识号.Text)
    End If
End Sub

Private Sub txt标识号_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        Select Case mstrCurFindtype
            Case "标识号"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "就诊卡"
                Dim blnCard As Boolean
    
                '去掉磁卡的其他的特殊字符
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = InputIsCard(Me.Txt标识号, KeyAscii)
                
                '刷卡完成或确认输入
                If blnCard And Len(Me.Txt标识号.Text) = Val(gbytCardLen) - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.Txt标识号.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.Txt标识号.Text = Me.Txt标识号.Text & Chr(KeyAscii)
                        Me.Txt标识号.SelStart = Len(Me.Txt标识号.Text)
                    End If
                    KeyAscii = 0
                    Me.Txt标识号.Text = UCase(Me.Txt标识号)
                    Me.Txt标识号.SetFocus
                End If
            Case "单据号"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (Txt标识号.Text = "" Or Txt标识号.SelLength = Len(Txt标识号.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "姓名"
            
        End Select
    End If
End Sub

Private Sub Txt标识号_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txt标识号_Validate(Cancel As Boolean)
    If mstrCurFindtype = "单据号" Then
        If IsNumeric(Txt标识号.Text) Then
            Txt标识号.Text = GetFullNO(Txt标识号.Text, 0)
        End If
    End If
End Sub


Private Sub SeekNextPati(ByVal blnFirst As Boolean)
Dim blnOK As Boolean, l As Long, intB As Integer

    If rptList.FocusedRow Is Nothing Then '没记录
        Exit Sub
    ElseIf rptList.FocusedRow.GroupRow Then
        Exit Sub
    End If


    intB = 0
    If Not blnFirst Then
        intB = rptList.SelectedRows(l).Index + 1
        If intB > rptList.Records.Count Then intB = 0
    End If

    blnOK = False
    For l = intB To rptList.Rows.Count - 1 '在当前状态中查找
        Select Case mstrCurFindtype
            Case "标识号"
                If Nvl(rptList.Rows(l).Record.Item(mcol.标识号).Value, 0) = Txt标识号.Text Then blnOK = True
            Case "就诊卡", "ＩＣ卡"
                If Nvl(rptList.Rows(l).Record.Item(mcol.就诊卡号).Value, "") = Txt标识号.Text Then blnOK = True
            Case "单据号"
                If Nvl(rptList.Rows(l).Record.Item(mcol.NO).Value, "") = Txt标识号.Text Then blnOK = True
            Case "检查号"
                If Nvl(rptList.Rows(l).Record.Item(mcol.检查号).Value, "") = Txt标识号.Text Then blnOK = True
            Case "姓名"
                If Nvl(rptList.Rows(l).Record.Item(mcol.姓名).Value, "") Like Txt标识号.Text & "*" Then blnOK = True
                If zlCommFun.SpellCode(Nvl(rptList.Rows(l).Record.Item(mcol.姓名).Value, "")) Like UCase(Txt标识号.Text) & "*" Then blnOK = True
            Case "身份证"
                If Nvl(rptList.Rows(l).Record.Item(mcol.身份证号).Value, "") = Txt标识号.Text Then blnOK = True
        End Select
        
        If blnOK Then
            Txt标识号.Tag = Txt标识号.Text
            If rptList.FocusedRow.Index <> l Then     '若不是当前选中行则选中
                rptList.FocusedRow = rptList.Rows(l)
            End If
            rptList.SetFocus
            Exit Sub
        End If
    Next
End Sub

Private Sub Menu_Manage_更换检查设备()
    Dim strModality As String
    Dim rResult As VbMsgBoxResult
    Dim strSQL As String
    
    If rptList.FocusedRow Is Nothing Then '没记录
        Exit Sub
    ElseIf rptList.FocusedRow.GroupRow Then
        Exit Sub
    End If
    
    If UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.影像类别).Value)) = "CR" Or _
       UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.影像类别).Value)) = "DR" Or _
       UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.影像类别).Value)) = "DX" Or _
       UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.影像类别).Value)) = "RF" Then
       
       
       frmChangeDevice.ShowMe UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.影像类别).Value)), Me
       strModality = frmChangeDevice.strDeviceType
       
        If strModality <> "" Then
            strSQL = "Zl_影像检查_影像类别(" & rptList.FocusedRow.Record(mcol.医嘱ID).Value & "," & rptList.FocusedRow.Record(mcol.发送号).Value & ",'" & strModality & "')"
            ExecuteProc strSQL, Me.Caption
        End If
        
        '刷新病人列表
        Call RefreshRptlist
    End If
End Sub

Private Sub sub3DProcess(strCommand As String)
    Dim str3DCommand As String
    Dim str3DImgDir As String

    str3DImgDir = App.Path & "\TmpImage\3D\"

    '组织三维重建语句
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & str3DImgDir
    On Error Resume Next
    Shell str3DCommand
End Sub

Private Sub sub三维重建(strCommand As String)

    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.医嘱ID).Value, rptList.FocusedRow.Record(mcol.发送号).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.转出).Value = 1)
    End If
     
'    Call sub3DProcess("IDLE")
    '组织三维重建需要的图像
    Call mfrmPACSImg.zlMenuClick("三维重建")
    Call sub3DProcess(strCommand)
End Sub

