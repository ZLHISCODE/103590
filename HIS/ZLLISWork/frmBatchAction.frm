VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBatchAction 
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11100
   Icon            =   "frmBatchAction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11100
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   6540
      ScaleHeight     =   3945
      ScaleWidth      =   3825
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   900
      Width           =   3825
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2265
         Left            =   60
         TabIndex        =   33
         Top             =   300
         Width           =   3525
         _Version        =   589884
         _ExtentX        =   6218
         _ExtentY        =   3995
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox chkfilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "普通"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   32
         ToolTipText     =   "住院标本"
         Top             =   30
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkfilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "急诊"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   31
         ToolTipText     =   "门诊和直接登记标本"
         Top             =   30
         Value           =   1  'Checked
         Width           =   675
      End
   End
   Begin RichTextLib.RichTextBox RtfTxt 
      Height          =   585
      Left            =   6270
      TabIndex        =   25
      Top             =   5490
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1032
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmBatchAction.frx":6852
   End
   Begin VB.PictureBox picWhere 
      BorderStyle     =   0  'None
      Height          =   7065
      Left            =   210
      ScaleHeight     =   7065
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   330
      Width           =   5535
      Begin XtremeReportControl.ReportControl rptMachine 
         Height          =   3105
         Left            =   60
         TabIndex        =   17
         Top             =   3000
         Width           =   4965
         _Version        =   589884
         _ExtentX        =   8758
         _ExtentY        =   5477
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.OptionButton optSort 
         Caption         =   "病区"
         Height          =   180
         Index           =   2
         Left            =   4620
         TabIndex        =   39
         Top             =   1440
         Width           =   765
      End
      Begin VB.OptionButton optSort 
         Caption         =   "病人"
         Height          =   180
         Index           =   1
         Left            =   3930
         TabIndex        =   38
         Top             =   1440
         Width           =   705
      End
      Begin VB.OptionButton optSort 
         Caption         =   "标本"
         Height          =   180
         Index           =   0
         Left            =   3270
         TabIndex        =   36
         Top             =   1440
         Width           =   705
      End
      Begin VB.CheckBox chkAbnormal 
         Caption         =   "不显示异常结果标本"
         Height          =   180
         Left            =   60
         TabIndex        =   35
         Top             =   1740
         Width           =   2325
      End
      Begin VB.ComboBox cboDate 
         Height          =   300
         ItemData        =   "frmBatchAction.frx":68EF
         Left            =   90
         List            =   "frmBatchAction.frx":68F1
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   90
         Width           =   1095
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "包含打印过的标本"
         Height          =   180
         Left            =   3270
         TabIndex        =   29
         Top             =   1740
         Width           =   1755
      End
      Begin VB.CheckBox chkPatient 
         Caption         =   "同一个病人合并为一个报告单打印"
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   1980
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.ComboBox cbo病区 
         Height          =   300
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1605
      End
      Begin VB.ComboBox cbo来源 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox TxtModify 
         Height          =   285
         Left            =   1050
         MaxLength       =   15
         TabIndex        =   23
         Top             =   2385
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Frame Frame1 
         Height          =   105
         Left            =   -45
         TabIndex        =   21
         Top             =   2205
         Width           =   5415
      End
      Begin VB.CheckBox chkUnion 
         Caption         =   "不打印被合并的标本"
         Height          =   255
         Left            =   3270
         TabIndex        =   20
         Top             =   1980
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.ComboBox cboExeDept 
         Height          =   300
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1050
         Width           =   1605
      End
      Begin VB.ComboBox cboRequisitionDept 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1050
         Width           =   1425
      End
      Begin VB.ComboBox cboVerifyMan 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1395
         Width           =   1425
      End
      Begin VB.TextBox TxtSample 
         Height          =   285
         Left            =   3060
         TabIndex        =   7
         Top             =   405
         Width           =   2205
      End
      Begin VB.TextBox txtBatchNum 
         Height          =   285
         Left            =   1230
         TabIndex        =   6
         Top             =   405
         Width           =   915
      End
      Begin MSComCtl2.DTPicker DtpBegin 
         Height          =   285
         Left            =   1230
         TabIndex        =   2
         Top             =   90
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   149094403
         CurrentDate     =   39198
      End
      Begin MSComCtl2.DTPicker DtpEnd 
         Height          =   285
         Left            =   3330
         TabIndex        =   4
         Top             =   90
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   149094403
         CurrentDate     =   39198
      End
      Begin VB.Label lblLisSort 
         AutoSize        =   -1  'True
         Caption         =   "排序"
         Height          =   180
         Left            =   2850
         TabIndex        =   37
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "病人病区"
         Height          =   180
         Left            =   2850
         TabIndex        =   27
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "病 人 来 源"
         Height          =   180
         Left            =   150
         TabIndex        =   26
         Top             =   795
         Width           =   990
      End
      Begin VB.Label LabModify 
         AutoSize        =   -1  'True
         Caption         =   "(确定修改标本号的开始号码)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   2475
         TabIndex        =   24
         Top             =   2430
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label LabModify 
         AutoSize        =   -1  'True
         Caption         =   "修改标本号"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   2445
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "标本序号"
         Height          =   180
         Left            =   2265
         TabIndex        =   18
         Top             =   450
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "检验仪器(多选)"
         Height          =   180
         Left            =   90
         TabIndex        =   16
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "执行科室"
         Height          =   180
         Left            =   2850
         TabIndex        =   14
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "申 请 科 室"
         Height          =   180
         Left            =   150
         TabIndex        =   12
         Top             =   1110
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "检  验   人"
         Height          =   180
         Left            =   150
         TabIndex        =   10
         Top             =   1455
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "批       号"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   450
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   3090
         TabIndex        =   3
         Top             =   135
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "检验时间"
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   135
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   7485
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBatchAction.frx":68F3
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14499
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ImgList 
      Left            =   135
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":7185
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":71F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":778B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":7D25
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":82BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":EB21
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":15383
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":1BBE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":22447
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":28CA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   5820
      Top             =   990
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmBatchAction.frx":2F50B
      Left            =   5670
      Top             =   1680
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBatchAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngLeftWidth As Long           '左边的宽度
Private Const Dkp_ID_Left As Integer = 101
Private Const Dkp_ID_Right As Integer = 102
Private mintEditType As Integer         '操作类型 (1=打印 2=审核 3=无主删除)
Private mlngMachine As Long             '仪器ID
Private mstrPrivs As String             '权限
Private mstrAuditingMan As String       '审核人
Private mstrAuditingManID As String     '审核人ID
Private mintAuditing As Integer         '时间限制
Private mDateAuditing As Date           '审核开始时间
Private mDeptID As Long                 '执行科室ID
Private mintUnion As Integer            '是否区分仪器进行显示 0=不区分 1=区分
Private mMakeNoRule As String           '标本序号生成的日期规则
Private mstrPrintDepts As String        '可以打印的科室
Private mblnExec As Boolean              '是否正在执行

Private Enum mMCol          '仪器
    ID
    选择
    编码
    仪器
    类型
End Enum

Private Enum mCol           '列表
    选择 = 0
    急诊
    执行状态
    标本号
    标本类型
    核收时间
    核收人
    检验人
    申请时间
    申请人
    申请科室
    检验仪器
    执行科室
    医嘱id
    发送号
    转出
    标本id
    病人ID
    是否审核
    病人姓名
    相关ID
    标本序号
    仪器id
    病人来源
    婴儿
    开嘱科室ID
    报告结果
    主页ID
    采样人
    采样时间
End Enum

Private mclsUnzip As New cUnzip
Private mclsZip As New cZip


Private Sub cbo来源_Click()
    If cbo来源.ListIndex = 2 Or cbo来源.ListIndex = 0 Then
        cbo病区.Enabled = True
    Else
        cbo病区.ListIndex = 0
        cbo病区.Enabled = False
    End If
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
        Case conMenu_File_PrintSet                                                              '打印设置
            Call zlPrintSet
        Case conMenu_File_Exit                                                                  '退出
            Unload Me
        '---------------------------------------------------------------
        Case conMenu_Manage_ThingModi                                                           '全选
            Call RptSelect(Me.rptList.Records, True)
            Me.rptList.Populate
        Case conMenu_Manage_ThingDel                                                            '全清
            Call RptSelect(Me.rptList.Records, False)
            Me.rptList.Populate
        Case conMenu_File_Print                                                                 '报告打印
            Call SaveData
        Case conMenu_Edit_Audit                                                                 '审核
            Call SaveData
        Case conMenu_Edit_Delete                                                                '删除
            Call SaveData
        Case conMenu_Manage_Reset                                                               '批量修改标本号
            Call ModifySampleNumber
        
        '---------------------------------------------------------------
        Case conMenu_View_ToolBar                                                               '工具栏
        Case conMenu_View_ToolBar_Button                                                        '标准按钮
            Me.cbrthis(2).Visible = Not Me.cbrthis(2).Visible
            Me.cbrthis.RecalcLayout
        Case conMenu_View_ToolBar_Text                                                          '文本标签
            Dim cbrControl As CommandBarControl
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        Case conMenu_View_ToolBar_Size                                                          '大图标
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
        Case conMenu_View_StatusBar                                                             '状态栏
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbrthis.RecalcLayout
        Case conMenu_View_Refresh                                                               '刷新
            Call RefreshData
        '---------------------------------------------------------------
        Case conMenu_Help_Help                                                                  '帮助主题
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web                                                                   'WEB上的中联
            Call zlHomePage(hWnd)
        Case conMenu_Help_Web_Home                                                              '主页
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail                                                              '发送返馈
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About                                                                 '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
        
    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button:                                                                   '按钮
            Control.Checked = Me.cbrthis(2).Visible
        
        Case conMenu_View_ToolBar_Text:                                                                     '按钮文字
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        
        Case conMenu_View_ToolBar_Size:                                                                     '大图标
            Control.Checked = Me.cbrthis.Options.LargeIcons
        
        Case conMenu_View_StatusBar:                                                                        '状态栏
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkfilter_Click(Index As Integer)
    Dim Record  As ReportRecord
    For Each Record In Me.rptList.Records
        If Record.Item(mCol.急诊).Icon = 2 Then
            Record.Item(mCol.选择).Checked = (chkfilter(0).Value = 1)
        Else
            Record.Item(mCol.选择).Checked = (chkfilter(1).Value = 1)
        End If
    Next
    Me.rptList.Populate
End Sub

Private Sub dkpMain_Resize()
    Me.cbrthis.RecalcLayout
End Sub

Private Sub DtpEnd_Validate(Cancel As Boolean)
    '10765
    If DtpEnd.Value < DtpBegin.Value Then
        MsgBox "结束日期不能小于开始日期！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub DTPBegin_Validate(Cancel As Boolean)
    '10765
    If DtpBegin.Value > DtpEnd.Value Then
        MsgBox "开始日期不能大于结束日期！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim intSort As Integer
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
'    Me.cbrthis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "操作(&A)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "全选(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "全清(&R)")
        If mintEditType = 1 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 2 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审核(&A)"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 3 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 4 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "修改(&M)"): cbrControl.BeginGroup = True
        End If
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&F)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    '快键绑定
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_Select
        .Add FCONTROL, Asc("Z"), conMenu_Edit_DeSelect
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Audit
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With

    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "全选"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "全清")
        If mintEditType = 1 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 2 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审核"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 3 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 4 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "修改"): cbrControl.BeginGroup = True
        End If
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '分隔窗体
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane
    
    dkpMain.Options.HideClient = True
    mlngLeftWidth = Me.picWhere.Width - 250
    
    Set Pane1 = dkpMain.CreatePane(Dkp_ID_Left, 200, 150, DockLeftOf, Nothing)
    Pane1.Title = "条件选择"
    Pane1.Handle = Me.picWhere.hWnd
    Pane1.Options = PaneNoCaption

    Set Pane2 = dkpMain.CreatePane(Dkp_ID_Right, 200, 600, DockRightOf, Nothing)
    Pane2.Title = "列表"
    Pane2.Handle = Me.picList.hWnd
    Pane2.Options = PaneNoCaption
    
    Pane1.Select
    
    '初始化
    Me.DtpBegin = Now: Me.DtpEnd = Now
    '时间
    With Me.cboDate
        .Clear
        .AddItem "检验时间"
        .AddItem "报告时间"
        .ListIndex = 0
    End With
    
    
    Dim rsTmp As New ADODB.Recordset
    '病人病区
    With Me.cbo病区
        .Clear
        .AddItem "所有病区"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = " select Distinct B.编码,B.名称,A.病区id from 病区科室对应 A,部门表 B where A.病区id=B.id Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        Me.cbo病区.AddItem "" & rsTmp("名称")
        Me.cbo病区.ItemData(Me.cbo病区.NewIndex) = rsTmp("病区ID")
        rsTmp.MoveNext
    Loop
    If Me.cbo病区.ListCount > 0 Then Me.cbo病区.ListIndex = 0
    '病人来源
    With Me.cbo来源
        .Clear
        .AddItem "所有病人"
        .AddItem "门诊"
        .AddItem "住院"
        .AddItem "其他"
        .AddItem "体检"
    End With
    If Me.cbo来源.ListCount > 0 Then Me.cbo来源.ListIndex = 0
    '申请科室
    With Me.cboRequisitionDept
        .Clear
        .AddItem "所有科室"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = "SELECT A.名称 as 名称,ID FROM 部门表 A,部门性质说明 B " & _
               " WHERE A.ID=B.部门id AND B.工作性质='临床' ORDER BY A.编码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        With Me.cboRequisitionDept
            .AddItem rsTmp("名称")
            .ItemData(.NewIndex) = rsTmp("Id")
        End With
        rsTmp.MoveNext
    Loop
    If Me.cboRequisitionDept.ListCount > 0 Then Me.cboRequisitionDept.ListIndex = 0
    
    '执行科室
    With Me.cboExeDept
        .Clear
        .AddItem "所有科室"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = " SELECT A.名称 as 名称,ID FROM 部门表 A,部门性质说明 B " & _
              " WHERE A.ID=B.部门id AND b.工作性质 = '检验'  ORDER BY A.编码  "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        With Me.cboExeDept
            .AddItem rsTmp("名称")
            .ItemData(.NewIndex) = rsTmp("Id")
        End With
        rsTmp.MoveNext
    Loop
    If Me.cboExeDept.ListCount > 0 Then Me.cboExeDept.ListIndex = 0
        
    With Me.cboExeDept
        Dim lngIndex As Long
        If mDeptID > 0 And .ListCount > 0 Then
            For lngIndex = 0 To .ListCount - 1
                If mDeptID = .ItemData(lngIndex) Then
                    .ListIndex = lngIndex
                    Exit For
                End If
            Next
        End If
    End With
    '检验人
    With Me.cboVerifyMan
        .Clear
        .AddItem "所有人员"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = "Select Distinct 编号,姓名 As 名称, a.Id" & vbNewLine & _
            " From 人员表 a, 部门人员 b, 部门性质说明 c" & vbNewLine & _
            " Where a.Id = b.人员id And b.部门id = c.部门id And c.工作性质 = '检验'" & vbNewLine & _
            " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & vbNewLine & _
            " Order By 编号 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        With Me.cboVerifyMan
            .AddItem rsTmp("名称")
            .ItemData(.NewIndex) = rsTmp("Id")
        End With
        rsTmp.MoveNext
    Loop
    If Me.cboVerifyMan.ListCount > 0 Then Me.cboVerifyMan.ListIndex = 0
    
    '初使化列表列
    Dim Column As ReportColumn
    Dim intLoop As Integer
    Dim Record As ReportRecord
    
    With Me.rptMachine.Columns
        
        rptMachine.AllowColumnRemove = False
        rptMachine.ShowItemsInGroups = False
        
        With rptMachine.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptMachine.SetImageList ImgList
        
        Set Column = .Add(mMCol.ID, "ID", 18, False): Column.Visible = False
        Set Column = .Add(mMCol.选择, "选择", 18, False): Column.Icon = 0
        Set Column = .Add(mMCol.编码, "编码", 65, True)
        Set Column = .Add(mMCol.仪器, "仪器", 120, True)
        Set Column = .Add(mMCol.类型, "类型", 85, True)
        Me.rptMachine.Populate
    End With
    
    gstrSql = "select ID, 编码,名称,仪器类型 from 检验仪器 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    With Me.rptMachine
        .Records.DeleteAll
        .Populate
        Set Record = .Records.Add
        For intLoop = 0 To .Columns.Count
            Record.AddItem ""
        Next
        Record.Item(mMCol.ID).Value = 0
        Record.Item(mMCol.选择).HasCheckbox = True: Record.Item(mMCol.选择).Checked = False
        Record.Item(mMCol.仪器).Value = "手工"
    End With
    
    Do Until rsTmp.EOF
        With Me.rptMachine
            Set Record = .Records.Add
            For intLoop = 0 To .Columns.Count
                Record.AddItem ""
            Next
            
            Record.Item(mMCol.ID).Value = Nvl(rsTmp("ID"))
            Record.Item(mMCol.选择).HasCheckbox = True
            If mlngMachine = Nvl(rsTmp("ID")) Then
                Record.Item(mMCol.选择).Checked = True
            Else
                Record.Item(mMCol.选择).Checked = False
            End If
            Record.Item(mMCol.编码).Value = Nvl(rsTmp("编码"))
            Record.Item(mMCol.仪器).Value = Nvl(rsTmp("名称"))
            Record.Item(mMCol.类型).Value = Nvl(rsTmp("仪器类型"))
                        
        End With
        rsTmp.MoveNext
    Loop
    Me.rptMachine.Populate
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList.SetImageList ImgList
        
        
        Set Column = .Add(mCol.选择, "选择", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.急诊, "急诊", 18, False): Column.Icon = 1
        Set Column = .Add(mCol.执行状态, "状态", 18, False): Column.Icon = 4
        
        Set Column = .Add(mCol.标本号, "标本号", 80, True)
        Set Column = .Add(mCol.病人姓名, "病人姓名", 65, True)
        Set Column = .Add(mCol.标本类型, "标本类型", 65, True)
        Set Column = .Add(mCol.核收时间, "核收时间", 80, True)
        Set Column = .Add(mCol.核收人, "核收人", 65, True)
        Set Column = .Add(mCol.检验人, "检验人", 65, True)
        Set Column = .Add(mCol.申请时间, "申请时间", 65, True)
        Set Column = .Add(mCol.申请人, "申请人", 65, True)
        Set Column = .Add(mCol.申请科室, "申请科室", 65, True)
        Set Column = .Add(mCol.检验仪器, "检验仪器", 65, True)
        Set Column = .Add(mCol.执行科室, "执行科室", 65, True)
        Set Column = .Add(mCol.医嘱id, "医嘱id", 65, True): Column.Visible = False
        Set Column = .Add(mCol.发送号, "发送号", 65, True): Column.Visible = False
        Set Column = .Add(mCol.转出, "转出", 65, True): Column.Visible = False
        Set Column = .Add(mCol.标本id, "标本ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.病人ID, "病人ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.是否审核, "是否审核", 65, True): Column.Visible = False
        Set Column = .Add(mCol.相关ID, "相关Id", 65, True): Column.Visible = False
        Set Column = .Add(mCol.标本序号, "标本序号", 65, True): Column.Visible = False
        Set Column = .Add(mCol.仪器id, "仪器ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.婴儿, "婴儿", 65, True): Column.Visible = False
        Set Column = .Add(mCol.开嘱科室ID, "开嘱科室ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.报告结果, "报告结果", 65, True): Column.Visible = False
        Set Column = .Add(mCol.主页ID, "主页ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.采样人, "采样人", 65, True): Column.Visible = False
        Set Column = .Add(mCol.采样时间, "采样时间", 80, True): Column.Visible = False
        
        Me.rptMachine.Populate
    End With
   
    Select Case mintEditType
        Case 1
            Me.Caption = "批量打印"
        Case 2
            Me.Caption = "批量审核"
        Case 3
            Me.Caption = "批量删除无主"
        Case 4
            Me.Caption = "批量修改标本号"
    End Select
    If mintEditType = 4 Then
        Me.Label8.Top = Me.LabModify(0).Top + Me.LabModify(0).Height + 100
        chkPrint.Top = cboVerifyMan.Top
        chkPrint.Left = Label7.Left
        
        For intLoop = 0 To 2
            optSort(intLoop).Visible = False
        Next
    Else
        Me.Label8.Top = Frame1.Top + Frame1.Height + 100
    End If
    Me.rptMachine.Top = Me.Label8.Top + Me.Label8.Height + 100
    
    Me.chkUnion.Value = zlDatabase.GetPara("frmBatchAction_不打印被合并标本", 100, 1208, 1)
    Me.chkUnion.Value = zlDatabase.GetPara("frmBatchAction_同一个病人合并为一个报告单打印", 100, 1208, 0)
    mintUnion = zlDatabase.GetPara("不区分仪器显示核收项目", 100, 1208, 0)
    mMakeNoRule = zlDatabase.GetPara("标本序号生成规则", 100, 1208, "今  天")
    mstrPrintDepts = zlDatabase.GetPara("只打指定科室报告单", 100, 1208, "")
    cboDate.ListIndex = Val(zlDatabase.GetPara("批量打印时间条件", 100, 1208, 0))
    
    intSort = zlDatabase.GetPara("批量打印按病人排序", 100, 1208, 0)
    If intSort >= 0 And intSort <= 2 Then
        optSort(intSort).Value = True
    Else
        optSort(0).Value = True
    End If
    
    Me.chkAbnormal.Value = zlDatabase.GetPara("frmBatchAction_批审异常结果标本显示", 100, 1208, 0)
    
    Call RestoreWinState(Me, App.ProductName)                   '界面恢复
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane
    On Error Resume Next

    If Me.WindowState = 1 Then Exit Sub

    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Left)
    Pane1.MinTrackSize.SetSize mlngLeftWidth / Screen.TwipsPerPixelX, Pane1.MaxTrackSize.Height
    Pane1.MaxTrackSize.SetSize mlngLeftWidth / Screen.TwipsPerPixelX, Pane1.MaxTrackSize.Height
    
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTmp As String
    Dim intLoop As Integer
    If mblnExec = True Then
        If mintEditType = 1 Then
            strTmp = "打印"
        ElseIf mintEditType = 2 Then
            strTmp = "审核"
        Else
            strTmp = "无主删除"
        End If
        MsgBox "正在执行<" & strTmp & ">操作不能退出!", vbInformation, Me.Caption
        Cancel = True
        Exit Sub
    End If
        
    zlDatabase.SetPara "frmBatchAction_不打印被合并标本", Me.chkUnion.Value, 100, 1208
    zlDatabase.SetPara "frmBatchAction_同一个病人合并为一个报告单打印", Me.chkPatient.Value, 100, 1208
    zlDatabase.SetPara "批量打印时间条件", Me.cboDate.ListIndex, 100, 1208

    For intLoop = 0 To 2
        If Me.optSort(intLoop).Value = True Then
            Exit For
        End If
    Next
    zlDatabase.SetPara "批量打印按病人排序", intLoop, 100, 1208
    zlDatabase.SetPara "frmBatchAction_批审异常结果标本显示", Me.chkAbnormal.Value, 100, 1208
    frmLabMain.zlRefreshData
End Sub

Private Sub Option1_Click()

End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = 20
        .Width = Me.picList.ScaleWidth - 40
        .Height = Me.picList.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub picWhere_Resize()
    Me.rptMachine.Height = Me.picWhere.ScaleHeight - Me.rptMachine.Top - 100
End Sub
Private Sub RefreshData()
    '功能                   '刷新数据
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer, lngloop As Long
    Dim Record As ReportRecord
    Dim varItem As Variant
    Dim varBetween As Variant
    Dim blnMoved As Boolean
    Dim strSQLbak As String
    Dim strTmp As String
    Dim str病区 As String, str病区Where As String, str病区显示 As String, lng病人来源 As Long
    Dim strSQL As String, i As Integer, rsTmp1 As ADODB.Recordset
    Dim intCount As Integer
    Dim blnShow As Boolean
    
    On Error GoTo errH
    
    With Me.rptList
        .Records.DeleteAll
        .Populate
    End With
    '10765  检验开始时间和结束时间
    If DtpBegin.Value > DtpEnd.Value Then
        MsgBox "开始日期不能大于结束日期！", vbInformation, gstrSysName
        DtpBegin.SetFocus
        Exit Sub
    End If
    blnMoved = MovedByDate(Me.DtpBegin.Value)
    '病区
    If Me.cbo病区.ItemData(Me.cbo病区.ListIndex) > 0 Then
        str病区 = " ,(Select K.名称 As 病区,J.病区id,J.科室ID From 病区科室对应 J,部门表 K Where J.病区id=K.ID) I   "
        str病区Where = " And B.病人科室id=I.科室id And I.病区id= [8] "
        str病区显示 = ",病区ID    "
    Else
        If optSort(2).Value = True Then
            str病区 = " ,(Select K.名称 As 病区,J.病区id,J.科室ID From 病区科室对应 J,部门表 K Where J.病区id=K.ID) I   "
            str病区Where = " And B.病人科室id=I.科室id  "
            str病区显示 = ",病区ID    "
        Else
            str病区 = ""
            str病区Where = ""
        End If
    End If
    gstrSql = "select /*+ RULE */ DISTINCT B.相关ID AS ID,A.医嘱id,F.发送号,0 AS 选择," & _
                      " Decode(A.仪器id, Null, " & vbCrLf & _
                        " to_Char(Trunc(A.标本序号/10000)+1,'0000')|| '-'||to_Char(MOD(A.标本序号,10000),'0000'), A.标本序号) As 标本号, " & _
                      "A.标本类型," & _
                      "TO_CHAR(A.核收时间,'MM-DD HH24:MI') AS 核收时间," & _
                      "A.核收人," & _
                      "A.检验人," & _
                      "lpad(A.标本序号,8,'0') as 排序," & _
                      "TO_CHAR(B.开嘱时间,'MM-DD HH24:MI') AS 申请时间," & _
                      "B.开嘱医生 AS 申请人," & _
                      "C.名称 AS 申请科室," & _
                      "E.名称 AS 执行科室," & _
                      "A.id as 标本ID, a.采样人,a.采样时间, " & _
                      "B.病人id, " & _
                      "D.名称 AS 检验仪器,0 As 转出,Decode(A.标本类别,1,'√','') As 急诊, " & _
                      "decode(a.审核时间,Null,'否','是') as 是否审核, " & _
                      "Decode(a.样本状态, 1, '检验中', 2, '已检验') As 执行状态, " & _
                      "Decode(a.是否传送, 1, '', '传送失败') As 传送, a.打印次数,a.微生物标本, " & _
                      "a.姓名,a.标本序号,a.仪器ID,a.病人来源,a.婴儿,b.开嘱科室ID,a.报告结果,b.主页ID  " & str病区显示 & _
                 "from 检验标本记录 A, 病人医嘱记录 B, 部门表 C, 检验仪器 D,部门表 E,病人医嘱发送 F,病人信息 G, " & _
                 " (Select * From Table(Cast(f_str2list([6]) As zltools.t_strlist))) H " & _
                  str病区 & _
                 " WHERE A.医嘱ID = B.相关ID(+) AND B.开嘱科室ID = C.ID(+) AND B.ID=F.医嘱id(+) AND " & _
                      "A.仪器ID = D.ID(+) AND B.执行科室id = E.ID(+) AND A.样本状态 IN (1,2) AND a.病人ID = G.病人ID(+)  " & _
                      "  " & str病区Where
                      
    '处理使用核收还是报告时间
    If cboDate.Text = "检验时间" Then
        gstrSql = gstrSql & " and 核收时间 between [1] and [2] "
    Else
        gstrSql = gstrSql & " and 检验时间 between [1] and [2] "
    End If
                      
    Select Case mintEditType
        Case 1
            '-------- 晋煤增加
            If Me.chkPrint.Value <> 1 Then
                gstrSql = gstrSql & " And nvl(a.打印次数,0) = 0 "
            End If
            '-------- 晋煤增加

            gstrSql = gstrSql & " and a.样本状态 in (1,2)  and a.姓名 is not null " & _
                                IIf(Me.chkUnion.Value = 1, " and nvl(a.合并ID,0) = 0 ", "")
            If InStr(mstrPrivs, "未审核打印") <= 0 Then
                gstrSql = gstrSql & " And 样本状态 = 2 "
            End If
        Case 2
            gstrSql = gstrSql & " and a.样本状态 = 1  and a.姓名  is not null"
        Case 3
            gstrSql = gstrSql & " and a.姓名 is null  and nvl(是否质控品,0) = 0  "
    End Select
    
    '申请科室
    If Me.cboRequisitionDept.ItemData(Me.cboRequisitionDept.ListIndex) > 0 Then
        gstrSql = gstrSql & " And b.开嘱科室ID = [3] "
    End If
    
    '执行科室
    If Me.cboExeDept.ItemData(Me.cboExeDept.ListIndex) > 0 Then
        gstrSql = gstrSql & " and a.执行科室Id = [4] "
    End If
    
    '检验人
    If Me.cboVerifyMan.ItemData(Me.cboVerifyMan.ListIndex) > 0 Then
        gstrSql = gstrSql & " and a.检验人 = [5] "
        
    End If
    
    '病人来源
    If Me.cbo来源.ListIndex > 0 Then
        gstrSql = gstrSql & " And a.病人来源=[9] "
        
        Select Case Me.cbo来源.List(Me.cbo来源.ListIndex)
        Case "门诊": lng病人来源 = 1
        Case "住院": lng病人来源 = 2
        Case "其他": lng病人来源 = 3
        Case "体检": lng病人来源 = 4
        End Select
    End If
    
    '只显示紧急医嘱
    
    
    '处理标本号
    If Trim(TxtSample.Text) <> "" Then
        TxtSample.Text = Replace(Replace(TxtSample.Text, "～", "~"), "-", "~")
        If Check_Sample = False Then Exit Sub '10861
        varItem = Split(Trim(TxtSample.Text), ",")
        For lngloop = 0 To UBound(varItem)
            varBetween = Split(varItem(lngloop), "~")
            If UBound(varBetween) > 0 Then
                strTmp = strTmp & "  OR lpad(A.标本序号,8,'0') BETWEEN lpad(" & IIf(Trim(Me.txtBatchNum) <> "", TransSampleNO(Val(Me.txtBatchNum) & "-" & Val(varBetween(0))), Val(varBetween(0))) & _
                        ",8,'0') AND lpad(" & IIf(Trim(Me.txtBatchNum) <> "", TransSampleNO(Val(Me.txtBatchNum) & "-" & Val(varBetween(1))), Val(varBetween(1))) & ",8,'0')"
            Else
                strTmp = strTmp & " OR A.标本序号='" & IIf(Trim(Me.txtBatchNum) <> "", TransSampleNO(Val(Me.txtBatchNum) & "-" & Val(varItem(lngloop))), Val(varItem(lngloop))) & "'"
            End If
        Next
            
    Else
        '只有批号时选择所有
        If Trim(Me.txtBatchNum) <> "" Then
            strTmp = strTmp & " or a.标本序号 between " & TransSampleNO(Val(Me.txtBatchNum) & "-0001") & " And " & TransSampleNO(Val(Me.txtBatchNum) & "-9999")
        End If
    End If
                              
    If strTmp <> "" Then gstrSql = gstrSql & " AND (1=2 " & strTmp & ")"
    strTmp = ""
    
    With Me.rptMachine
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mMCol.选择).Checked = True Then
                strTmp = strTmp & "," & .Records(intLoop).Item(mMCol.ID).Value
            End If
        Next
    End With
    
    If strTmp = "" Then
        MsgBox "请选择一个设备！", vbInformation
        Exit Sub
    Else
        strTmp = Mid(strTmp, 2)
    End If
    
    gstrSql = gstrSql & " And nvl(a.仪器ID,0) = h.Column_Value  "
    
    
    If InStr(mstrPrivs, "审核限制") > 0 And mintEditType = 2 Then
        '--- 20007-08-30 10783 批量审核时，审核人不能和检验人相同
        gstrSql = gstrSql & " And A.检验人 <> [7] "
    End If
    
    zlCommFun.ShowFlash "正在刷新数据请稍候..."
    Me.MousePointer = 11
    
    
'    If blnMoved Then
'        strSQLBak = gstrSql
'        strSQLBak = Replace(strSQLBak, "0 As 转出", "1 As 转出")
'        strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
'        strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
'        strSQLBak = Replace(strSQLBak, "检验标本记录", "H检验标本记录")
'        gstrSql = gstrSql & " Union ALL " & strSQLBak
'    End If
    '处理排序
    
    If optSort(0).Value = True Then
        gstrSql = gstrSql & "  Order by " & " 排序 "
    ElseIf optSort(1).Value = True Then
        gstrSql = gstrSql & "  Order by " & "病人ID,排序 "
    ElseIf optSort(2).Value = True Then
        gstrSql = gstrSql & "  Order by " & "病区ID,排序 "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(Me.DtpBegin, "yyyy-mm-dd 00:00:00")), _
                                         CDate(Format(Me.DtpEnd, "yyyy-mm-dd 23:23:59")), _
                                         CLng(Me.cboRequisitionDept.ItemData(Me.cboRequisitionDept.ListIndex)), _
                                         CLng(Me.cboExeDept.ItemData(Me.cboExeDept.ListIndex)), _
                                         CStr(Me.cboVerifyMan.Text), strTmp, IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan), _
                                         CLng(Me.cbo病区.ItemData(Me.cbo病区.ListIndex)), lng病人来源)
                                         
    With Me.rptList
        .Records.DeleteAll
        .Populate
    End With
    
    Do Until rsTmp.EOF
        With Me.rptList
            If chkAbnormal.Visible = True And chkAbnormal.Value = 1 Then
                blnShow = Not CheckAbnormal(rsTmp("标本ID"))
            Else
                 blnShow = True
            End If
            
            If blnShow = True Then
        
                Set Record = .Records.Add
                For intLoop = 0 To .Columns.Count
                    Record.AddItem ""
                Next
                
                Record.Item(mCol.选择).HasCheckbox = True
                Record.Item(mCol.相关ID).Value = Nvl(rsTmp("ID"))
                Record.Item(mCol.标本id).Value = Nvl(rsTmp("标本ID"))
                Record.Item(mCol.标本号).Value = Val(Nvl(rsTmp("标本序号")))
                Record.Item(mCol.标本号).Caption = Trim(Nvl(rsTmp("标本号")))
                '-----------------------------------------------------晋煤加
                If CInt(Nvl(rsTmp("打印次数"), "0")) > 0 Then
                    Record.Item(mCol.执行状态).Value = "已打印"
                    Record.Item(mCol.执行状态).Icon = 7
                ElseIf Nvl(rsTmp("执行状态")) = "已检验" Then
                    Record.Item(mCol.执行状态).Value = "已检验"
                    Record.Item(mCol.执行状态).Icon = 6
                ElseIf Nvl(rsTmp("传送")) = "" Then
                    Record.Item(mCol.执行状态).Value = "已传送"
                    Record.Item(mCol.执行状态).Icon = 5
                End If
                
                If Val("" & rsTmp!微生物标本) = 0 Then
                    strSQL = "Select Count(A.ID) - Sum(Decode(A.检验结果, Null, 0, 1)) As 无结果记录,Count(A.ID) as 结果数 " & vbNewLine & _
                            "From 检验普通结果 A" & vbNewLine & _
                            "Where A.检验标本id = [1]"
                    Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & rsTmp("标本ID")))
                    If rsTmp1.EOF Then
                        For i = 0 To rptList.Columns.Count - 1
                            Record.Item(i).BackColor = vbWhite
                        Next
                    Else
                        If Val("" & rsTmp1.Fields("无结果记录")) = 0 And Val("" & rsTmp1.Fields("结果数")) > 0 Then
                            For i = 0 To rptList.Columns.Count - 1
                                Record.Item(i).BackColor = &HFDD6C6
                            Next
                        Else
                            For i = 0 To rptList.Columns.Count - 1
                                Record.Item(i).BackColor = vbWhite
                            Next
    
                        End If
                    End If
                Else
                    For i = 0 To rptList.Columns.Count - 1
                        Record.Item(i).BackColor = vbWhite
                    Next
                End If
                
                '-----------------------------------------------------
                
                Record.Item(mCol.标本类型).Value = Nvl(rsTmp("标本类型"))
                Record.Item(mCol.病人ID).Value = Nvl(rsTmp("病人ID"))
                Record.Item(mCol.病人姓名).Value = Nvl(rsTmp("姓名"))
                Record.Item(mCol.发送号).Value = Nvl(rsTmp("发送号"))
                Record.Item(mCol.核收人).Value = Nvl(rsTmp("核收人"))
                Record.Item(mCol.核收时间).Value = Nvl(rsTmp("核收时间"))
                Record.Item(mCol.急诊).Icon = IIf(Nvl(rsTmp("急诊")) = "√", 2, -1)
                Record.Item(mCol.检验人).Value = Nvl(rsTmp("检验人"))
                Record.Item(mCol.检验仪器).Value = Nvl(rsTmp("检验仪器"))
                Record.Item(mCol.申请科室).Value = Nvl(rsTmp("申请科室"))
                Record.Item(mCol.申请人).Value = Nvl(rsTmp("申请人"))
                Record.Item(mCol.申请时间).Value = Nvl(rsTmp("申请时间"))
                Record.Item(mCol.是否审核).Value = Nvl(rsTmp("是否审核"))
                Record.Item(mCol.医嘱id).Value = Nvl(rsTmp("医嘱ID"))
                Record.Item(mCol.执行科室).Value = Nvl(rsTmp("执行科室"))
                Record.Item(mCol.转出).Value = Nvl(rsTmp("转出"))
                Record.Item(mCol.标本序号).Value = Nvl(rsTmp("标本序号"))
                Record.Item(mCol.仪器id).Value = Nvl(rsTmp("仪器ID"))
                Record.Item(mCol.病人来源).Value = Nvl(rsTmp("病人来源"), 3)
                Record.Item(mCol.婴儿).Value = Nvl(rsTmp("婴儿"), 0)
                Record.Item(mCol.开嘱科室ID).Value = Nvl(rsTmp("开嘱科室ID"))
                Record.Item(mCol.报告结果).Value = Nvl(rsTmp("报告结果"))
                Record.Item(mCol.主页ID).Value = Nvl(rsTmp("主页ID"))
                Record.Item(mCol.采样人).Value = Nvl(rsTmp("采样人"))
                Record.Item(mCol.采样时间).Value = Nvl(rsTmp("采样时间"))
            Else
                intCount = intCount + 1
            End If
        End With
        rsTmp.MoveNext
    Loop
    Me.rptList.Populate
    zlCommFun.StopFlash
    stbThis.Panels(2).Text = "当前已找到" & rsTmp.RecordCount & "条记录。"
    Call chkfilter_Click(0)
    Me.MousePointer = 0
    Exit Sub
errH:
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub



Private Function CheckAbnormal(lngSample As Long) As Boolean
    '功能             检验是否标本有异常的（警示上限和下限 结果标志=5或6）
    '返回             有异常时返回为真
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select count(id) id from 检验普通结果 where 检验标本id = [1]  and 结果标志 in (5,6) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查标本是否有异常数据", lngSample)
    If rsTmp.RecordCount > 0 Then
        If Val(rsTmp("id") & "") > 0 Then
            CheckAbnormal = True
            Exit Function
        End If
    End If
    
End Function
Private Sub RptSelect(Records As ReportRecords, blTrue As Boolean)
    '功能                           选择或取消选择
    '参数                           Records = 列表对象
    '                               blTrue  True = 选择 False = 取消选择
    Dim intLoop As Integer
    Me.chkfilter(0).Value = IIf(blTrue, 1, 0)
    Me.chkfilter(1).Value = IIf(blTrue, 1, 0)
    For intLoop = 0 To Records.Count - 1
        Records(intLoop).Item(mCol.选择).Checked = blTrue
    Next
End Sub


Private Sub SaveData(Optional blnPrintNoAuditing As Boolean = False)
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim blPrint As Boolean
    Dim blnAutoPrint As Boolean
    Dim lngloop As Long
    Dim bln审核通过  As Boolean
    Dim strMsg As String '显示未审核通过的提示信息
    Dim strErrInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim intLoop As Integer
    Dim strChart(1 To 9) As String
    Dim intPrivacy As Integer                               '发送报告单到医生站时是否显示隐私项目
    Dim blnCheckExesState As Boolean                        '出院病人存在记帐划价费用
    Dim lngAdvice As Long                                   '医嘱ID
    Dim intPrintCount As Integer                            '打印计数
    Dim intOnPrint As Integer                               '还有几个未审核的标本未打印
    Dim lngPatient As Long                                  '记录病人ID
    Dim str医嘱ID As String                                 '医嘱ID用","分隔
    Dim str标本ID As String                                 '标本ID用","分隔
    Dim intItem As Integer                                  '临时记录
    Dim astrItem() As String                                '数据用于记录ID
    Dim strPrintCode As String                              '打印单据编码
    Dim lng医嘱ID As Long                                   '医嘱ID
    Dim lng标本ID As Long                                   '标本ID
    Dim lng病人ID As Long                                   '病人ID
    Dim lng申请科室ID As Long                               '申请科室Id
    Dim blnRollBack As Boolean                              '是否回滚
    Dim astrSQL() As String                                 '要执行的数组
    Dim strTmp() As String
    Dim blng强制审核 As Boolean                             '强制审核通过
    Dim strDate As String                                   '采集时间不能通过
    
    On Error GoTo ErrHand
    
    ReDim astrSQL(0)
    Me.MousePointer = 11
    mblnExec = True
    blnAutoPrint = zlDatabase.GetPara("审核打印", 100, 1208, 0)
    '填写到新的电子病历里
    intPrivacy = zlDatabase.GetPara("报告单是否显示隐私项目", 100, 1208, 0)
    intPrintCount = 0
    intOnPrint = 0
    blPrint = blnPrintNoAuditing
    
    If Me.chkPatient.Value = 1 Then
        Me.rptList.SortOrder.DeleteAll
        Me.rptList.SortOrder.Add Me.rptList.Columns(mCol.病人ID)
        Me.rptList.Populate
    End If
    
    With Me.rptList
        For lngloop = 0 To .Records.Count - 1
            
            If .Records(lngloop).Item(mCol.选择).Checked = True And Val(.Records(lngloop).Item(mCol.相关ID).Value) > 0 Then
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                lng申请科室ID = Val(.Records(lngloop).Item(mCol.开嘱科室ID).Value)
                If mintEditType = 1 And InStr("," & mstrPrintDepts & ",", lng申请科室ID) > 0 Then
                    If Me.chkPatient.Value = 0 Then
                        '==按当前标本进行打印
                        '生成图形供自定义报表调用
                        gstrSql = "select id from 检验图像结果 where 标本id = [1] order by id"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(.Records(lngloop).Item(mCol.标本id).Value))
                        
                        For intLoop = 1 To 9
                            strChart(intLoop) = ""
                        Next
                        intLoop = 1
                        Do Until rsTmp.EOF
                            If intLoop > 9 Then Exit For
                            strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                            Debug.Print strChart(intLoop)
                            Call LoadImageData(App.path, rsTmp("ID"))
                            
                            intLoop = intLoop + 1
                            rsTmp.MoveNext
                        Loop
                        If mintEditType = 1 Then '打印
                            zlCommFun.ShowFlash "正在打印数据,已完成(" & lngloop + 1 & "/" & .Records.Count & ")"
                            If GetReportCode(Val(.Records(lngloop).Item(mCol.医嘱id).Value), Val(.Records(lngloop).Item(mCol.发送号).Value), strReportCode, strReportParaNo, bytReportParaMode, _
                                Val(.Records(lngloop).Item(mCol.转出).Value) = 1) Then
                                
                                If .Records(lngloop).Item(mCol.是否审核).Value = "否" And blPrint = False Then
                                    intOnPrint = intOnPrint + 1
                                Else
                                    If .Records(lngloop).Item(mCol.是否审核).Value = "是" Or InStr(mstrPrivs, "未审核打印") > 0 Then
                                        If intPrintCount = 0 Then Call ReportTaskBegin
                                        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, _
                                            "医嘱ID=" & Val(.Records(lngloop).Item(mCol.医嘱id).Value), _
                                            "病人ID=" & Val(.Records(lngloop).Item(mCol.病人ID).Value), _
                                            "标本ID=" & Val(.Records(lngloop).Item(mCol.标本id).Value), _
                                            "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                                            "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                                            "图形9=" & strChart(9), 2)
                                            intPrintCount = intPrintCount + 1
                                        .Records(lngloop).Item(mCol.选择).Checked = False
                                        .Populate
                                    End If
                                End If
                                '对已审核的标本做打印标志
                                If .Records(lngloop).Item(mCol.是否审核).Value = "是" Then
                                    If mintUnion = 0 Then
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_标本质控(" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ",'',1)"
                                    Else
                                        gstrSql = "select ID from 检验标本记录 where 医嘱ID = [1] "
                                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.医嘱id).Value))
                                        Do Until rsTmp.EOF
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_标本质控(" & rsTmp("ID") & ",'',1)"
                                            rsTmp.MoveNext
                                        Loop
                                    End If
                                End If
                            End If
                        End If
                    Else
                        '==按病人进行打印
                        If lngPatient <> Val(.Records(lngloop).Item(mCol.病人ID).Value) Then
                            If str标本ID <> "" Then
                                intLoop = 1
                                For intLoop = 1 To 9
                                    strChart(intLoop) = ""
                                Next
                                str标本ID = Mid(str标本ID, 2)
                                str医嘱ID = Mid(str医嘱ID, 2)
                                lng医嘱ID = Split(str医嘱ID, ",")(0)
                                lng标本ID = Split(str标本ID, ",")(0)
                                If strPrintCode = "" Then
                                    '有多个格式时得到格式
                                    frmLabMainPrintFormat.ShowMe Me, str医嘱ID, strPrintCode
                                End If
                                astrItem = Split(str标本ID, ",")
                                For intItem = 0 To UBound(astrItem)
                                    gstrSql = "select id from 检验图像结果 where 标本id = [1] order by id"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(astrItem(intItem)))
                                    Do Until rsTmp.EOF
                                        If intLoop > 9 Then Exit For
                                        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                                        Call LoadImageData(App.path, rsTmp("ID"))
                                        intLoop = intLoop + 1
                                        rsTmp.MoveNext
                                    Loop
                                Next
                                If intPrintCount = 0 Then Call ReportTaskBegin
                                Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & str医嘱ID, _
                                        "病人ID=" & lngPatient, "标本ID=" & str标本ID, "多个医嘱=" & str医嘱ID, "多个标本=" & str标本ID, _
                                        "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                                        "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                                        "图形9=" & strChart(9), 2)
                                intPrintCount = intPrintCount + 1
                                For intItem = 0 To UBound(astrItem)
                                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                    astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_标本质控(" & astrItem(intItem) & ",'',1)"
                                Next
                                str标本ID = "": str医嘱ID = ""
                            End If
                        End If
                        str标本ID = str标本ID & "," & Val(.Records(lngloop).Item(mCol.标本id).Value)
                        str医嘱ID = str医嘱ID & "," & Val(.Records(lngloop).Item(mCol.医嘱id).Value)
                        lngPatient = Val(.Records(lngloop).Item(mCol.病人ID).Value)
                                    
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If mintEditType = 2 Then '审核
                    strDate = ""
                    If .Records(lngloop).Item(mCol.采样人).Value <> "" Then
                        If .Records(lngloop).Item(mCol.采样时间).Value <> "" Then
                            If CDate(.Records(lngloop).Item(mCol.采样时间).Value) > zlDatabase.Currentdate Then
                                strDate = "UN"
                                strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.标本号).Value & " " & .Records(lngloop).Item(mCol.病人姓名).Value & " 采样时间，大于当前时间，不能进行审核！"
                            End If
                        End If
                    End If
                    If strDate <> "UN" Then
                        bln审核通过 = False
                        zlCommFun.ShowFlash "正在审核数据,已完成(" & lngloop + 1 & "/" & .Records.Count & ")"
                        
                        '21137 已归档报告不能取消
                        gstrSql = "Select Decode(病案状态, 1, '1-等待审查', 2, '2-拒绝审查', 3, '3-正在审查', 4, '4-审查反馈', 5, '5-审查归档') As 病案状态" & vbNewLine & _
                        "From 检验标本记录 A, 病案主页 B ,病案提交记录 C" & vbNewLine & _
                        "Where A.病人id = B.病人id And A.主页id = B.主页id And A.病人来源 = 2 And Nvl(B.病案状态, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                        " And b.病人id = c.病人Id and B.主页id = C.主页ID "
                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.标本id).Value))
                        If rsTmp.EOF = False Then
                            strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.标本号).Value & " " & .Records(lngloop).Item(mCol.病人姓名).Value & " 本次住院的病案已提交审查，不能审核！"
                        Else
                            '-------------------------------------------------------------------------------------------
                            If VerifyAuditingRule(Val(.Records(lngloop).Item(mCol.标本id).Value), strErrInfo, 2) = 1 Then
                                strErrInfo = ""
                                
                                blng强制审核 = (InStr(mstrPrivs, "批量强制审核规则") > 0)
                                If blng强制审核 = True Then
                                    strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.标本号).Value & " " & .Records(lngloop).Item(mCol.病人姓名).Value & " 的检验结果使用强制审核权限通过审核！" & vbNewLine & strErrInfo
                                Else
                                    strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.标本号).Value & " " & .Records(lngloop).Item(mCol.病人姓名).Value & " 的检验结果未通审核！" & vbNewLine & strErrInfo
                                End If
                            Else
                                blng强制审核 = True
                            End If
                            If blng强制审核 = True Then
                                If CheckChargeState(Val(.Records(lngloop).Item(mCol.相关ID).Value), False) = False Then
                                    blnCheckExesState = CheckExesState(Val(.Records(lngloop).Item(mCol.标本id).Value))
                                    If mintUnion = 0 Then
                                        '未收费
                                        If InStr(mstrPrivs, "未收费审核") > 0 And blnCheckExesState = True Then
                                            '签名不成功时退出
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ";" & mstrAuditingManID
                                            
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_报告审核(" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ",'" & _
                                                                         IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "','" & _
                                                                         UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                                            
                                            bln审核通过 = True
                                            
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Zl_检验报告单_Update(" & Val(.Records(lngloop).Item(mCol.标本id).Value) & "," & intPrivacy & ",'" & gstrUnitName & "')"           '审核后处理病历报告单
                                            

                                            
                                        Else
                                            strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.标本号).Value & " " & .Records(lngloop).Item(mCol.病人姓名).Value & _
                                                        IIf(blnCheckExesState, " 未收费！", " 出院病人记帐划价费用不能审核")
                                        End If
                                    Else
                                        gstrSql = "select ID from 检验标本记录 where 医嘱ID = [1] "
                                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.医嘱id).Value))
                                        Do Until rsTmp.EOF
                                            '未收费
                                            If InStr(mstrPrivs, "未收费审核") > 0 And blnCheckExesState = True Then
                                                '签名不成功时退出
                                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                                astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ";" & mstrAuditingManID
                                                
                                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                                astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_报告审核(" & rsTmp("ID") & ",'" & _
                                                                             IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "','" & _
                                                                             UserInfo.编号 & "','" & UserInfo.姓名 & "')"

                                                bln审核通过 = True
                                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                                astrSQL(UBound(astrSQL)) = "Zl_检验报告单_Update(" & rsTmp("ID") & "," & intPrivacy & ",'" & gstrUnitName & "')"            '审核后处理病历报告单

                                                
                                            Else
                                                strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.标本号).Value & " " & .Records(lngloop).Item(mCol.病人姓名).Value & _
                                                            IIf(blnCheckExesState, " 未收费！", " 出院病人记帐划价费用不能审核")
                                            End If
                                            rsTmp.MoveNext
                                        Loop
                                    End If '未收费 End
                                Else  '检查收费状态
                                    If mintUnion = 0 Then
                                        '签名不成功时退出
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ";" & mstrAuditingManID
                                        
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_报告审核(" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ",'" & _
                                                                         IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "','" & _
                                                                         UserInfo.编号 & "','" & UserInfo.姓名 & "')"

                                        bln审核通过 = True
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "Zl_检验报告单_Update(" & Val(.Records(lngloop).Item(mCol.标本id).Value) & "," & intPrivacy & ",'" & gstrUnitName & "')"           '审核后处理病历报告单

                                        
                                    Else
                                        gstrSql = "select ID from 检验标本记录 where 医嘱ID = [1] "
                                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.医嘱id).Value))
                                        Do Until rsTmp.EOF
                                           '签名不成功时退出
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ";" & mstrAuditingManID
                                            
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_报告审核(" & rsTmp("ID") & ",'" & _
                                                                         IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan) & "','" & _
                                                                         UserInfo.编号 & "','" & UserInfo.姓名 & "')"
 
                                            bln审核通过 = True
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Zl_检验报告单_Update(" & rsTmp("ID") & "," & intPrivacy & ",'" & gstrUnitName & "')"         '审核后处理病历报告单

                                            
                                            rsTmp.MoveNext
                                        Loop
                                    End If  '是否区分仪器进行显示 End
                                End If '检查收费状态 End
                            End If '审核规则 End
                        End If
                        '-------------------------------------------------------------------------------------------
                    End If
                    If blnAutoPrint And bln审核通过 And InStr("," & mstrPrintDepts & ",", "," & lng申请科室ID & ",") > 0 Then

                        If GetReportCode(Val(.Records(lngloop).Item(mCol.医嘱id).Value), Val(.Records(lngloop).Item(mCol.发送号).Value), strReportCode, strReportParaNo, bytReportParaMode, _
                             False) Then
                            '生成图形供自定义报表调用
                            'frmLabMainImage.zlRefresh .Records(lngLoop).Item(mCol.标本ID).Value, True
'                            frmLabMain.ReadImageData .Records(lngLoop).Item(mCol.标本ID).Value, True
                            gstrSql = "select id from 检验图像结果 where 标本id = [1] order by id "
                            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(.Records(lngloop).Item(mCol.标本id).Value))
                            
                            For intLoop = 1 To 9
                                strChart(intLoop) = ""
                            Next
                            intLoop = 1
                            Do Until rsTmp.EOF
                                If intLoop > 9 Then Exit For
                                strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                                Call LoadImageData(App.path, rsTmp("ID"))
                                intLoop = intLoop + 1
                                rsTmp.MoveNext
                            Loop
                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                            astrSQL(UBound(astrSQL)) = "打印字串;" & strReportCode & ";;" & strReportParaNo & ";" & bytReportParaMode & ";" & Val(.Records(lngloop).Item(mCol.医嘱id).Value) & ";" & _
                                            Val(.Records(lngloop).Item(mCol.病人ID).Value) & ";" & Val(.Records(lngloop).Item(mCol.标本id).Value) & _
                                            ";" & strChart(1) & ";" & strChart(2) & ";" & strChart(3) & ";" & strChart(4) & ";" & strChart(5) & ";" & strChart(6) & ";" & strChart(7) & _
                                            ";" & strChart(8) & ";" & strChart(9)
'                            If intPrintCount = 0 Then Call ReportTaskBegin
'                            Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, _
'                                "医嘱ID=" & Val(.Records(lngLoop).Item(mCol.医嘱id).Value), _
'                                "病人ID=" & Val(.Records(lngLoop).Item(mCol.病人ID).Value), _
'                                "标本ID=" & Val(.Records(lngLoop).Item(mCol.标本id).Value), _
'                                "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
'                                "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
'                                "图形9=" & strChart(9), 2)
'                                intPrintCount = intPrintCount + 1
                            '标记打印
                            If mintUnion = 0 Then
                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_标本质控(" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ",'',1)"
                            Else
                                gstrSql = "select ID from 检验标本记录 where 医嘱ID = [1] "
                                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.医嘱id).Value))
                                Do Until rsTmp.EOF
                                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                    astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_标本质控(" & rsTmp("ID") & ",'',1)"
                                    rsTmp.MoveNext
                                Loop
                            End If
                        End If
                    End If
                    
                    .Records(lngloop).Item(mCol.选择).Checked = False
                    .Populate
                End If
            End If
            '删除无主
            If .Records(lngloop).Item(mCol.选择).Checked = True And mintEditType = 3 Then
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_标本删除(" & Val(.Records(lngloop).Item(mCol.标本id).Value) & ")"
                .Records(lngloop).Item(mCol.选择).Checked = False
                .Populate
            End If
            '10张报告单在一个任务中打印
            If intPrintCount >= 10 Then intPrintCount = 0:   Call ReportTaskEnd
            DoEvents
        Next
        
        '===================处理合并打印中的最后一个标本====================
        If Me.chkPatient.Value = 1 And str标本ID <> "" And InStr("," & mstrPrintDepts & ",", lng申请科室ID) > 0 Then
            intLoop = 1
            For intLoop = 1 To 9
                strChart(intLoop) = ""
            Next
            str标本ID = Mid(str标本ID, 2)
            str医嘱ID = Mid(str医嘱ID, 2)
            lng医嘱ID = Split(str医嘱ID, ",")(0)
            lng标本ID = Split(str标本ID, ",")(0)
            If strPrintCode = "" Then
                '有多个格式时得到格式
                frmLabMainPrintFormat.ShowMe Me, str医嘱ID, strPrintCode
            End If
            astrItem = Split(str标本ID, ",")
            For intItem = 0 To UBound(astrItem)
                gstrSql = "select id from 检验图像结果 where 标本id = [1] order by id"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(astrItem(intItem)))
                Do Until rsTmp.EOF
                    If intLoop > 9 Then Exit For
                    strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                    Call LoadImageData(App.path, rsTmp("ID"))
                    intLoop = intLoop + 1
                    rsTmp.MoveNext
                Loop
            Next
            If intPrintCount = 0 Then Call ReportTaskBegin
            Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & str医嘱ID, _
                    "病人ID=" & lngPatient, "标本ID=" & str标本ID, "多个医嘱=" & str医嘱ID, "多个标本=" & str标本ID, _
                    "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                    "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                    "图形9=" & strChart(9), 2)
            intPrintCount = intPrintCount + 1
            For intItem = 0 To UBound(astrItem)
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_标本质控(" & astrItem(intItem) & ",'',1)"
            Next
            str标本ID = "": str医嘱ID = ""
        End If
        '==================================================================================
        If intPrintCount > 0 Then Call ReportTaskEnd
    End With
    
'    gcnOracle.BeginTrans
'    blnRollBack = True
    
    For lngloop = 1 To UBound(astrSQL)
        If Trim(astrSQL(lngloop)) <> "" Then
            If UCase(Mid(astrSQL(lngloop), 1, 3)) = "ZL_" Then
                zlDatabase.ExecuteProcedure astrSQL(lngloop), Me.Caption
            ElseIf UCase(Mid(astrSQL(lngloop), 1, 4)) = "打印字串" Then
                strTmp = Split(astrSQL(lngloop), ";")
                If intPrintCount = 0 Then Call ReportTaskBegin
                Call ReportOpen(gcnOracle, glngSys, strTmp(1), Me, "NO=" & strTmp(3), "性质=" & strTmp(4), "医嘱ID=" & strTmp(5), "病人ID=" & strTmp(6), _
                "标本ID=" & strTmp(7), "图形1=" & strTmp(8), "图形2=" & strTmp(9), "图形3=" & strTmp(10), _
                , "图形4=" & strTmp(11), "图形5=" & strTmp(12), "图形6=" & strTmp(13), "图形7=" & strTmp(14), "图形8=" & strTmp(15), "图形9=" & strTmp(16), 2)
                intPrintCount = intPrintCount + 1
                
                '10张报告单在一个任务中打印
                If intPrintCount >= 10 Then intPrintCount = 0:   Call ReportTaskEnd
            Else
                '签名不成功时退出
                If Signature(Val(Split(astrSQL(lngloop), ";")(1)), mstrAuditingManID) = False Then
'                    gcnOracle.RollbackTrans
'                    blnRollBack = False
                     zlCommFun.StopFlash
                    mblnExec = False
                    Exit Sub
                End If
            End If
        End If
    Next
    If intPrintCount > 0 Then Call ReportTaskEnd
    
    blnRollBack = False
'    gcnOracle.CommitTrans
    
    zlCommFun.StopFlash
    Me.MousePointer = 0
    
    If strMsg <> "" Then
        MsgBox "以下记录未通过审核！" & strMsg, vbInformation, Me.Caption
    End If
    
    If intOnPrint > 0 Then
        If MsgBox("发现未审核的报告单" & intOnPrint & "张，是否打印?", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            Call SaveData(True)
        End If
    End If
    Call RefreshData
    On Error Resume Next
    '删除图形文件
    For intLoop = 1 To 9
        If strChart(intLoop) <> "" Then
            Kill strChart(intLoop)
        End If
    Next
    mblnExec = False
    Exit Sub
    
ErrHand:
    If blnRollBack = True Then gcnOracle.RollbackTrans: blnRollBack = False
    zlCommFun.StopFlash
    Me.MousePointer = 0
    mblnExec = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ShowMe(objfrm As Object, intEditType As Integer, Optional lngMachine As Long, Optional strPrivs As String, Optional strAuditingMan As String, _
                  Optional intAuditing As Integer, Optional DateAuditing As Date, Optional DeptID As Long, Optional strAuditingManID As String)
    '----------------------------------------------------------------------------------------------------------------
    '功能：显示本编辑窗体
    '参数：             lngMachine = 仪器ID strprivs = 权限 strAuditingMan = 申请人 intAuditing = 时间限制
    '                   DataAuditing = 开始时间  objfrm =  父窗体, intEditType = 操作类型(1=打印 2=审核 3=无主删除 4=批量修改标本号)
    '返回：
    '-----------------------------------------------------------------------------------------------------------------
    
    mintEditType = intEditType
    mlngMachine = lngMachine
    mstrPrivs = strPrivs
    mstrAuditingMan = strAuditingMan
    mstrAuditingManID = strAuditingManID
    mintAuditing = intAuditing
    mDateAuditing = DateAuditing
    mDeptID = DeptID
    stbThis.Panels(2).Text = "准备！"
    If mintEditType = 1 Then
        Me.chkUnion.Visible = True
        Me.chkPatient.Visible = True
    End If
    If mintEditType = 4 Then
        Frame1.Top = Me.chkPatient.Top - 50
        LabModify(0).Top = Frame1.Top + Frame1.Height + 100
        LabModify(1).Top = LabModify(0).Top
        TxtModify.Top = LabModify(0).Top - 50
        Label8.Top = LabModify(0).Top + LabModify(0).Height + 100
        rptMachine.Top = Me.Label8.Top + Me.Label8.Height + 100
        Me.LabModify(0).Visible = True: Me.LabModify(1).Visible = True: Me.TxtModify.Visible = True
    End If
    Me.Show , objfrm
    
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptList
        Set hitColumn = .HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "选择" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mCol.选择).Checked
                For Each Record In .Records
                    Record.Item(mCol.选择).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub rptMachine_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptMachine
        Set hitColumn = .HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "选择" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mMCol.选择).Checked
                For Each Record In .Records
                    Record.Item(mMCol.选择).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub txtSample_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call RefreshData: TxtSample.SelStart = 0: TxtSample.SelLength = Len(TxtSample)
End Sub

Private Sub TxtSample_Validate(Cancel As Boolean)
    '   10861 程序中是否应限制%,?等字符的输入
    If Check_Sample = False Then
        Cancel = True
    End If
End Sub

Private Function Check_Sample() As Boolean
    '   10861 程序中是否应限制%,?等字符的输入
    Dim i As Long, str字符 As String
    str字符 = ""
    If Len(TxtSample) > 0 Then
        For i = 1 To Len(TxtSample)
            If InStr("0123456789,~", Mid(TxtSample, i, 1)) <= 0 Then
                str字符 = str字符 & Mid(TxtSample, i, 1)
            End If
        Next
    End If

    If str字符 <> "" Then
        MsgBox "不能输入" & str字符, vbQuestion, gstrSysName
        Check_Sample = False
    Else
        Check_Sample = True
    End If

End Function
Private Function ModifySampleNumber() As Boolean
    '功能               批量修改标本号
    '参数               intModifyNumber   开始修改的标本号
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim astrSQL() As String
    Dim intCount As Integer
    Dim lngMachine As Long
    Dim rsTmp  As New ADODB.Recordset
    Dim blnUnChecked As Boolean, strMsg As String
    Dim strStartDate As String
    Dim strEndDate As String
    Dim blnBegin As Boolean
    
    On Error GoTo errH
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在准备修改标本号请等待..."
    
    If Trim(Me.TxtModify.Text) = "" Then
        strMsg = "请输入修改标本号的开始号码!"
        blnUnChecked = True
    ElseIf Not IsNumeric(Trim(Me.TxtModify.Text)) Then
        '11484 开始号码为非数字时，报错。
        strMsg = "请把开始号码改为数字!"
        blnUnChecked = True
    End If
        
    If blnUnChecked Then
        MsgBox strMsg, vbQuestion, gstrSysName
        zlCommFun.StopFlash
        Me.MousePointer = 0
        Me.TxtModify.SetFocus
        Exit Function
    End If
    
    With Me.rptList
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mCol.选择).Checked = True Then
                intCount = intCount + 1
                ReDim Preserve astrSQL(1 To intCount)
                astrSQL(intCount) = "ZL_检验标本记录_标本序号(" & .Records(intLoop).Item(mCol.标本id).Value & ",'" & intCount + TxtModify.Text - 1 & _
                "',null,null,to_date('" & Now & "','yyyy-mm-dd hh24:mi:ss')," & "to_date('" & Now & "','yyyy-mm-dd hh24:mi:ss'))"
                lngMachine = Val(.Records(intLoop).Item(mCol.仪器id).Value)
                
                strStartDate = GetDateTime(mMakeNoRule, 1, .Records(intLoop).Item(mCol.核收时间).Value)
                strEndDate = GetDateTime(mMakeNoRule, 2, .Records(intLoop).Item(mCol.核收时间).Value)
                
                gstrSql = "Select Id From 检验标本记录 Where 标本序号 = [1] " & IIf(lngMachine > 0, " And 仪器id = [2] ", "") & " And " & _
                          " 核收时间 Between [3] And [4] And ID <> [5] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, intCount + TxtModify.Text - 1, lngMachine, CDate(strStartDate), _
                            CDate(strEndDate), CStr(.Records(intLoop).Item(mCol.标本id).Value))
                
                zlCommFun.ShowFlash "正在更新标本" & .Records(intLoop).Item(mCol.标本序号).Value
                If rsTmp.EOF = False Then
                    zlCommFun.StopFlash
                    Me.MousePointer = 0
                    MsgBox "标本号" & intCount + TxtModify.Text - 1 & "已被使用！请检查后再重新修改标本号!", vbInformation, Me.Caption
                    Exit Function
                End If
            End If
        Next
        If intCount = 0 Then
            zlCommFun.StopFlash
            Me.MousePointer = 0
            Exit Function
        End If
        gcnOracle.BeginTrans
        blnBegin = True
        For intLoop = 1 To UBound(astrSQL)
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        Next
        gcnOracle.CommitTrans
    End With
    zlCommFun.StopFlash
    Me.MousePointer = 0
    RefreshData
Exit Function
errH:
    If blnBegin = True Then
        gcnOracle.RollbackTrans
    End If
    RefreshData
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function


