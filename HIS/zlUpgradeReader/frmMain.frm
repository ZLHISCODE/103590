VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "升级说明阅读器"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   13575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":6852
   ScaleHeight     =   7815
   ScaleWidth      =   13575
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5310
      Left            =   2655
      TabIndex        =   1
      Top             =   2145
      Width           =   7260
      _Version        =   589884
      _ExtentX        =   12806
      _ExtentY        =   9366
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox pic风险评估 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   11550
      ScaleHeight     =   2010
      ScaleWidth      =   2550
      TabIndex        =   18
      Top             =   5730
      Width           =   2550
      Begin RichTextLib.RichTextBox txt风险评估 
         Height          =   525
         Left            =   510
         TabIndex        =   26
         Top             =   510
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   926
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":D0A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox pic说明 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   10305
      ScaleHeight     =   2010
      ScaleWidth      =   2550
      TabIndex        =   17
      Top             =   4605
      Width           =   2550
      Begin RichTextLib.RichTextBox txt说明 
         Height          =   525
         Left            =   270
         TabIndex        =   25
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   926
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":D133
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   11145
      ScaleHeight     =   4875
      ScaleWidth      =   3075
      TabIndex        =   16
      Top             =   2565
      Width           =   3075
      Begin RichTextLib.RichTextBox txt需求 
         Height          =   525
         Left            =   420
         TabIndex        =   24
         Top             =   45
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   926
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":D1C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox cbo系统 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMain.frx":D251
      Left            =   11430
      List            =   "frmMain.frx":D253
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   165
      Width           =   2955
   End
   Begin VB.Frame fraFind 
      Caption         =   "查询条件"
      Height          =   1125
      Left            =   2670
      TabIndex        =   3
      Top             =   720
      Width           =   10710
      Begin VB.ComboBox cbo影响类型 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6975
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdReLoad 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   9495
         TabIndex        =   23
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   8310
         TabIndex        =   22
         Top             =   180
         Width           =   1100
      End
      Begin VB.ComboBox cbo用户 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   2715
      End
      Begin VB.ComboBox cbo培训 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4695
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   1350
      End
      Begin VB.ComboBox cbo是否阅读 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   800
      End
      Begin VB.ComboBox cbo风险等级 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox cbo结束版本 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6825
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   225
         Width           =   1350
      End
      Begin VB.ComboBox cbo开始版本 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4695
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "影响类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6090
         TabIndex        =   13
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "#"
         Height          =   180
         Left            =   9120
         TabIndex        =   21
         Top             =   675
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "登记用户"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "是否培训"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3810
         TabIndex        =   11
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否阅读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2055
         TabIndex        =   9
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lbl风险 
         AutoSize        =   -1  'True
         Caption         =   "风险等级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lbl开始版本 
         AutoSize        =   -1  'True
         Caption         =   "版本范围               －"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3810
         TabIndex        =   4
         Top             =   300
         Width           =   2625
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   9615
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D255
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D7EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD89
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E123
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E6BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EC57
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11039
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1341B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17BDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17F79
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTree 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6435
      Left            =   105
      ScaleHeight     =   6435
      ScaleWidth      =   2565
      TabIndex        =   0
      Top             =   675
      Width           =   2565
      Begin MSComctlLib.ImageList imgTree 
         Left            =   585
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView tvwLeft 
         Height          =   6090
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   10742
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgTree"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   7050
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgMenu 
      Left            =   900
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":18313
   End
   Begin XtremeCommandBars.CommandBars cbsMenu 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":1BCC3
      Left            =   1650
      Top             =   345
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsSheet As ADODB.Recordset
Private mlngItemID As Long

Private Const Dkp_ID_Tree As Integer = 1                          '列表
Private Const Dkp_ID_Find As Integer = 2                          '查找
Private Const Dkp_ID_Rept As Integer = 3
Private Const Dkp_ID_Right As Integer = 4
Private Const Dkp_ID_说明 As Integer = 5
Private Const Dkp_ID_风险 As Integer = 6

Private mIntType As Integer '显示方式  0-未登录方式 1-已登录方式

Enum 用户影响
    未填写
    正面作用
    负面作用
    无影响
End Enum

Private ItemHot As ReportRecordItem         '当前培训项焦点
Private rowLink As ReportRow        '当前超连接焦点行
Private mblnEdit As Boolean                 '是否修改过项目值
Private mstrFileName As String
Private mLastFileName As String
Private Type T条件
    用户     As String
    模块     As String
    开始版本 As String
    结束版本 As String
    风险等级 As String
    是否阅读 As String
    是否培训 As String
    影响类型 As String
End Type

Private m条件 As T条件
Private mstr关联问题 As String
Private mstr标题 As String
Private mLastNode As Node

Private Sub cbo结束版本_Click()
    If cbo结束版本.List(cbo结束版本.ListIndex) < cbo开始版本.List(cbo开始版本.ListIndex) Then
        
    End If
End Sub

Private Sub cbo结束版本_Validate(Cancel As Boolean)
    Dim intIndex As Integer
    If cbo结束版本.List(cbo结束版本.ListIndex) < cbo开始版本.List(cbo开始版本.ListIndex) Then
        For intIndex = 0 To cbo结束版本.ListCount - 1
            If cbo结束版本.List(intIndex) > cbo开始版本.List(cbo开始版本.ListIndex) Then
                cbo结束版本.ListIndex = intIndex
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cbsMenu_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.Id
        Case conMenu_File_Save
            Control.Enabled = mblnEdit
            cmdSave.Enabled = Control.Enabled
        Case conMenu_View_ShowPrivewText                '显示用户需求
            Control.Checked = rptList.PreviewMode
        Case conMenu_View_ShowGroupBox                  '显示分组框
            Control.Checked = rptList.ShowGroupBox
        Case conMenu_View_ShowRelation
            Control.Enabled = mstr关联问题 <> ""         '显示关联问题
    End Select
End Sub


Private Sub cmdReLoad_Click()
    Call ReLoad
End Sub

Private Sub cmdSave_Click()
     Call SaveItem: mblnEdit = False
End Sub

Private Sub Form_Load()
    
    Call initCommbar    '初始化菜单
    Call initDockPane   '初始化界面各区域
    
    Call LoadInitIcon   '装入系统图标

    Call initRptList(rptList, ImgList, lbl风险.Font, True)     '初始化问题列表
    
    Call initSYS        '初始化可选系统
    '设置默认系统
    
        
    ' 更新标题
    If mIntType = 1 Then
        Me.Caption = Me.Caption & "－" & "阅读者：" & gstrDBUser
        
    ElseIf mIntType = 0 Then
        Me.Caption = Me.Caption & "－(未登录) "
    End If
    mstr标题 = Me.Caption
    
    txt说明.Locked = True
    txt需求.Locked = True
    txt风险评估.Locked = True
    
    mblnEdit = False

    '恢复界面设置
    Dim strTmp As String
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & TypeName(rptList), rptList.Name, "")
    If strTmp <> "" Then rptList.LoadSettings strTmp
    If cbo系统.ListCount > 0 Then cbo系统.ListIndex = 0

End Sub

Private Sub cbo系统_Click()
    Dim intI As Integer
    If mblnEdit Then
        If MsgBox("所做的修改未保存，是否继续？", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
        mblnEdit = False
    End If
    mLastFileName = mstrFileName
    mstrFileName = ""
    mblnEdit = False
    
    cbo风险等级.Clear
    cbo风险等级.AddItem "所有"
    
    cbo风险等级.AddItem "高"
    cbo风险等级.AddItem "中"
    cbo风险等级.AddItem "低"
    cbo风险等级.AddItem "不确定"
    
    cbo是否阅读.Clear
    cbo是否阅读.AddItem "所有"
    cbo是否阅读.AddItem "已阅"
    cbo是否阅读.AddItem "未阅"
    
    cbo培训.Clear
    cbo培训.AddItem "所有"
    cbo培训.AddItem "未填写"
    cbo培训.AddItem "已培训"
    cbo培训.AddItem "无需培训"

    m条件.风险等级 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "风险等级", "所有")
    m条件.是否培训 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "是否培训", "所有")
    m条件.是否阅读 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "是否阅读", "所有")
    m条件.用户 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "用户", "所有")
    m条件.模块 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "模块", "所有模块")
    m条件.开始版本 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "开始版本", "0.0.0")
    m条件.结束版本 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "结束版本", "100.100.100")
    m条件.影响类型 = GetSetting("ZLSOFT", "公共模块\UpgradeReader", "影响类型", "所有")
    
    cbo风险等级.ListIndex = 0
    If m条件.风险等级 <> "所有" Then
        For intI = 0 To cbo风险等级.ListCount - 1
            If m条件.风险等级 = cbo风险等级.List(intI) Then
                cbo风险等级.ListIndex = intI
                Exit For
            End If
        Next
    End If
    
    cbo是否阅读.ListIndex = 0
    If m条件.是否阅读 <> "所有" Then
        For intI = 0 To cbo是否阅读.ListCount - 1
            If m条件.是否阅读 = cbo是否阅读.List(intI) Then
                cbo是否阅读.ListIndex = intI
                Exit For
            End If
        Next
    End If
    
    cbo培训.ListIndex = 0
    If m条件.是否培训 <> "所有" Then
        For intI = 0 To cbo培训.ListCount - 1
            If m条件.是否培训 = cbo培训.List(intI) Then
                cbo培训.ListIndex = intI
                Exit For
            End If
        Next
    End If
    
    Call initTree       '初始化模块列表
    If mLastFileName = "" Then mLastFileName = mstrFileName
    Me.Caption = mstr标题 & " " & mstrFileName
End Sub

Private Sub cbsMenu_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long


    Select Case Control.Id

    Case conMenu_View_Expend_CurCollapse                            '折叠当前组
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                rptList.SelectedRows(0).Expanded = False
            ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                If rptList.SelectedRows(0).ParentRow.GroupRow Then
                    rptList.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '因折叠定位到分组上,不会自动激活该事件
        Call rptList_SelectionChanged

    Case conMenu_View_Expend_CurExpend                              '展开当前组
        If rptList.SelectedRows.Count > 0 Then
            rptList.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse                            '折叠所有组
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '因折叠定位到分组上,不会自动激活该事件
        Call rptList_SelectionChanged
    Case conMenu_View_Expend_AllExpend                              '展开所有组
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    Case conMenu_View_ShowPrivewText                                '显示用户需求
        rptList.PreviewMode = Not rptList.PreviewMode
    Case conMenu_View_ShowGroupBox
        rptList.ShowGroupBox = Not rptList.ShowGroupBox             '显示分组框
    Case conMenu_View_ShowRelation
        Call frmRelation.ShowRelation(mstrFileName, mstr关联问题)   '显示关联问题
        
    Case conMenu_File_Save                                          '保存
        Call SaveItem: mblnEdit = False
    Case conMenu_File_Exit        '退出
        Unload Me
        
    Case conMenu_View_Refresh     '刷新
        Call ReLoad
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Id = Dkp_ID_Tree Then
        Item.Handle = picTree.hwnd
    ElseIf Item.Id = Dkp_ID_Find Then
        Item.Handle = fraFind.hwnd
    ElseIf Item.Id = Dkp_ID_Rept Then
        Item.Handle = rptList.hwnd
    ElseIf Item.Id = Dkp_ID_Right Then
        Item.Handle = picRight.hwnd
    ElseIf Item.Id = Dkp_ID_说明 Then
        Item.Handle = pic说明.hwnd
    ElseIf Item.Id = Dkp_ID_风险 Then
        Item.Handle = pic风险评估.hwnd
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEdit Then
        If MsgBox("所做的修改未保存，是否继续？", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
        mblnEdit = False
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & TypeName(rptList), rptList.Name, rptList.SaveSettings
    '关闭公共部件的窗体
    CloseWindows
    '关闭应用工具包部件的窗体
    mclsAppTool.CloseWindows
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With Me.txt需求
        .Left = picRight.ScaleLeft
        .Top = picRight.ScaleTop
        .Width = picRight.ScaleWidth - 45
        .Height = picRight.ScaleHeight - 45
    End With
End Sub

Private Sub picTree_Resize()
    On Error Resume Next
    Me.tvwLeft.Left = 0
    Me.tvwLeft.Top = 0
    Me.tvwLeft.Width = picTree.ScaleWidth
    Me.tvwLeft.Height = picTree.ScaleHeight - Me.tvwLeft.Top
End Sub

Private Sub pic风险评估_Resize()
    On Error Resume Next
    With Me.txt风险评估
        .Left = pic风险评估.ScaleLeft
        .Top = pic风险评估.ScaleTop
        .Width = pic风险评估.ScaleWidth - 45
        .Height = pic风险评估.ScaleHeight - 45
    End With
End Sub

Private Sub pic说明_Resize()
    On Error Resume Next
    With Me.txt说明
        .Left = pic说明.ScaleLeft
        .Top = pic说明.ScaleTop
        .Width = pic说明.ScaleWidth - 45
        .Height = pic说明.ScaleHeight - 45
    End With
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub rptList_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
     Dim RecordItem As ReportRecordItem
    If (Row.Record(mCol.影响评估).Value = 未填写) Then
        For Each RecordItem In Row.Record
            RecordItem.Bold = True
        Next
    Else
        For Each RecordItem In Row.Record
            RecordItem.Bold = False
        Next
    End If
        
    If (Item.Index = mCol.风险) Then
        Select Case Item.Value
            Case 0: Item.Icon = ICON_Unknown    '不确定
            Case 1: Item.Icon = ICON_Low        '低
            Case 2: Item.Icon = ICON_Center     '中
            Case 3: Item.Icon = ICON_High       '高
        End Select
    End If
    
    If (Item.Index = mCol.编号) Then
        If Row.Record(mCol.连接).Value = "是" Then
            Set Metrics.Font = fntUnderLine
            Metrics.ForeColor = vbBlue
        End If
    End If
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim strLinkFile As String
    If Button = 1 Then
'        If (Not ItemHot Is Nothing) Then
'            If ItemHot.Value = "无需培训" Then Exit Sub '无需培训
'             ItemHot.Value = IIf(ItemHot.Value = "已培训", "　", "已培训") '修改培训状态
'            'If ItemHot.Icon = -1 Then Exit Sub
''            ItemHot.Icon = IIf(ItemHot.Icon = ICON_Train, 6, 5)
''            ItemHot.Value = IIf(ItemHot.Icon = ICON_Train, 6, 5)
'            mblnEdit = True
'        End If
        
        If (Not rowLink Is Nothing) Then
            If rowLink.Record(mCol.连接).Value = "是" Then
                strLinkFile = Mid(mstrFileName, 1, InStrRev(mstrFileName, "\")) & "Document\" & rowLink.Record(mCol.编号).Value & ".htm"
                If Dir(strLinkFile) <> "" Then
                    Call ShellExecute(Me.hwnd, "open", "file:///" & Replace(strLinkFile, "\", "/"), vbNullString, vbNullString, 1)
                Else
                    MsgBox "未找到对应的html文件，打开文件失败！", vbInformation, gstrSysname
                End If
            End If
        End If
    End If
End Sub

Private Sub rptList_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    Dim htInfo As ReportHitTestInfo
    Set htInfo = rptList.HitTest(X, Y)
    
    Dim Item As ReportRecordItem
    Dim objRow As ReportRow

    If (Not htInfo.Item Is Nothing) Then
        If (htInfo.Item.Index = mCol.培训) Then
            Set Item = htInfo.Item
        End If
        
        If (htInfo.Item.Index = mCol.编号) Then
            Set objRow = htInfo.Row
        End If
    End If

    If (Not objRow Is rowLink) Then
        If (Not objRow Is Nothing) Then
            If objRow.Record(mCol.连接).Value = "是" Then
                objRow.Record(mCol.编号).BackColor = RGB(255, 238, 99)
            End If
            
        End If
        
        If (Not rowLink Is Nothing) Then
            rowLink.Record(mCol.编号).BackColor = -1
        End If
        
        Set rowLink = objRow
        rptList.Redraw
    End If
    
    If (Not Item Is ItemHot) Then
        If (Not Item Is Nothing) Then
            If Item.Value = "无需培训" Then Exit Sub '无需培训
'            If Item.Icon = -1 Then Exit Sub
'            Item.BackColor = IIf(Item.Icon = ICON_Train, RGB(207, 93, 96), RGB(255, 238, 194))
            Item.BackColor = IIf(Item.Value = "　", RGB(207, 93, 96), RGB(255, 238, 194))
        End If

        If (Not ItemHot Is Nothing) Then
            If ItemHot.Value = "" Then Exit Sub
            ItemHot.BackColor = -1
        End If
        Set ItemHot = Item
        rptList.Redraw
    End If
    

End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = cbsMenu.ActiveMenuBar.FindControl(, conMenu_View)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
        
        If (Not ItemHot Is Nothing) Then
            If ItemHot.Value = "无需培训" Then Exit Sub '无需培训
             ItemHot.Value = IIf(ItemHot.Value = "已培训", "　", "已培训") '修改培训状态
            'If ItemHot.Icon = -1 Then Exit Sub
'            ItemHot.Icon = IIf(ItemHot.Icon = ICON_Train, 6, 5)
'            ItemHot.Value = IIf(ItemHot.Icon = ICON_Train, 6, 5)
            mblnEdit = True
            rptList.Redraw
        End If
End Sub

Private Sub rptList_SelectionChanged()
    '#
    txt说明 = ""
    txt需求 = ""
    txt风险评估 = ""
    mstr关联问题 = ""
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If Not rptList.FocusedRow.GroupRow Then
        txt说明 = rptList.FocusedRow.Record(mCol.说明).Value
        txt需求 = rptList.FocusedRow.Record(mCol.需求).Value
        txt风险评估 = rptList.FocusedRow.Record(mCol.风险评估).Value
        mstr关联问题 = Trim(rptList.FocusedRow.Record(mCol.关联问题).Value)
    End If
    
End Sub

Private Sub rptList_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If (Item.Index = mCol.影响评估) Then
        Dim ItemRead As ReportRecordItem
        Set ItemRead = Item.Record(mCol.阅读)
        
        If (Item.Value = 未填写) Then
            ItemRead.Icon = ICON_NoRead
        Else
            ItemRead.Icon = ICON_Read
        End If
        Item.Record(mCol.修改).Value = "1"
        mblnEdit = True
    
    End If

End Sub

Private Sub tvwLeft_NodeClick(ByVal Node As MSComctlLib.Node)
    '调用
    If mstrFileName <> "" Then
        If mblnEdit Then
            If MsgBox("所做的修改未保存，是否继续？", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then
                If Not mLastNode Is Nothing Then tvwLeft.SelectedItem = mLastNode
                Exit Sub
            End If
        End If
        Call LoadSheet(mstrFileName)
        Set mLastNode = Node
    End If
    
End Sub

'-----------------以下是本窗体自定义过程
Public Sub Show_me(ByVal intType As Integer)
    mIntType = intType
    Me.Show
    
End Sub

Private Sub OpenExcel(ByVal str系统 As String)
    Dim strSheet As String
    Dim strFilename As String
    Dim strPath As String
   
    If mstrFileName = "" Then
        
        strPath = App.Path & "\"
        strPath = Mid(strPath, 1, InStrRev(strPath, "\"))
        strPath = strPath & ReadFromIni(App.Path & "\" & App.EXEName & ".ini", str系统, "Path")
        strFilename = Dir(strPath & "\*.xls")
        
        If strFilename = "" Then
            strPath = GetSetting("ZLSOFT", "公共全局", "程序路径", App.Path & "\")
            strPath = Mid(strPath, 1, InStrRev(strPath, "\"))
            strPath = strPath & ReadFromIni(App.Path & "\" & App.EXEName & ".ini", str系统, "Path")
            strFilename = Dir(strPath & "\*.xls")
        End If
        
        If strFilename = "" Then
            If MsgBox("在默认目录未找到EXCEL文件，请手工指定升级说明文件。", vbYesNo + vbDefaultButton1, gstrSysname) = vbYes Then
                dlgFile.DialogTitle = "打开一个升级说明文件进行浏览"
                dlgFile.InitDir = App.Path
                dlgFile.Filename = ""
                dlgFile.Filter = "升级说明文件|*.xls"
                dlgFile.ShowOpen
                If dlgFile.Filename = "" Then Exit Sub
                strFilename = dlgFile.Filename
            End If
        Else
            strFilename = strPath & "\" & strFilename
        End If
        mstrFileName = strFilename
    End If
    
    
    
End Sub

Private Function initSYS() As Boolean
    '添加系统
    Dim strIniFileName As String, i As Integer, strRemoveItem As String
    Dim str系统 As String, varTmp As Variant, str已安装系统 As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    str系统 = Trim(ReadFromIni(App.Path & "\" & App.EXEName & ".ini", "系统", "名称"))
    If str系统 = "" Then
        MsgBox "配置文件丢失，不能继续运行！", vbQuestion, gstrSysname
        Exit Function
    End If
    
    cbo系统.Clear
    
    If mIntType = 0 Then
        '未登录方式，显示配置文件中的系统
        If InStr(str系统, "|") > 0 Then
            varTmp = Split(str系统, "|")
            For i = LBound(varTmp) To UBound(varTmp)
                If Trim(varTmp(i)) <> "" Then
                    cbo系统.AddItem varTmp(i)
                End If
            Next
        Else
            cbo系统.AddItem str系统
        End If
    ElseIf mIntType = 1 Then
        '如果是已登录的，显示有权访问的系统
        str已安装系统 = ""
        Set rsTmp = gcnOracle.Execute("Select 名称,编号 From zlsystems Where 编号 In(" & gstrSystems & ") Order by 编号")
        Do Until rsTmp.EOF
            str已安装系统 = Trim("" & rsTmp!名称)
            If str已安装系统 <> "" Then
                cbo系统.AddItem str已安装系统
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    
    initSYS = True
    Exit Function
errHandle:
    initSYS = False
    MsgBox Err.Number & " " & Err.Description, vbInformation, gstrSysname
End Function


Private Sub initCommbar()
    '初始化菜单栏，工具栏
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim objPopup As CommandBarPopup
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMenu.VisualTheme = xtpThemeOffice2003
    With Me.cbsMenu.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    cbsMenu.EnableCustomization False
    Set cbsMenu.Icons = imgMenu.Icons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMenu.ActiveMenuBar.Title = "菜单"
    cbsMenu.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMenu.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_File, "文件(&F)", -1, False)
    objMenu.Id = conMenu_File
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "保存(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
    End With


    
    Set objMenu = cbsMenu.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View, "查看(&V)", -1, False)
    objMenu.Id = conMenu_View
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowGroupBox, "显示分组框(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowPrivewText, "预览用户需求(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowRelation, "查看关联问题(&R)")
        
'        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "查找(&F)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找一下个(&N)")
'        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "筛选(&I)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordPrev, "前一条记录(&P)")
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordNext, "后一条记录(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True

    End With
    

    
'    Set objMenu = cbsMenu.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Help, "帮助(&H)", -1, False)
'    objMenu.Id = conMenu_Help
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): objControl.BeginGroup = True
'    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMenu.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowGroupBox, "分组框"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowPrivewText, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowRelation, "关联")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): objControl.BeginGroup = True
        
'        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "查找"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "筛选"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordPrev, "前一条")
'        Set objControl = .Add(xtpControlButton, conMenu_View_RecordNext, "后一条")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlLabel, conMenu_Custom_System - 1, "系统")
        objControl.Flags = xtpFlagRightAlign
    
        Set objCustom = .Add(xtpControlCustom, conMenu_Custom_System, "系统")
        objCustom.ShortcutText = "系统"
        objCustom.Handle = Me.cbo系统.hwnd
        objCustom.Flags = xtpFlagRightAlign
        objCustom.Style = xtpButtonIconAndCaption
    End With
    
    '前一条，后一条记录不显示文字说明
    For Each objControl In objBar.Controls
        If objControl.Id <> conMenu_View_RecordPrev And objControl.Id <> conMenu_View_RecordNext And objControl.Id <> conMenu_Custom_System And objControl.Id <> conMenu_Custom_System - 1 Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMenu.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_File_Save
'        .Add FCONTROL, vbKeyF, conMenu_View_Find
'        .Add 0, vbKeyF3, conMenu_View_FindNext
'        .Add FCONTROL, vbKeyI, conMenu_View_Filter
'        .Add FCONTROL, vbKeyLeft, conMenu_View_RecordPrev
'        .Add FCONTROL, vbKeyRight, conMenu_View_RecordNext
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
    
    'MDI Tab
'    '-----------------------------------------------------
'    cbsMenu.ActiveMenuBar.SetFlags xtpFlagHideMDIButtons, 0
'    Set mWorkSpace = cbsMenu.ShowTabWorkspace(True)
'    cbsMenu.TabWorkspace.AutoTheme = False
'    cbsMenu.TabWorkspace.PaintManager.Appearance = xtpTabAppearanceVisualStudio
'    cbsMenu.TabWorkspace.PaintManager.Color = xtpTabColorOffice2003
'    cbsMenu.TabWorkspace.PaintManager.ClientFrame = xtpTabFrameSingleLine
    
    '状态栏
    '-----------------------------------------------------
'    cbsMenu.StatusBar.Visible = True
'    cbsMenu.StatusBar.AddPane 1
'    cbsMenu.StatusBar.SetPaneStyle 1, SBPS_STRETCH
'    cbsMenu.StatusBar.SetPaneText 1, ""
'    cbsMenu.StatusBar.AddPane 2
'    cbsMenu.StatusBar.SetPaneWidth 2, 100
'    cbsMenu.StatusBar.SetPaneText 2, ""
'    cbsMenu.StatusBar.AddPane 3
'    cbsMenu.StatusBar.SetPaneWidth 3, 60
'    cbsMenu.StatusBar.SetPaneText 3, ""
'    cbsMenu.StatusBar.IdleText = ""
    
    picRight.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
    picTree.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
    pic说明.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
    pic风险评估.BackColor = cbsMenu.GetSpecialColor(STDCOLOR_BTNFACE)
End Sub

Private Sub initDockPane()
    Dim paneTree As Pane, paneFind As Pane, paneEdit As Pane, paneList As Pane, paneRight As Pane, pane说明 As Pane, pane风险 As Pane
    
    Me.dkpMain.SetCommandBars Me.cbsMenu
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    
    Me.dkpMain.Options.HideClient = True
    
    Set paneList = Me.dkpMain.CreatePane(Dkp_ID_Rept, 900, 700, DockTopOf, Nothing)
    paneList.Title = "问题清单"
    paneList.Handle = Me.rptList.hwnd
    paneList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneTree = Me.dkpMain.CreatePane(Dkp_ID_Tree, 180, 90, DockLeftOf, Nothing)
    paneTree.Title = "模块"
    paneTree.Handle = Me.picTree.hwnd
    
    Set paneFind = Me.dkpMain.CreatePane(Dkp_ID_Find, 50, 180, DockTopOf, paneList)
    paneFind.Title = "查询条件"
    paneFind.Handle = Me.fraFind.hwnd
    paneFind.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane说明 = Me.dkpMain.CreatePane(Dkp_ID_说明, 800, 500, DockBottomOf, paneList)
    pane说明.Title = "修改说明"
    pane说明.Handle = Me.pic说明.hwnd
    pane说明.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneRight = Me.dkpMain.CreatePane(Dkp_ID_Right, 100, 300, DockBottomOf, pane说明)
    paneRight.Title = "用户需求"
    paneRight.Handle = Me.picRight.hwnd
    paneRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane风险 = Me.dkpMain.CreatePane(Dkp_ID_风险, 100, 300, DockRightOf, paneRight)
    pane风险.Title = "风险评估"
    pane风险.Handle = Me.pic风险评估.hwnd
    pane风险.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
End Sub

Private Sub initTree()
    '
    Dim objNode As Node, intGrant As Integer
    Dim objRootNode As Node
    Dim strSheet As String, varSheet As Variant, i As Integer, intList As Integer
    Dim strNodeText As String, str模块 As String
    Dim blnAdd As Boolean
    Dim str最低版本 As String, str最高版本 As String, strTmp版本 As String, str用户 As String, str影响模块 As String, var影响模块 As Variant, str影响类型 As String
    Dim str用户版本 As String, rsTmp As ADODB.Recordset, StrSQL As String
    
    On Error Resume Next
    
    If cbo系统.ListIndex < 0 Then Exit Sub
    Call OpenExcel(cbo系统.List(cbo系统.ListIndex))
    If mstrFileName = "" Then
        mstrFileName = mLastFileName
        Exit Sub
    End If
    
    Me.tvwLeft.Nodes.Clear
    Set objRootNode = Me.tvwLeft.Nodes.Add(, , "Root", "所有模块", "K_" & 141)
    
    cbo开始版本.Clear
    'cbo开始版本.AddItem "0.0.0"
    cbo结束版本.Clear
    'cbo结束版本.AddItem "100.100.100"
    cbo用户.Clear
    cbo用户.AddItem "所有"
    If mIntType = 1 And gstr用户单位名称 <> "" Then
        cbo用户.AddItem gstr用户单位名称
        cbo用户.ListIndex = cbo用户.ListCount - 1
    End If
    
    If mstrFileName <> "" Then
        strSheet = OpenExcelFile(mstrFileName)
        If InStr(strSheet, "|") <= 0 Then Exit Sub
    End If
    
    cbo影响类型.Clear
    cbo影响类型.AddItem "所有"
    cbo影响类型.ListIndex = 0
    
    '----
    varSheet = Split(strSheet, "|")
    For i = LBound(varSheet) To UBound(varSheet)
        Set mrsSheet = OpenExcelSheet(varSheet(i))
        Do Until mrsSheet.EOF
            '取版本
            str用户 = Trim("" & mrsSheet(Excel_Col.登记用户).Value)
'            If mIntType = 0 Or (mIntType = 1 And str用户 = gstr用户单位名称) Then
                strTmp版本 = "" & mrsSheet(Excel_Col.发布版本).Value
                If strTmp版本 <> "" Then
                    If str最低版本 = "" Then
                        cbo开始版本.AddItem strTmp版本
                        str最低版本 = strTmp版本
                    Else
                        If str最低版本 > strTmp版本 Then
                            str最低版本 = strTmp版本
                        End If
                        
                        blnAdd = True
                        For intList = 0 To cbo开始版本.ListCount - 1
                            If cbo开始版本.List(intList) = strTmp版本 Then
                                blnAdd = False
                            End If
                        Next
                        
                        If blnAdd Then
                            cbo开始版本.AddItem strTmp版本
                        End If
                    
                    End If
                
                    If str最高版本 = "" Then
                        cbo结束版本.AddItem strTmp版本
                        str最高版本 = strTmp版本
                    Else
                        If str最高版本 < strTmp版本 Then
                            str最高版本 = strTmp版本
                        End If
                        blnAdd = True
                        For intList = 0 To cbo结束版本.ListCount - 1
                            If cbo结束版本.List(intList) = strTmp版本 Then blnAdd = False
                        Next
                        If blnAdd Then
                            cbo结束版本.AddItem strTmp版本
                        End If
                    End If
                End If 'end If strTmp版本 <> ""
'            End If
            
            '取影响类型
            str影响类型 = Trim("" & mrsSheet(Excel_Col.影响类型).Value)
            If str影响类型 <> "" Then
                blnAdd = True
                For intList = 0 To cbo影响类型.ListCount - 1
                    If cbo影响类型.List(intList) = str影响类型 Then blnAdd = False
                Next
                If blnAdd Then cbo影响类型.AddItem str影响类型
            End If
            
            If mIntType = 0 Then
                '取用户名
                str用户 = Trim("" & mrsSheet(Excel_Col.登记用户).Value)
                If str用户 <> "" Then
                    blnAdd = True
                    For intList = 0 To cbo用户.ListCount - 1
                        If cbo用户.List(intList) = str用户 Then blnAdd = False
                    Next
                    If blnAdd Then cbo用户.AddItem str用户
                End If
                

            End If
            mrsSheet.MoveNext
        Loop
    Next
    '加入固定模块
    Dim str固定模块 As String, var固定模块 As Variant, int固定 As Integer
    str固定模块 = Trim(ReadFromIni(App.Path & "\" & App.EXEName & ".ini", "固定模块", "模块"))
    var固定模块 = Split(str固定模块, "|")
    For int固定 = LBound(var固定模块) To UBound(var固定模块)
        If var固定模块(int固定) <> "" Then
            Set objNode = Me.tvwLeft.Nodes.Add(, , "G" & Format(int固定 + 1, "000"), var固定模块(int固定), "K_" & 99)
            Call addModleToTree(var固定模块(int固定), "G" & Format(int固定 + 1, "000"))
        End If
    Next
    
    If mIntType = 0 Then
        '未登录 加模块
        varSheet = Split(strSheet, "|")
        For i = LBound(varSheet) To UBound(varSheet)
            Set mrsSheet = OpenExcelSheet(varSheet(i))
            Do Until mrsSheet.EOF
                Call AddExcelToTree("")
                mrsSheet.MoveNext
            Loop
        Next

    '-----
    ElseIf mIntType = 1 Then
'        If gblnOwner Then
'            '是系统所有者
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "G01", "管理工具", "K_" & 207)
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0101", "装卸管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010101", "系统装卸管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010102", "系统升迁管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010103", "对象检查修复", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010104", "置换安装脚本", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0101", 4, "G010105", "编译无效对象", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0102", "数据管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010201", "数据转移", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010202", "数据导出", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010203", "数据导入", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010204", "数据调出", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010205", "数据调入", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0102", 4, "G010206", "数据清除", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0103", "运行管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010301", "用户注册管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010302", "运行状态监控", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010303", "后台作业管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010304", "运行日志管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010305", "错误日志管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010306", "系统运行选项", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010307", "站点部件升级", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010308", "站点运行控制", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0103", 4, "G010309", "站点文件收集", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0104", "权限管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0104", 4, "G010401", "角色授权管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0104", 4, "G010402", "用户授权管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0104", 4, "G010403", "菜单重组规划", "K_" & 99)
'
'            Set objNode = Me.tvwLeft.Nodes.Add("G01", 4, "G0105", "专项工具", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0105", 4, "G010501", "报表管理", "K_" & 99)
'            Set objNode = Me.tvwLeft.Nodes.Add("G0105", 4, "G010502", "函数管理", "K_" & 99)
'        End If
        
'        Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & 工具清单.导航功能清单, "导航功能清单", "K_" & 210)
'
'        intGrant = zlRegTool
'        If ((intGrant And 4) = 4) Then
'            If InStr(1, GetPrivFunc(0, 工具清单.消息收发工具), "基本") <> 0 Then
'                Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & 工具清单.消息收发工具, "消息收发工具", "K_" & 145)
'            End If
'        End If
'        If ((intGrant And 8) = 8) Then
'            If InStr(1, GetPrivFunc(0, 工具清单.EXCEL报表工具), "基本") Then
'                Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & 工具清单.EXCEL报表工具, "EXCEL报表工具", "K_" & 217)
'            End If
'        End If
'
'        If InStr(1, GetPrivFunc(0, 工具清单.本地参数管理), "基本") Then
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & 工具清单.本地参数管理, "本地参数管理", "K_" & 135)
'        End If
'
'        If InStr(1, GetPrivFunc(0, 工具清单.系统选项设置), "基本") Then
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & 工具清单.系统选项设置, "系统选项设置", "K_" & 147)
'        End If
'
'        If InStr(1, GetPrivFunc(0, 工具清单.字典管理工具), "基本") Then
'            Set objNode = Me.tvwLeft.Nodes.Add(, , "T" & 工具清单.字典管理工具, "字典管理工具", "K_" & 144)
'        End If
        
        '已登录
        If cbo系统.List(cbo系统.ListIndex) = "体检管理系统" Then
            Set rsTmp = rsMenuPEIS
        Else
            Set rsTmp = rsMenu
        End If
        With rsTmp
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not .EOF
                'On Error Resume Next
                If UCase(.Fields("部件").Value) <> UCase("zl9Report") Then
                    If .Fields("模块").Value = 0 Then
                        If .Fields("上级") = 0 Then
                            Set objNode = Me.tvwLeft.Nodes.Add(, , "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, 99, !图标))
                        Else
                            Set objNode = Me.tvwLeft.Nodes.Add("_" & .Fields("上级").Value, 4, "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, 99, !图标))
                        End If
                    Else
                        '子功能
                        If .Fields("上级") = 0 Then
                            Set objNode = Me.tvwLeft.Nodes.Add(, , "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, 99, !图标))
                        Else
                            Set objNode = Me.tvwLeft.Nodes.Add("_" & .Fields("上级").Value, 4, "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, 99, !图标))
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        '已登录 未找到的模块加到“其他中”
        Set objNode = Me.tvwLeft.Nodes.Add(, , "Q01", "其他", "K_" & 172)
        varSheet = Split(strSheet, "|")
        For i = LBound(varSheet) To UBound(varSheet)
            Set mrsSheet = OpenExcelSheet(varSheet(i))
            Do Until mrsSheet.EOF
                Call AddExcelToTree("Q01")
                mrsSheet.MoveNext
            Loop
        Next
        
        
        '取用户版本
        StrSQL = "Select 版本号 From zlsystems Where 名称='" & cbo系统.List(cbo系统.ListIndex) & "'"
        Set rsTmp = gcnOracle.Execute(StrSQL)
        Do Until rsTmp.EOF
            str用户版本 = "" & rsTmp("版本号").Value
            rsTmp.MoveNext
        Loop
    End If
    
    '--开始版本
    If cbo开始版本.ListCount > 0 Then
        For intList = 0 To cbo开始版本.ListCount - 1
            If cbo开始版本.List(intList) = str最低版本 Then
               cbo开始版本.ListIndex = intList: Exit For
            End If
        Next
    End If
    For intList = 0 To cbo开始版本.ListCount - 1
        If cbo开始版本.List(intList) = m条件.开始版本 Then
             cbo开始版本.ListIndex = intList: Exit For
        End If
    Next
    If str用户版本 <> "" Then
        If str用户版本 < str最高版本 Then
            For intList = 0 To cbo开始版本.ListCount - 1
                If cbo开始版本.List(intList) = str用户版本 Then
                     cbo开始版本.ListIndex = intList: Exit For
                End If
            Next
        End If
    End If
    '--结束版本
    If cbo结束版本.ListCount > 0 Then
        For intList = 0 To cbo结束版本.ListCount - 1
            If cbo结束版本.List(intList) = str最高版本 Then
                cbo结束版本.ListIndex = intList: Exit For
            End If
        Next
    End If
    For intList = 0 To cbo结束版本.ListCount - 1
        If cbo结束版本.List(intList) = m条件.结束版本 Then
            cbo结束版本.ListIndex = intList: Exit For
        End If
    Next
    
    '--用户
    If cbo用户.ListCount > 0 Then cbo用户.ListIndex = 0
    For intList = 0 To cbo用户.ListCount - 1
        If cbo用户.List(intList) = m条件.用户 Then
           cbo用户.ListIndex = intList: Exit For
        End If
    Next
    
    cbo影响类型.ListIndex = 0
    If m条件.影响类型 <> "所有" Then
        For intList = 0 To cbo影响类型.ListCount - 1
            If m条件.影响类型 = cbo影响类型.List(intList) Then
                cbo影响类型.ListIndex = intList
                Exit For
            End If
        Next
    End If
    '－－模块
    For intList = 1 To tvwLeft.Nodes.Count
        If tvwLeft.Nodes(intList).Text = m条件.模块 Then
            'Set tvwLeft.SelectedItem = tvwLeft.Nodes(intList)
            Set objRootNode = tvwLeft.Nodes(intList)
            
            Exit For
        End If
    Next
        
    Set tvwLeft.SelectedItem = objRootNode
    Call tvwLeft_NodeClick(tvwLeft.SelectedItem)
    Call ReLoad
End Sub

Private Sub LoadSheet(ByVal Filename As String)

    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptItem1 As ReportRecordItem
    Dim rptItem2 As ReportRecordItem
    Dim rptRow As ReportRow
    Dim rptColum As ReportColumn
    
    Dim strSheet As String, varSheet As Variant, i As Integer, intNode As Integer, str模块 As String
    Dim lngCount As Long
    Dim strTxt As String
    On Error GoTo errHandle
    If Filename = "" Then Exit Sub
    strSheet = OpenExcelFile(Filename)
    
    If strSheet = "" Then Exit Sub
    
    '清空列表中的模块个数
    
    For i = 1 To tvwLeft.Nodes.Count
        tvwLeft.Nodes(i).Tag = 0
        strTxt = Replace(tvwLeft.Nodes(i).Text, "(新版)", "")
        If InStr(strTxt, "(") > 0 Then strTxt = Mid(strTxt, 1, InStr(strTxt, "(") - 1)
        tvwLeft.Nodes(i).Text = strTxt
    Next
    
    rptList.Records.DeleteAll '清空原列表
    lblA.Caption = ""
    If InStr(strSheet, "|") <= 0 Then Exit Sub
    
    varSheet = Split(strSheet, "|")
    For i = LBound(varSheet) To UBound(varSheet)
        Set mrsSheet = OpenExcelSheet(varSheet(i))
        lngCount = 0
        Do Until mrsSheet.EOF
            
            '加入明细
            With rptList
                    '添加模块个数，
                If IsAdd(False) Then
                    str模块 = Replace("" & mrsSheet(Excel_Col.登记模块).Value, "(新版)", "")
                    For intNode = 1 To tvwLeft.Nodes.Count
                        strTxt = tvwLeft.Nodes(intNode).Text
                        If InStr(strTxt, "(") > 0 Then strTxt = Mid(strTxt, 1, InStr(strTxt, "(") - 1)
                        If str模块 = strTxt Then
                            tvwLeft.Nodes(intNode).Tag = Val(tvwLeft.Nodes(intNode).Tag) + 1
                            tvwLeft.Nodes(intNode).Text = strTxt & "(" & Val(tvwLeft.Nodes(intNode).Tag) & ")"
                            Call AddParent(intNode)
                        End If
                    Next
                End If
                
                If IsAdd(True) Then      '符合条件则加入

                    lngCount = lngCount + 1
                    Set rptRcd = Me.rptList.Records.Add()
                    
                    '已读 = 0: 风险: 培训: 版本: 分类: 编号: 模块: 影响模块: 风险评估: 用户: 需求: 说明: 关联问题: 备注: 影响评估: 连接
                    Set rptItem = rptRcd.AddItem(""): rptItem.Focusable = False
                    If Val("" & mrsSheet(Excel_Col.对用户影响评估).Value) = 0 Then
                        rptItem.Icon = ICON_NoRead
                    Else
                        rptItem.Icon = ICON_Read
                    End If
                        
                    If "" & mrsSheet(Excel_Col.问题风险).Value = "高" Then
                        Set rptItem1 = rptRcd.AddItem(3)
                    ElseIf "" & mrsSheet(Excel_Col.问题风险).Value = "中" Then
                        Set rptItem1 = rptRcd.AddItem(2)
                    ElseIf "" & mrsSheet(Excel_Col.问题风险).Value = "低" Then
                        Set rptItem1 = rptRcd.AddItem(1)
                    Else
                        Set rptItem1 = rptRcd.AddItem(0)
                    End If
                    rptItem1.Caption = " ": rptItem1.Focusable = False

                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.发布版本).Value)): rptItem.Focusable = False
                    
                    Set rptItem = rptRcd.AddItem(CStr(Replace(varSheet(i), "$", ""))): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.问题编号).Value)): rptItem.Focusable = False
                   
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.登记模块).Value)): rptItem.Focusable = False
                    
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.影响模块).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.影响类型).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.风险评估说明).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.登记用户).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.用户需求).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.修改说明).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.相关问题).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & mrsSheet(Excel_Col.备注).Value)): rptItem.Focusable = False
                    
                    '---- 用户操作项
'                    Set rptItem2 = rptRcd.AddItem("")
                    If "" & mrsSheet(Excel_Col.是否需要培训) = "否" Then
'                         rptItem2.Icon=-1
                        Set rptItem2 = rptRcd.AddItem("无需培训") '无需培训
                        
                    Else
                        If "" & mrsSheet(Excel_Col.操作培训情况).Value = "" Then
'                            rptItem2.Icon = ICON_UnTrain
                            Set rptItem2 = rptRcd.AddItem("　")

                        Else
'                            rptItem2.Icon = ICON_Train
                            Set rptItem2 = rptRcd.AddItem("已培训")
       
                        End If
                    End If
                    
                    rptRcd.AddItem Val(CStr("" & mrsSheet(Excel_Col.对用户影响评估).Value))   '0-未填写 1-正面影响 2-负面影响 3-无影响
                    
                    '--- 超链接,是否修改
                    rptRcd.AddItem CStr("" & mrsSheet(Excel_Col.是否有HTML文档).Value)
                    rptRcd.AddItem CStr("0")
                    rptRcd.PreviewText = "" & mrsSheet(Excel_Col.用户需求).Value
                End If
                
            End With
            
            mrsSheet.MoveNext
        Loop
        If lngCount > 0 Then lblA.Caption = Trim(lblA.Caption & " " & Replace(varSheet(i), "$", "") & "(" & lngCount & ")")
    Next
    Set mrsSheet = Nothing
        
    '定位到上次选中行
    If mlngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.编号).Value) = mlngItemID Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    
    '展开选中行
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    
    rptList.Populate
    Call rptList_SelectionChanged '触发选择事件
    
    '保存查询参数
    m条件.风险等级 = cbo风险等级.List(cbo风险等级.ListIndex)
    m条件.是否阅读 = cbo是否阅读.List(cbo是否阅读.ListIndex)
    m条件.是否培训 = cbo培训.List(cbo培训.ListIndex)
    m条件.模块 = Replace(tvwLeft.SelectedItem.Text, "(新版)", "")
    If InStr(m条件.模块, "(") > 0 Then m条件.模块 = Mid(m条件.模块, 1, InStr(m条件.模块, "(") - 1)
    
    m条件.开始版本 = cbo开始版本.List(cbo开始版本.ListIndex)
    m条件.结束版本 = cbo结束版本.List(cbo结束版本.ListIndex)
    m条件.用户 = cbo用户.List(cbo用户.ListIndex)
    
    m条件.影响类型 = cbo影响类型.List(cbo影响类型.ListIndex)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "风险等级", m条件.风险等级)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "是否阅读", m条件.是否阅读)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "是否培训", m条件.是否培训)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "开始版本", m条件.开始版本)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "结束版本", m条件.结束版本)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "模块", m条件.模块)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "用户", m条件.用户)
    Call SaveSetting("ZLSOFT", "公共模块\UpgradeReader", "影响类型", m条件.影响类型)

'    If int条件 > 0 Then
'        '至少有一个原来保存的条件失效，需要调用刷新
'        Call ReLoad
'    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Number & " " & Err.Description, vbQuestion, gstrSysname
    
End Sub

Private Function LoadInitIcon()
    '装入ICON到 imgTree控件
    Dim intIcon As Integer
    Dim strIcon As String
    
    strIcon = ","
    With imgTree
        .ListImages.Clear
        .ImageHeight = 16
        .ImageWidth = 16
    End With
    
    For intIcon = 99 To 240
        imgTree.ListImages.Add , "K_" & intIcon, mclsAppTool.GetIcon(intIcon)
    Next
    Set Me.tvwLeft.ImageList = imgTree

End Function

Private Sub SaveItem()
'保存修改结果
    Dim str编号 As String, strSheet As String, str影响评估 As String, str培训 As String
    Dim i As Long
    With rptList
        For i = 0 To .Records.Count - 1
            
            If .Records(i).Item(mCol.修改).Value = "1" Or .Records(i).Item(mCol.培训).Value <> " " Then
                str编号 = .Records(i).Item(mCol.编号).Value
                strSheet = .Records(i).Item(mCol.分类).Value & "$"
                str影响评估 = Val(.Records(i).Item(mCol.影响评估).Value)
                str培训 = .Records(i).Item(mCol.培训).Value
                If str编号 <> "" Then
                
                    Set mrsSheet = OpenExcelSheet(strSheet)
                    mrsSheet.Filter = 0
                    
                    If mrsSheet.RecordCount > 0 Then mrsSheet.MoveFirst
                    
                    mrsSheet.Filter = "" & mrsSheet(Excel_Col.问题编号).Name & " = " & str编号
                    Do Until mrsSheet.EOF
                        mrsSheet(Excel_Col.对用户影响评估) = Switch(str影响评估 = 0, "0-未填写", str影响评估 = 1, "1-正面作用", str影响评估 = 2, "2-负面作用", str影响评估 = 3, "3-无影响")
                        mrsSheet(Excel_Col.阅读记录) = gstrDBUser
                        If mrsSheet(Excel_Col.是否需要培训) <> "否" And str培训 <> "无需培训" Then
                            mrsSheet(Excel_Col.操作培训情况) = IIf(str培训 = "已培训", "已培训", "")
                        End If
                        mrsSheet.Update
                        mrsSheet.MoveNext
                    Loop
                    
                    Set mrsSheet = Nothing
                End If
            End If
        Next
    End With
    
End Sub

Private Function IsAdd(ByVal bln包含模块条件 As Boolean) As Boolean
    '   是否符合条件，符合返回true
    Dim str模块 As String, str开始版本 As String, str结束版本 As String, strTmp As String
    Dim str用户 As String, str影响模块 As String
    Dim str风险等级 As String, str是否阅读 As String, str是否培训 As String, str影响类型 As String
    IsAdd = False
        
    '
    If Val("" & mrsSheet(Excel_Col.问题编号).Value) = 0 Then Exit Function
    If Trim("" & mrsSheet(Excel_Col.发布版本).Value) = "" Then Exit Function
    '0-未登录 显示所有用户
    '1-已登录 只能当前用户。
'    If mIntType = 1 And gstr用户单位名称 <> "" Then
'        If "" & mrsSheet(Excel_Col.登记用户) <> gstr用户单位名称 Then Exit Function
'    ElseIf mIntType = 0 Then
        
        If mstrFileName <> mLastFileName Then
            '初次调用，取注册表中的条件
            str用户 = m条件.用户
        Else
            If cbo用户.ListIndex >= 0 Then str用户 = cbo用户.List(cbo用户.ListIndex)
        End If

        
        If str用户 <> "所有" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.登记用户))
            If str用户 <> strTmp Then Exit Function
        End If
'    End If
    
    '-- 模块 ,    '-- 影响模块

    If bln包含模块条件 Then '添加模块个数时，不区分模块
         
        If mLastFileName <> mstrFileName Then
            str模块 = m条件.模块
        Else
            If Not tvwLeft.SelectedItem Is Nothing Then
                If tvwLeft.SelectedItem.Key <> "Root" Then
                    str模块 = tvwLeft.SelectedItem.Text
                Else
                    str模块 = "所有模块"
                End If
            Else
                str模块 = "所有模块"
            End If
        End If
        
        If InStr(str模块, "所有模块") <= 0 Then
            str模块 = Trim(Replace(str模块, "(新版)", ""))
            If InStr(str模块, "(") > 0 Then str模块 = Mid(str模块, 1, InStr(str模块, "(") - 1)
            strTmp = Trim(Replace("" & mrsSheet(Excel_Col.登记模块), "(新版)", ""))
            str影响模块 = Trim(Replace("" & mrsSheet(Excel_Col.影响模块), "(新版)", ""))
            
            If strTmp <> str模块 Then
                If str影响模块 <> "" Then
                    If InStr(str影响模块, str模块) <= 0 Then Exit Function
                Else
                    Exit Function
                End If
            End If
        End If
    End If
    
    '-- 版本
    
        If mLastFileName <> mstrFileName Then
            str开始版本 = m条件.开始版本
            str结束版本 = m条件.结束版本
        Else
            If cbo开始版本.ListIndex >= 0 And cbo结束版本.ListIndex >= 0 Then
                str开始版本 = cbo开始版本.List(cbo开始版本.ListIndex)
                str结束版本 = cbo结束版本.List(cbo结束版本.ListIndex)
            Else
                str开始版本 = "10.19.0"
                str结束版本 = "90.19.0"
            End If
        End If
        strTmp = Trim("" & mrsSheet(Excel_Col.发布版本))

        If strTmp = "" Then Exit Function
        If strTmp < str开始版本 Or strTmp > str结束版本 Then Exit Function
    
    
    '-- 风险等级
    If cbo风险等级.ListIndex >= 0 Then
        str风险等级 = cbo风险等级.List(cbo风险等级.ListIndex)
        If str风险等级 <> "所有" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.问题风险))
            If strTmp <> str风险等级 Then Exit Function
        End If
    End If
    
    '-- 是否阅读
    If cbo是否阅读.ListIndex >= 0 Then
        str是否阅读 = cbo是否阅读.List(cbo是否阅读.ListIndex)
        If str是否阅读 <> "所有" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.对用户影响评估))
            If str是否阅读 = "已阅" Then
                If Val(strTmp) = 0 Then Exit Function
            Else
                If Val(strTmp) <> 0 Then Exit Function
            End If
        End If
    End If
    
    '-- 是否培训
    If cbo培训.ListIndex >= 0 Then
        str是否培训 = cbo培训.List(cbo培训.ListIndex)
        If str是否培训 <> "所有" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.是否需要培训))
            
            If strTmp = "是" Then
                strTmp = "" & mrsSheet(Excel_Col.操作培训情况)
            Else
                strTmp = "无需培训"
            End If
            
            If strTmp = "" Then strTmp = "　" '未填就是未培训
            If InStr(strTmp, str是否培训) <= 0 Then Exit Function
        End If
    End If
    
    '-- 影响类型
    If cbo影响类型.ListIndex >= 0 Then
        str影响类型 = cbo影响类型.List(cbo影响类型.ListIndex)
        If str影响类型 <> "所有" Then
            strTmp = Trim("" & mrsSheet(Excel_Col.影响类型))
            If strTmp <> str影响类型 Then Exit Function
        End If
    End If

    If mLastFileName <> mstrFileName Then mLastFileName = mstrFileName
    IsAdd = True
    
End Function

Private Sub AddExcelToTree(ByVal strNodeKey As String)

    '添加Excel中的模块到模块列表中
    Dim str编号 As String, str模块 As String, blnAdd As Boolean, IntCount As Integer
    Dim objNote As Node, strNoteTxt As String
    str模块 = Replace(Trim("" & mrsSheet(Excel_Col.登记模块).Value), "(新版)", "")
    str编号 = Val(Trim("" & mrsSheet(Excel_Col.问题编号)))
    
    If str模块 = "" Or Val(str编号) = 0 Then Exit Sub
    With tvwLeft
        If tvwLeft.Nodes.Count > 0 Then
            blnAdd = True
            For IntCount = 1 To tvwLeft.Nodes.Count
                strNoteTxt = tvwLeft.Nodes(IntCount).Text
                
                If InStr(strNoteTxt, "(") > 0 Then strNoteTxt = Mid(strNoteTxt, 1, InStr(strNoteTxt, "(") - 1)
                If str模块 = strNoteTxt Then
                    blnAdd = False
                    Exit For
                End If
            Next
            
            If blnAdd Then
                If strNodeKey = "" Then
                    Set objNote = tvwLeft.Nodes.Add(, , "E" & Val(str编号), str模块, "K_" & 170)
                Else
                    Set objNote = tvwLeft.Nodes.Add(strNodeKey, 4, "E" & Val(str编号), str模块, "K_" & 170)
                End If
            End If
        End If
    End With
End Sub

Private Sub ReLoad()
    '刷新
    
    If mblnEdit = True Then
        If MsgBox("所做的修改未保存，是否继续？", vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
        mblnEdit = False
    End If
    'Call initTree
    Dim objNode As Node
    If Not tvwLeft.SelectedItem Is Nothing Then
        Set objNode = tvwLeft.SelectedItem
        Call tvwLeft_NodeClick(objNode)
    End If
End Sub
Private Sub AddParent(ByVal intIndex As Integer)
    '更新父节点的模块个数
    Dim strTxt As String
    If tvwLeft.Nodes(intIndex).Parent Is Nothing Then Exit Sub
    strTxt = tvwLeft.Nodes(intIndex).Parent.Text
    If InStr(strTxt, "(") > 0 Then strTxt = Mid(strTxt, 1, InStr(strTxt, "(") - 1)
    tvwLeft.Nodes(intIndex).Parent.Tag = Val(tvwLeft.Nodes(intIndex).Parent.Tag) + 1
    tvwLeft.Nodes(intIndex).Parent.Text = strTxt & "(" & Val(tvwLeft.Nodes(intIndex).Parent.Tag) & ")"
    Call AddParent(tvwLeft.Nodes(intIndex).Parent.Index)
End Sub

Private Sub addModleToTree(ByVal str上级模块 As String, ByVal strKey As String)
    '加入模块到列表
    Dim str模块 As String, var模块 As Variant, int模块 As Integer
    Dim objNode As Node
    If str上级模块 = "" Then Exit Sub
    
    str模块 = Trim(ReadFromIni(App.Path & "\" & App.EXEName & ".ini", str上级模块, "模块"))
    var模块 = Split(str模块, "|")
    For int模块 = LBound(var模块) To UBound(var模块)
        If var模块(int模块) <> "" Then
            Set objNode = Me.tvwLeft.Nodes.Add(strKey, 4, strKey & Format(int模块 + 1, "000"), var模块(int模块), "K_" & 99)
            Call addModleToTree(var模块(int模块), strKey & Format(int模块 + 1, "000"))
        End If
    Next
End Sub

