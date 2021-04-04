VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoleGrant 
   Caption         =   "角色授权"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14025
   Icon            =   "frmRoleGrant.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14025
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame fraHSplit 
      Height          =   30
      Left            =   240
      MousePointer    =   7  'Size N S
      TabIndex        =   31
      Top             =   4800
      Width           =   9615
   End
   Begin VB.Frame fraVSplit 
      Height          =   7095
      Left            =   6360
      MousePointer    =   9  'Size W E
      TabIndex        =   18
      Top             =   480
      Width           =   30
   End
   Begin VB.CommandButton cmdUnSel 
      Caption         =   "全消(&R)"
      Height          =   350
      Left            =   3960
      TabIndex        =   30
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   3000
      TabIndex        =   29
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "全部展开(&D)"
      Height          =   350
      Left            =   1680
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "检查对象权限(&V)"
      Height          =   350
      Left            =   9240
      TabIndex        =   27
      Top             =   7680
      Width           =   1695
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Index           =   1
      Left            =   6480
      ScaleHeight     =   6975
      ScaleWidth      =   5175
      TabIndex        =   19
      Top             =   600
      Width           =   5175
      Begin MSComctlLib.TreeView tvwMenu 
         Height          =   3690
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   480
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   6509
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgTreeview"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView lvwFunc 
         Height          =   2415
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   4560
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   4260
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15724768
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "功能"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "排列"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "说明"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwModRelas 
         Height          =   4890
         Left            =   3600
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   8625
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgTreeview"
         Appearance      =   1
      End
      Begin VB.Label lblNotice 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         Height          =   180
         Left            =   1080
         TabIndex        =   26
         Top             =   180
         Width           =   90
      End
      Begin VB.Label lblFuncNote 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         Caption         =   "关联模块功能"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   4320
         Width           =   1080
      End
      Begin VB.Label lblMenuNote 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         Caption         =   "关联模块"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   7155
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   4035
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   14
      Top             =   7680
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -60
      TabIndex        =   11
      Top             =   585
      Width           =   11070
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   12750
      TabIndex        =   13
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   11520
      TabIndex        =   12
      Top             =   7680
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgTreeview 
      Left            =   5280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":000C
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":0E5E
            Key             =   "分类"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":76C0
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":859A
            Key             =   "分类_选中"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":EDFC
            Key             =   "Function"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":FC4E
            Key             =   "Optional"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":164B0
            Key             =   "Fixed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Top             =   8108
      Visible         =   0   'False
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   8064
      Width           =   14028
      _ExtentX        =   24739
      _ExtentY        =   635
      SimpleText      =   $"frmRoleGrant.frx":1CD12
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRoleGrant.frx":1CD59
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21828
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
   Begin MSComctlLib.ListView lvwTmp 
      Height          =   495
      Left            =   3720
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      View            =   1
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "功能"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "排列"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "说明"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Index           =   0
      Left            =   240
      ScaleHeight     =   6975
      ScaleWidth      =   6135
      TabIndex        =   17
      Top             =   600
      Width           =   6135
      Begin VB.CheckBox chkShowDisReport 
         Caption         =   "显示停用报表(&R)"
         Height          =   345
         Left            =   4800
         TabIndex        =   32
         Top             =   105
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.CheckBox chkOnlyShow 
         Caption         =   "仅已授权(&G)"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   150
         Width           =   1290
      End
      Begin VB.CheckBox chkVirtual 
         Caption         =   "含虚拟模块(&M)"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   150
         Width           =   1545
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Top             =   120
         Width           =   1530
      End
      Begin MSComctlLib.TreeView tvwMenu 
         Height          =   3690
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   6509
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgTreeview"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvwFunc 
         Height          =   2415
         Index           =   0
         Left            =   15
         TabIndex        =   8
         Top             =   4560
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   4260
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "功能"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "排列"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "说明"
            Object.Width           =   12347
         EndProperty
      End
      Begin VB.Label lblMenuNote 
         AutoSize        =   -1  'True
         Caption         =   "模块菜单"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblFuncNote 
         AutoSize        =   -1  'True
         Caption         =   "模块功能"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   4320
         Width           =   720
      End
      Begin VB.Label lblSearch 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "定位(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1155
         TabIndex        =   3
         Top             =   180
         Width           =   630
      End
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      Caption         =   "应用系统(&S)"
      Height          =   180
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin VB.Image imgRoleGrant 
      Height          =   480
      Left            =   120
      Picture         =   "frmRoleGrant.frx":1D5ED
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "对角色“门诊收费员”的授权处理。(Ctrl+F查找,F3下一个)"
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuState 
         Caption         =   "功能项目(&1)"
         Index           =   0
      End
      Begin VB.Menu mnuPopuState 
         Caption         =   "功能详细(&2)"
         Checked         =   -1  'True
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRoleGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'树形菜单枚举
Private Enum MenuType
    MT_模块 = 0
    MT_关联模块 = 1
End Enum
'界面样式
Private Enum ShowFaceType
    SFT_应用系统 = 0 '应用系统的样式
    SFT_自定义报表 = 1 '基础工具以及自定义报表
    SFT_其他 = 2 '基础编码与基础函数
End Enum
'读取数据类型
Private Enum ReadDataType
    RDT_Menu = 0
    RDT_Module = 1
    RDT_Function = 2
    RDT_Table = 3
    RDT_Systems = 4
    RDT_ModRelas = 5
End Enum
'清空控件类型
Private Enum ClearType
    CT_关联功能 = 0 '切换关联模块，需要清空关联功能
    CT_关联模块 = 1 '切换功能以模块与功能间的切换需要清空关联模块
    CT_功能 = 2 '模块切换
    CT_Sys = 3 '系统切换
End Enum

Private Const BLN_TEST = False
Private mblnOk As Boolean
Private mstrRole As String

Private WithEvents mclsPrivilege As clsPrivilege
Attribute mclsPrivilege.VB_VarHelpID = -1

Private mrsTree As ADODB.Recordset        '模块菜单
Private mrsModule As ADODB.Recordset      '模块授予情况记录
'模块授权情况的总体情况，以及模块授权操作的缓存，在保存授权时会将信息更新到mrsModule中
Private mrsModsInfo As ADODB.Recordset
Private mrsTable  As ADODB.Recordset      '编码表授予情况记录
Private mrsFunction  As ADODB.Recordset   '函数授予情况记录
Private mrsSys   As ADODB.Recordset   '所有安装系统
Private mrsModRelas As ADODB.Recordset   '模块关系数据
Private mrsRelasTree As ADODB.Recordset  '关联模块树形

Private mrsRelas As ADODB.Recordset '权限关系(主从)
Private mrsRelExcl As ADODB.Recordset ' 权限关系(互斥)
Private mrsGroup As ADODB.Recordset '权限分组
Private mblnVirtual As Boolean

Private mintUpdate As Integer
Private mblnItem As Boolean
'快捷键控制变量
Private mblnExpanded As Boolean '是否全部展开
'其他变量
Private mlngSys As Long
Private msftStyle As ShowFaceType  '界面样式
Private mcllHaveSys As Collection '建立模块菜单关系的系统,key=系统,值=1
Private mcllKeyModule As Collection '模块菜单关系，Key=系统_模块,值=菜单Key1,菜单Key2...
Private mcllCodeModule As Collection '编号，模块关系，Key=编号,值=系统_模块
Private mstrFind As String '查找字符串
Private mlngCurPos As Long '当前查找位置
Private mblnClear As Boolean
Private mcllTip As Collection '悬浮提示
Private mstrCurRelas As String
Private mblnReturn As Boolean
Private mintActive As Integer
Private mblnUnRefresh As Boolean
Private mblnSaveClick As Boolean '记录是否保存修改

Private mrsRptGroups As ADODB.Recordset      '记录报表分组数据
Private mrsReports As ADODB.Recordset        '记录报表数据
Private mrsGroups As ADODB.Recordset         '记录报表组数据

Public Function GrantToRole(ByVal strRole As String) As Boolean
    mblnOk = False
    If gstr注册码 = "" Then
        MsgBox "该用户的注册码无效，请重新注册。", vbExclamation, gstrSysName
        Exit Function
    End If
    mstrRole = strRole
    Set mcllHaveSys = New Collection
    Set mcllKeyModule = New Collection
    Set mcllCodeModule = New Collection
    Set mrsModule = ReadData(RDT_Module)
    If mrsModule.RecordCount = 0 Then
        Set mrsFunction = ReadData(RDT_Function)
        If mrsFunction.RecordCount = 0 Then
            Set mrsTable = ReadData(RDT_Table)
            If mrsTable.RecordCount = 0 Then
                MsgBox "你不具有可授予的权力，不能进行授权操作。", vbInformation, gstrSysName
                On Error Resume Next
                Unload Me
                err.Clear: On Error GoTo 0
                Exit Function
            End If
        End If
    End If
    Me.Show vbModal, frmMDIMain
    GrantToRole = mblnSaveClick
    mblnSaveClick = False
End Function

Private Sub chkOnlyShow_Click()
    
    LockWindowUpdate Me.hwnd
    If msftStyle <> SFT_其他 Then Call AdjustRelasTree
    mlngCurPos = 0
    Call SetOnlyShow
    If tvwMenu(MT_模块).Nodes.Count = 0 Then
        Call ClearFace(CT_Sys)
    End If
    LockWindowUpdate 0
End Sub

Private Sub chkShowDisReport_Click()
    Dim objNode As Node
    
    LockWindowUpdate Me.hwnd
    mrsModsInfo.Filter = "系统=0 And 序号 >=100"
    mrsModsInfo.Sort = "序号"
    tvwMenu(MT_模块).Nodes.Clear
    tvwMenu(MT_模块).Nodes.Add , , "K_0", "所有报表", "分类", "分类_选中"
    If mrsGroups.RecordCount <> 0 Then mrsGroups.MoveFirst
    With mrsGroups
        Do While Not .EOF
            If IsNull(!上级id) = True Then
                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_0", tvwChild, "K_" & !Id, !名称, "分类", "分类_选中")
            Else
                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & !上级id, tvwChild, "K_" & !Id, !名称, "分类", "分类_选中")
            End If
            .MoveNext
        Loop
    End With
    With mrsModsInfo
        Do While Not .EOF
            mrsRptGroups.Filter = IIf(chkShowDisReport.value, "", "是否停用 = 0 and ") & "程序id = " & !序号
            mrsReports.Filter = IIf(chkShowDisReport.value, "", "是否停用 = 0 and ") & "程序id = " & !序号
            If mrsRptGroups.RecordCount = 1 Then
                '该模块为报表组发布的
                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & Nvl(mrsRptGroups!分类id, 0), tvwChild, "M_0_" & !序号, "【" & Format(!序号, "000000") & "】" & !标题, "Module")
                objNode.Checked = !授权否 = 1
                mrsModsInfo.Update "模块类型", 2
            ElseIf mrsReports.RecordCount = 1 Then
                '该模块为报表发布的
                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & Nvl(mrsReports!分类id, 0), tvwChild, "M_0_" & !序号, "【" & Format(!序号, "000000") & "】" & !标题, "Module")
                objNode.Checked = !授权否 = 1
                mrsModsInfo.Update "模块类型", 2
            End If
            .MoveNext
        Loop
    End With

    '删除没有子项的文件夹
    Call DeleteNodes
    LockWindowUpdate 0
End Sub

Private Sub chkVirtual_Click()
    Dim objNode As Node
    mlngCurPos = 0
    If Not mblnVirtual Then Call SetVirtualVisual(chkVirtual.value <> 0)
    '处理，这种情况，先勾选虚拟模块，选中虚拟模块，再取消勾选，选中存在的模块，
    '由于要需找以前选中的节点（现在已经不存在），取消以前选择节点的加粗状态(这样处理会报错），所以这样处理
    On Error Resume Next
    Set objNode = tvwMenu(MT_模块).Nodes(tvwMenu(MT_模块).Tag)
    If err.Number <> 0 Then
        err.Clear: tvwMenu(MT_模块).Tag = ""
        If tvwMenu(MT_模块).Nodes.Count <> 0 Then
            Call tvwMenu_NodeClick(MT_模块, tvwMenu(MT_模块).Nodes(1))
        End If
    End If
End Sub

Private Sub cmbSystem_Click()
    Dim strPre As String
    Dim objNode As Node, objNodeChild As Node
    Dim blnHaveSys As Boolean
    Dim strTMp As String
    Dim strFirstNode As String
    Dim strPreKey As String
    Dim i As Long
    
    mblnUnRefresh = True
    '判断是否切换系统，切换了才刷新数据
    If Val(cmbSystem.Tag) <> cmbSystem.ListIndex Then
        LockWindowUpdate Me.hwnd
        Call ClearFace(CT_Sys)
        chkShowDisReport.Visible = False
        mlngSys = Val(cmbSystem.ItemData(cmbSystem.ListIndex))
        cmbSystem.Tag = cmbSystem.ListIndex
        If mlngSys = 0 Then '非应用系统授权
            chkVirtual.Visible = False
            Select Case cmbSystem.Text
                Case "基础编码"
                    msftStyle = SFT_其他
                    If mrsTable Is Nothing Then Set mrsTable = ReadData(RDT_Table)
                    If glngSysNo <> -1 Then
                        mrsTable.Filter = "系统 = " & glngSysNo
                    Else
                        mrsTable.Filter = ""
                    End If
                    mrsTable.Sort = "系统,表名"
                    With mrsTable
                        Do While Not .EOF
                            If strPre <> !系统 & "" Then
                                Set objNode = tvwMenu(MT_模块).Nodes.Add(, , "K_" & !系统, "【" & !系统 & "】" & !系统名, "分类", "分类_选中")
                                strPre = !系统 & "": objNode.Checked = True
                            End If
                            Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & !系统, tvwChild, "T_" & !系统 & "_" & !表名, RPAD("【" & !表名 & "】", 20) & !说明, "Table")
                            objNode.Checked = !授权否 = 1
                            '默认勾选，若有一个子级不勾选，则取消勾选
                            If Not objNode.Checked Then objNode.Parent.Checked = False
                            '不默认勾选节点，防止速度较慢
'                            '获取第一个节点与第一个勾选节点
'                            If objNode.Checked And tvwMenu(MT_模块).Tag = "" Then
'                                tvwMenu(MT_模块).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            .MoveNext
                        Loop
                    End With
                Case "取数函数"
                    msftStyle = SFT_其他
                    If mrsFunction Is Nothing Then Set mrsFunction = ReadData(RDT_Function)
                    If glngSysNo <> -1 Then
                        mrsFunction.Filter = "系统 = " & glngSysNo
                    Else
                        mrsFunction.Filter = ""
                    End If
                    mrsFunction.Sort = "系统,函数名"
                    With mrsFunction
                        Do While Not .EOF
                            If strPre <> !系统 & "" Then
                                Set objNode = tvwMenu(MT_模块).Nodes.Add(, , "K_" & !系统, "【" & !系统 & "】" & !系统名, "分类", "分类_选中")
                                strPre = !系统 & "": objNode.Checked = True
                            End If
                            Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & !系统, tvwChild, "F_" & !系统 & "_" & !函数名, RPAD("【" & !函数名 & "（中文名：" & !中文名 & "）】", 52) & !说明, "Function")
                            objNode.Checked = !授权否 = 1
                            '不默认勾选节点，防止速度较慢
                            '默认勾选，若有一个子级不勾选，则取消勾选
                            If Not objNode.Checked Then objNode.Parent.Checked = False
'                            '获取第一个节点与第一个勾选节点
'                            If objNode.Checked And tvwMenu(MT_模块).Tag = "" Then
'                                tvwMenu(MT_模块).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            .MoveNext
                        Loop
                    End With
                Case "基础工具"
                    If mrsModsInfo Is Nothing Then Set mrsModsInfo = GetModuleInfo
                    mrsModsInfo.Filter = "系统=0 And 序号<100"
                    mrsModsInfo.Sort = "序号"
                    With mrsModsInfo
                        Do While Not .EOF
                            Set objNode = tvwMenu(MT_模块).Nodes.Add(, , "M_0_" & !序号, "【" & Format(!序号, "000000") & "】" & !标题, "Module")
                            objNode.Checked = !授权否 = 1
                            '标记模块类型
                            mrsModsInfo.Update "模块类型", 2
                            '不默认勾选节点，防止速度较慢
'                            '获取第一个节点与第一个勾选节点
'                            If objNode.Checked And tvwMenu(MT_模块).Tag = "" Then
'                                tvwMenu(MT_模块).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            .MoveNext
                        Loop
                    End With
                Case "自定义报表"
                    msftStyle = SFT_自定义报表
                    chkShowDisReport.Visible = True
                    On Error GoTo errH
                    If mrsModsInfo Is Nothing Then Set mrsModsInfo = GetModuleInfo
                    mrsModsInfo.Filter = "系统=0 And 序号 >=100"
                    mrsModsInfo.Sort = "序号"
                    
                    gstrSQL = "Select Id, 上级id, 名称, 说明" & vbNewLine & _
                                "From (Select Id, 上级id, 名称, 说明 From Zlrptclasses)" & vbNewLine & _
                                "Start With 上级id Is Null" & vbNewLine & _
                                "Connect By Prior Id = 上级id"
                    Set mrsGroups = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "查找分组")
                    gstrSQL = "Select Id, 程序id, 名称, 分类id, Nvl(是否停用, 0) 是否停用 From Zlrptgroups Where 系统 Is Null And 程序id Is Not Null"
                    Set mrsRptGroups = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "查找已发布的分组")
                    gstrSQL = "Select 程序id, 分类id, Nvl(是否停用, 0) 是否停用  From Zlreports Where 系统 Is Null And 程序id Is Not Null"
                    Set mrsReports = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "查找已发布的报表")
                    
                    tvwMenu(MT_模块).Nodes.Add , , "K_0", "所有报表", "分类", "分类_选中"
                    
                    With mrsGroups
                        Do While Not .EOF
                            If IsNull(!上级id) = True Then
                                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_0", tvwChild, "K_" & !Id, !名称, "分类", "分类_选中")
                            Else
                                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & !上级id, tvwChild, "K_" & !Id, !名称, "分类", "分类_选中")
                            End If
                            .MoveNext
                        Loop
                    End With
                    With mrsModsInfo
                        Do While Not .EOF
                            mrsRptGroups.Filter = "是否停用 = 0 and 程序id = " & !序号
                            mrsReports.Filter = "是否停用 = 0 and 程序id = " & !序号
                            If mrsRptGroups.RecordCount = 1 Then
                                '该模块为报表组发布的
                                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & Nvl(mrsRptGroups!分类id, 0), tvwChild, "M_0_" & !序号, "【" & Format(!序号, "000000") & "】" & !标题, "Module")
                                objNode.Checked = !授权否 = 1
                                '标记模块类型
                                mrsModsInfo.Update "模块类型", 2
                            ElseIf mrsReports.RecordCount = 1 Then
                                '该模块为报表发布的
                                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & Nvl(mrsReports!分类id, 0), tvwChild, "M_0_" & !序号, "【" & Format(!序号, "000000") & "】" & !标题, "Module")
                                objNode.Checked = !授权否 = 1
                                '标记模块类型
                                mrsModsInfo.Update "模块类型", 2
                            End If
                            .MoveNext
                        Loop
                    End With
                    
                    '删除没有子项的文件夹
                    Call DeleteNodes
            End Select
        Else '应用系统授权
            msftStyle = SFT_应用系统
            '判断节点模块关系是否存储
            On Error Resume Next
            strTMp = mcllHaveSys("S_" & mlngSys)
            If err.Number <> 0 Then err.Clear: strTMp = ""
            blnHaveSys = strTMp <> ""
            On Error GoTo errH
            If mrsModsInfo Is Nothing Then Set mrsModsInfo = GetModuleInfo
            If mrsTree Is Nothing Then Set mrsTree = ReadData(RDT_Menu)
            With mrsTree
                .Filter = "系统=" & mlngSys
                Do While Not .EOF
                    If !模块 = 0 Then
                        If !上级 = 0 Then
                            Set objNode = tvwMenu(MT_模块).Nodes.Add(, , "K_" & Format(!编号, "000000"), !标题, "分类", "分类_选中")
                        Else
                            Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & Format(!上级, "000000"), tvwChild, "K_" & Format(!编号, "000000"), !标题, "分类", "分类_选中")
                        End If
                    Else
                        mrsModsInfo.Filter = "系统=" & mlngSys & " And 序号=" & !模块
                        If Not mrsModsInfo.EOF Then '没有找到对应模块则不显示
                            If !上级 = 0 Then
                                Set objNode = tvwMenu(MT_模块).Nodes.Add(, , "M_" & Format(!编号, "000000") & "_" & !模块, "【" & Format(!模块, "000000") & "】" & !标题, "Module")
                                If Not blnHaveSys Then mcllCodeModule.Add objNode.Key, "K_" & !编号
                            Else
                                '增加对模块的下级节点同样是模块的支持
                                On Error Resume Next
                                strPreKey = mcllCodeModule("K_" & !上级)
                                If err.Number <> 0 Then
                                    err.Clear
                                    Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & Format(!上级, "000000"), tvwChild, "M_" & Format(!编号, "000000") & "_" & !模块, "【" & Format(!模块, "000000") & "】" & !标题, "Module")
                                Else
                                    Set objNode = tvwMenu(MT_模块).Nodes.Add(strPreKey, tvwChild, "M_" & Format(!编号, "000000") & "_" & !模块, "【" & Format(!模块, "000000") & "】" & !标题, "Module")
                                End If
                                On Error GoTo errH
                                If Not blnHaveSys Then mcllCodeModule.Add objNode.Key, "K_" & !编号
                            End If
                            
                            objNode.Checked = mrsModsInfo!授权否 = 1
                            '不默认勾选节点，防止速度较慢
'                            '获取第一个节点与第一个勾选节点
'                            If objNode.Checked And tvwMenu(MT_模块).Tag = "" Then
'                                tvwMenu(MT_模块).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            '标记模块类型，存储节点模块关系
                            If Not blnHaveSys Then
                                mrsModsInfo.Update "模块类型", 1
                                On Error Resume Next
                                strTMp = mcllKeyModule("K_" & mlngSys & "_" & !模块)
                                If err.Number <> 0 Then
                                    err.Clear: strTMp = ""
                                    strTMp = objNode.Key
                                Else
                                    mcllKeyModule.Remove "K_" & mlngSys & "_" & !模块
                                    strTMp = strTMp & "," & objNode.Key
                                End If
                                mcllKeyModule.Add strTMp, "K_" & mlngSys & "_" & !模块
                                On Error GoTo errH
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End With
            '判断是否展示虚拟模块，获取虚拟模块的第一个节点，与第一个勾选的节点
            Call SetVirtualVisual(tvwMenu(MT_模块).Nodes.Count = 0 Or chkVirtual.value <> 0, strFirstNode)
            '标记该系统的模块菜单关系已经记录
            If Not blnHaveSys Then
                mcllHaveSys.Add "1", "S_" & mlngSys
            End If
            '设置整体勾选状态
            Call CheckNode(tvwMenu(MT_模块))
        End If
        '不默认勾选节点，防止速度较慢
'        '没有勾选节点且有一个第一节点，则选择第一节点
'        If tvwMenu(MT_模块).Tag = "" And strFirstNode <> "" Then
'            tvwMenu(MT_模块).Tag = strFirstNode
'        End If
'        '展开标记节点
'        If tvwMenu(MT_模块).Tag <> "" Then
'            strTmp = tvwMenu(MT_模块).Tag: tvwMenu(MT_模块).Tag = ""
'            Call SetNodeExpand(tvwMenu(MT_模块), strTmp) '展开节点
'            Call tvwMenu_NodeClick(MT_模块, tvwMenu(MT_模块).Nodes(strTmp))
'        End If
        If chkOnlyShow.value = 1 Then
            Call SetOnlyShow
        End If
        Call Form_Resize
        LockWindowUpdate 0
        mblnUnRefresh = False
        Call RefreshState
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "cmbSystem_Click:" & err.Description, vbInformation, Me.Caption
End Sub

Private Function DeleteNodes(Optional ByVal strKey As String) As Boolean
'删除tvwMenu(MT_模块)指定节点下所有空节点
'入参:tvwMenu(MT_模块)指定节点的key值，如果为空，则默认为根节点
    Dim objNode As Node
    Dim strDelKeys As String
    Dim arrTmp As Variant, i As Long
    Dim blnNotDel As Boolean

    '获取初始节点，以便循环
    If tvwMenu(MT_模块).Nodes.Count = 0 Then Exit Function
    If strKey = "" Then
        Set objNode = tvwMenu(MT_模块).Nodes(1)
    ElseIf tvwMenu(MT_模块).Nodes(strKey).Children <> 0 Then
        Set objNode = tvwMenu(MT_模块).Nodes(strKey).Child
    End If
    '获取可以删除的节点
    Do While Not objNode Is Nothing
        '若子级被选中，则父级选中
        If objNode.Key Like "M*" Then
            blnNotDel = True
        ElseIf Not DeleteNodes(objNode.Key) Then
            strDelKeys = strDelKeys & "|" & objNode.Key
        Else
            blnNotDel = True
        End If
        Set objNode = objNode.Next
    Loop
    
    '删除可以删除的节点
    arrTmp = Split(Mid(strDelKeys, 2), "|")
    For i = LBound(arrTmp) To UBound(arrTmp)
        tvwMenu(MT_模块).Nodes.Remove arrTmp(i)
    Next
    DeleteNodes = blnNotDel
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    frmModuleCheck.ShowMe IIf(glngSysNo = -1, 0, glngSysNo)
End Sub

Private Sub cmdExp_Click()
    Call Form_KeyDown(vbKeyD, vbCtrlMask)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "ZL9Svrtools\" & Me.name
End Sub

Private Sub cmdOK_Click()
    'Dim sglTimer0 As Single
    Dim str所有者() As String, i As Long
    
    mblnSaveClick = True
    mrsSys.Filter = "": mrsSys.Sort = "编号"
    ReDim str所有者(mrsSys.RecordCount - 1)
    For i = LBound(str所有者) To UBound(str所有者)
        str所有者(i) = mrsSys!所有者 & ""
        mrsSys.MoveNext
    Next
    MousePointer = 13
    'sglTimer0 = Timer
    If mrsTable Is Nothing Then Set mrsTable = ReadData(RDT_Table)
    If mrsFunction Is Nothing Then Set mrsFunction = ReadData(RDT_Function)
    If mrsRelasTree Is Nothing Then Call GetRelasTree(True)
    '缓存界面值，并拆解模块授权的整体情况
    '#将数据更新到缓存
    Call UpdateGrantState
    pgb.Visible = True
    Set mclsPrivilege = New clsPrivilege
    Call mclsPrivilege.InitOracle(gcnOracle)
    If mclsPrivilege.InitPrivilege(str所有者, mstrRole, mrsModule, mrsTable, mrsFunction) Then
        '一、清除**********************************************************************************************************
        If mclsPrivilege.RevokePrivilege Then
        End If
        '二、授权**********************************************************************************************************
        If mclsPrivilege.GrantPrivilege Then
            '插入重要操作日志
            Call SaveAuditLog(2, "角色授权", "修改角色“" & Split(mstrRole, "_")(1) & "”的权限")
        End If
    End If
'    MsgBox Timer - sglTimer0
    MousePointer = 0
    mblnOk = True
    If mclsPrivilege.FailInfo <> "" Then
        MsgBox "由于授权模块对象不存在或权限错误等原因，" & vbCr & "以下权限未正常授予：" & mclsPrivilege.FailInfo, vbExclamation, gstrSysName
    End If
    Set mclsPrivilege = Nothing
    Unload Me
End Sub

Private Sub cmdSelAll_Click()
    Dim objCur As Object
    If mintActive > 1 Then
        Set objCur = lvwFunc(mintActive Mod 2)
    Else
        Set objCur = tvwMenu(mintActive Mod 2)
    End If
    Call SetSel(objCur, True)
End Sub

Private Sub cmdUnSel_Click()
    Dim objCur As Object
    If mintActive > 1 Then
        Set objCur = lvwFunc(mintActive Mod 2)
    Else
        Set objCur = tvwMenu(mintActive Mod 2)
    End If
    Call SetSel(objCur, False)
End Sub

Private Sub Form_Activate()
    Dim strSql As String
    Dim lStyle As Long
    
    Call ApplyOEM(stbThis)
    lblNote.Caption = "对角色“" & Mid(mstrRole, 4) & "”的授权处理"
    cmbSystem.Tag = "-1" '方便处理
    Call SendMessage(tvwMenu(MT_关联模块).hwnd, TVM_SETBKCOLOR, 0, ByVal &HEFF0E0)
    lStyle = GetWindowLong(tvwMenu(MT_关联模块).hwnd, GWL_STYLE)
    Call SetWindowLong(tvwMenu(MT_关联模块).hwnd, GWL_STYLE, lStyle - TVS_HASLINES)
    Call SetWindowLong(tvwMenu(MT_关联模块).hwnd, GWL_STYLE, lStyle)
    mblnExpanded = False
    If mrsRelas Is Nothing Then
    '--- 初始化权限关系记录集
        strSql = "Select 系统, 序号, 组号, 功能, nvl(主项,0) as 主项 From zlProgrelas Where 主项=1"
        Set mrsRelas = gcnOracle.Execute(strSql)
        
        strSql = "Select 系统, 序号, 组号, 功能, 主项 From zlProgrelas Where 关系=1"
        Set mrsRelExcl = gcnOracle.Execute(strSql)
        
        strSql = "Select 系统, 序号, 功能, 组号, 关系, 主项, 主项关系 From Zlprogrelas"
        Set mrsGroup = gcnOracle.Execute(strSql)
    End If
    Call FillSystem
    If cmbSystem.ListCount < 1 Then
        Unload Me
        Exit Sub
    Else
        If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0
    End If
    If cmbSystem.ListCount = 1 Then
        cmbSystem.Enabled = False
    End If
    Call LvwFlatColumnHeader(lvwFunc(MT_关联模块))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, objNode As Node, lst As ListItem
    Dim strKey As String, strAll As String, strGrant As String
    Dim arrTmp As Variant
    
    If KeyCode = vbKeyF3 Then '查找下一个
        mlngCurPos = FindModule(mlngCurPos)
    ElseIf KeyCode = vbKeyReturn Then   '查找下一个
        If (TypeOf Me.ActiveControl Is TreeView) Then
            If Me.ActiveControl.Index = MT_模块 Then
                mlngCurPos = FindModule(mlngCurPos)
            End If
        End If
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then '折叠展开功能
        LockWindowUpdate Me.hwnd
        Call SynchronizeState '先同步状态
        If cmdExp.Caption = "全部折叠(&D)" Then cmdExp.Tag = 1
        mblnExpanded = Not (mblnExpanded)
        For i = 0 To IIf(msftStyle = SFT_应用系统, 1, 0)
            For Each objNode In tvwMenu(i).Nodes
                objNode.Expanded = mblnExpanded
            Next
            If tvwMenu(i).Nodes.Count > 0 Then
                tvwMenu(i).Nodes(1).Selected = True
                tvwMenu(i).Nodes(1).EnsureVisible
            End If
        Next
        cmdExp.Tag = 0
        cmdExp.Caption = IIf(mblnExpanded, "全部折叠(&D)", "全部展开(&D)")
        LockWindowUpdate 0
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then '全选
        Call SetSel(Me.ActiveControl, True)
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then '全消
        Call SetSel(Me.ActiveControl, False)
    End If
End Sub

Private Sub SetSel(ByRef objCur As Object, Optional ByVal blnSelALl As Boolean = True)
    Dim i As Long, objNode As Node, lst As ListItem
    Dim strKey As String, strAll As String, strGrant As String
    Dim arrTmp As Variant
    
    LockWindowUpdate Me.hwnd
    If blnSelALl Then
        If (TypeOf objCur Is TreeView) Then '定位到树形菜单则只对该树形采取全选或取消权限
            For Each objNode In objCur.Nodes
                objNode.Checked = True
                strKey = GetUpdateKey(objNode.Key)
                If strKey <> "" Then Call UpdateGrantState(strKey, True)
            Next
            If msftStyle <> SFT_其他 Then Call SynchronizeState '同步状态
            If objCur.Tag <> "" Then
                On Error Resume Next
                Set objNode = objCur.Nodes(objCur.Tag)
                If err.Number = 0 Then
                    Call tvwMenu_NodeClick(objCur.Index, objCur.Nodes(objCur.Tag))
                Else
                    err.Clear: objCur.Tag = ""
                End If
                On Error GoTo 0
            End If
        ElseIf (TypeOf objCur Is ListView) Then   '定位到树形菜单则只对该树形采取全选或取消权限
            If objCur.Enabled Then
                For Each lst In objCur.ListItems
                    lst.Checked = True
                    strGrant = strGrant & "," & lst.Text
                Next
                strGrant = Mid(strGrant, 2)
                strKey = GetUpdateKey(tvwMenu(objCur.Index).Tag)
                If strKey <> "" Then
                    arrTmp = Split(strKey, "_")
                    strGrant = CheckFunc(Val(arrTmp(1)), Val(arrTmp(2)), objCur, strGrant)
                    Call UpdateGrantState(strKey, True, strGrant, 1)
                End If
                If msftStyle <> SFT_其他 Then Call SynchronizeState '再次同步状态
            End If
        End If
    Else
        If (TypeOf objCur Is TreeView) Then '定位到树形菜单则只对该树形采取全选或取消权限
            For Each objNode In objCur.Nodes
                objNode.Checked = False
                strKey = GetUpdateKey(objNode.Key)
                If strKey <> "" Then Call UpdateGrantState(strKey, False)
            Next
            If msftStyle <> SFT_其他 Then Call SynchronizeState   '同步状态
            If objCur.Tag <> "" Then
                On Error Resume Next
                Set objNode = objCur.Nodes(objCur.Tag)
                If err.Number = 0 Then
                    Call tvwMenu_NodeClick(objCur.Index, objCur.Nodes(objCur.Tag))
                Else
                    err.Clear: objCur.Tag = ""
                End If
                On Error GoTo 0
            End If
        ElseIf (TypeOf objCur Is ListView) Then   '定位到树形菜单则只对该树形采取全选或取消权限
            If objCur.Enabled Then
                For Each lst In objCur.ListItems
                    lst.Checked = False
                Next
                strKey = GetUpdateKey(tvwMenu(objCur.Index).Tag)
                If strKey <> "" Then
                    arrTmp = Split(strKey, "_")
                    strGrant = CheckFunc(Val(arrTmp(1)), Val(arrTmp(2)), objCur)
                    Call UpdateGrantState(strKey, True, strGrant, 1)
                End If
                If msftStyle <> SFT_其他 Then Call SynchronizeState '同步状态
            End If
        End If
    End If
    LockWindowUpdate 0
End Sub

Private Sub InitTips()
    Dim ObjTip  As clsTipSwap
    Dim i As Integer
    
    Set mcllTip = New Collection
    For i = 0 To 1
        Set ObjTip = New clsTipSwap
        Set ObjTip.ParentControl = lvwFunc(i)
        ObjTip.Icon = TTIconInfo
        ObjTip.Style = TTBalloon
        ObjTip.Create
        mcllTip.Add ObjTip, "T_" & i
    Next

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim lngHeight As Long, lngTop As Long
    On Error Resume Next
    '设置整体界面高度
    If Me.Height < 7000 Then Me.Height = 7000
    If Me.Width < 9300 Then Me.Width = 9300
    '设置角色授权上方区域
    cmbSystem.Left = Me.ScaleWidth - cmbSystem.Width - 60
    lblSys.Left = cmbSystem.Left - lblSys.Width - 30
    fraLine.Width = Me.Width + 100
    '设置角色授权下方区域位置
    lngHeight = Me.ScaleHeight - stbThis.Height
    pgb.Top = lngHeight + (stbThis.Height - pgb.Height) / 2
    pgb.Left = stbThis.Panels(2).Left + Me.TextWidth("字") * 12
    pgb.Width = stbThis.Panels(2).Left + stbThis.Panels(2).Width - pgb.Left - 100
    cmdCancel.Top = lngHeight - cmdCancel.Height - 60
    cmdOK.Top = cmdCancel.Top
    cmdHelp.Top = cmdCancel.Top
    cmdExp.Top = cmdCancel.Top
    cmdSelAll.Top = cmdCancel.Top
    cmdUnSel.Top = cmdCancel.Top
    cmdCheck.Top = cmdCancel.Top
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    cmdCheck.Left = cmdOK.Left - cmdCheck.Width - 500
    lngHeight = cmdCancel.Top - 60 - picMenu(MT_模块).Top
    If fraHSplit.Tag = "" Then '没有拖动就6-4分屏
        fraHSplit.Top = picMenu(MT_模块).Top + lngHeight * 0.6
    End If
    If fraHSplit.Top - fraLine.Top - fraLine.Height < 2000 Then fraHSplit.Top = fraLine.Top + 2000 + fraLine.Height
    fraHSplit.Width = Me.ScaleWidth + 300
    If msftStyle = SFT_其他 Then
        fraHSplit.Visible = False
        lngTop = lngHeight
    Else
        fraHSplit.Visible = True
        lngTop = fraHSplit.Top - picMenu(MT_模块).Top
    End If
    
    '设置角色授权中间区域
    If msftStyle = SFT_应用系统 Then
        picMenu(MT_关联模块).Visible = True
        fraVSplit.Visible = True
        lvwFunc(MT_模块).Visible = True
        lblFuncNote(MT_模块).Visible = True
        If fraVSplit.Tag = "" Then '没有拖动就7-3分屏
            fraVSplit.Left = Me.ScaleWidth * 0.7
        ElseIf Me.ScaleWidth - fraVSplit.Left < 2000 Then
            fraVSplit.Left = Me.ScaleWidth - 2000
        End If
        picMenu(MT_模块).Width = fraVSplit.Left - 15 - picMenu(MT_模块).Left
        picMenu(MT_关联模块).Left = fraVSplit.Left + fraVSplit.Width + 15
        picMenu(MT_关联模块).Width = Me.ScaleWidth - picMenu(MT_关联模块).Left
        For i = 0 To 1
            picMenu(i).Height = lngHeight
            tvwMenu(i).Width = picMenu(i).ScaleWidth - tvwMenu(i).Left
            lvwFunc(i).Width = tvwMenu(i).Width
            tvwMenu(i).Height = lngTop - tvwMenu(i).Top - 30
            lblFuncNote(i).Top = lngTop + fraHSplit.Height + 30
            lvwFunc(i).Top = lblFuncNote(i).Top + lblFuncNote(i).Height + 30
            lvwFunc(i).Height = picMenu(i).ScaleHeight - lvwFunc(i).Top - 30
        Next
        fraVSplit.Top = fraLine.Top - 120
        fraVSplit.Height = cmdCancel.Top - 60 - fraVSplit.Top
    Else
        picMenu(MT_关联模块).Visible = False
        fraVSplit.Visible = False
        picMenu(MT_模块).Width = Me.ScaleWidth - 30 - picMenu(MT_模块).Left
        picMenu(MT_模块).Height = cmdCancel.Top - 60 - picMenu(MT_模块).Top
        tvwMenu(MT_模块).Width = picMenu(MT_模块).ScaleWidth - tvwMenu(MT_模块).Left
        If msftStyle = SFT_自定义报表 Then
            lvwFunc(MT_模块).Visible = True
            lblFuncNote(MT_模块).Visible = True
            lvwFunc(MT_模块).Width = tvwMenu(MT_模块).Width
            tvwMenu(MT_模块).Height = lngTop - tvwMenu(MT_模块).Top - 30
            lblFuncNote(MT_模块).Top = lngTop + fraHSplit.Height + 30
            lvwFunc(MT_模块).Top = lblFuncNote(MT_模块).Top + lblFuncNote(MT_模块).Height + 30
            lvwFunc(MT_模块).Height = picMenu(MT_模块).ScaleHeight - lvwFunc(MT_模块).Top - 30
        Else
             lvwFunc(MT_模块).Visible = False
             lblFuncNote(MT_模块).Visible = False
             tvwMenu(MT_模块).Height = picMenu(MT_模块).ScaleHeight - tvwMenu(MT_模块).Top - 30
        End If
    End If
    '说明列设置
    If lvwFunc(0).View = lvwReport Then
        For i = 0 To 1
            lvwFunc(i).ColumnHeaders(3).Width = lvwFunc(i).Width - lvwFunc(i).ColumnHeaders(1).Width - lvwFunc(i).ColumnHeaders(2).Width
        Next
    End If
    '测试状态，展示所有关联关系
    If BLN_TEST Then
        If picMenu(MT_关联模块).Visible Then
            tvwModRelas.Visible = True
            tvwModRelas.Left = tvwMenu(MT_关联模块).Left + tvwMenu(MT_关联模块).Width * 0.5
            tvwModRelas.Width = tvwMenu(MT_关联模块).Width * 0.5
            tvwModRelas.Top = tvwMenu(MT_关联模块).Top
            tvwModRelas.Height = tvwMenu(MT_关联模块).Height
            tvwModRelas.ZOrder
        End If
    End If
    Exit Sub
'ErrH:
'    If 0 = 1 Then
'        Resume
'    End If
'    MsgBox "From_Resize:" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ClearDataAndVar(True)
End Sub

Private Sub fraHSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then fraHSplit.Top = fraHSplit.Top + Y
End Sub

Private Sub fraHSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fraHSplit.Top - fraLine.Top < 2500 Then fraHSplit.Top = fraLine.Top + 2500
    If fraHSplit.Top > picMenu(0).Height + picMenu(0).Top - 2000 Then fraHSplit.Top = picMenu(0).Height + picMenu(0).Top - 2000
    fraHSplit.Tag = "拖动"
    Call Form_Resize
End Sub

Private Sub fraVSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 1 Then fraVSplit.Left = fraVSplit.Left + X
End Sub

Private Sub fraVSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fraVSplit.Left < 3500 Then fraVSplit.Left = 3500
    If fraVSplit.Left > Me.ScaleWidth - 3500 Then fraVSplit.Left = Me.ScaleWidth - 3500
    fraVSplit.Tag = "拖动"
    Call Form_Resize
End Sub

Private Function ReadData(ByVal rdtInput As ReadDataType) As ADODB.Recordset
'功能：读取数据
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    Select Case rdtInput
        Case RDT_Menu
            strSql = "Select Level As 层次, Id As 编号, Nvl(上级id, 0) As 上级, 标题, Decode(Nvl(短标题, '空'), '空', 标题, 短标题) As 短标题, 快键, 说明," & vbNewLine & _
                            "       Nvl(模块, 0) As 模块, Nvl(系统, 0) As 系统, Nvl(图标, 0) As 图标" & vbNewLine & _
                            "From Zlmenus" & vbNewLine & _
                            "Where 组别 In ('缺省', '导诊')" & vbNewLine & _
                            "Start With 上级id Is Null" & vbNewLine & _
                            "Connect By Prior Id = 上级id"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "模块菜单读取", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "层次,编号,上级,标题,模块,系统,图标")
        Case RDT_Module
            If gblnInIDE Then '调试环境，不检查授权
                strSql = "Select F.序号, F.标题, F.系统, F.功能, F.缺省值, Decode(R.功能, Null, 0, 1) As 授权否" & vbNewLine & _
                            "From (Select C.序号, C.标题, Nvl(B.系统, 0) 系统, B.功能, Nvl(B.缺省值, 0) 缺省值" & vbNewLine & _
                            "       From Zlprogfuncs b, Zlprograms c" & vbNewLine & _
                            "       Where Nvl(C.系统, 0) = Nvl(B.系统, 0) And C.序号 = B.序号) f," & vbNewLine & _
                            "     (Select Nvl(A.系统, 0) 系统, A.序号, A.功能 From Zlrolegrant a Where A.角色 = [1]) r" & vbNewLine & _
                            "Where F.系统 = R.系统(+) And F.序号 = R.序号(+) And F.功能 = R.功能(+)" & vbNewLine & _
                            "Order By F.序号"
            Else
                strSql = "Select G.序号, G.标题, Nvl(G.系统,0) 系统, F.功能, Nvl(F.缺省值, 0) As 缺省值, Decode(R.功能, Null, 0, 1) As 授权否" & vbNewLine & _
                                "From Zlprograms  g," & vbNewLine & _
                                "     (Select 系统, 序号, 功能, 缺省值 From Zlprogfuncs Where 系统 Is Null Or (序号 Between 10000 And 19999)" & vbNewLine & _
                                "       Union" & vbNewLine & _
                                "       Select A.系统, A.程序id As 序号, A.功能, 1 As 缺省值 From Zlreports b, Zlrptputs a Where A.报表id = B.Id And B.系统 Is Null" & vbNewLine & _
                                "       Union" & vbNewLine & _
                                "       Select F.系统, F.序号, F.功能, F.缺省值 From Zlprogfuncs f, Zlregfunc r" & vbNewLine & _
                                "       Where Trunc(F.系统 / 100) = R.系统 And F.序号 = R.序号 And F.功能 = R.功能 And" & vbNewLine & _
                                "             1 = (Select 1 From Zlregaudit a Where A.项目 = '授权证章')) f," & vbNewLine & _
                                "     (Select Nvl(系统, 0) 系统, 序号, 角色, 功能 From Zlrolegrant Where 角色 = [1]) r" & vbNewLine & _
                                "Where Nvl(G.系统, 0) = Nvl(F.系统, 0) And G.序号 = F.序号 And F.序号 = R.序号(+) And F.功能 = R.功能(+) And Nvl(F.系统, 0) = R.系统(+)" & vbNewLine & _
                                "Order By 序号"
            End If
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "授权模块读取", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "序号,标题,系统,功能,授权否,缺省值")
        Case RDT_Function
            strSql = "Select S.名称 系统名, S.编号 系统, S.所有者, Upper(F.函数名) As 函数名, F.说明,F.中文名, Decode(U.Table_Name, Null, 0, 1) 授权否" & vbNewLine & _
                        "From Zlsystems s, Zlfunctions f," & vbNewLine & _
                        "     (Select Table_Schema As Owner, Grantee, Table_Name From All_Tab_Privs Where Table_Schema = User) u," & vbNewLine & _
                        "     (Select Table_Schema As 所有者, Table_Name As 函数, Privilege As 权限" & vbNewLine & _
                        "       From All_Tab_Privs" & vbNewLine & _
                        "       Where Privilege = 'EXECUTE' And Grantable = 'YES'" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select Owner, Object_Name, 'EXECUTE' From All_Objects Where Owner = User And Object_Type = 'FUNCTION') r" & vbNewLine & _
                        "Where F.系统 = S.编号 And S.所有者 = R.所有者 And Upper(F.函数名) = R.函数 And U.Grantee(+) = [1] And U.Owner(+) = R.所有者 And" & vbNewLine & _
                        "      U.Table_Name(+) = R.函数"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "基础函数读取", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "函数名,中文名,说明,系统,系统名,所有者,授权否,授权否 最初授权")
        Case RDT_Table
            strSql = "Select T.系统, T.系统名, T.所有者, T.表名, T.说明, Decode(R.Table_Name, Null, 0, 1) 授权否" & vbNewLine & _
                        "From (Select S.名称 系统名, S.编号 系统, S.所有者, B.表名, B.说明 From Zlsystems s, Zlbasecode b Where B.系统 = S.编号) t," & vbNewLine & _
                        "     (Select Table_Schema As 所有者, Table_Name As 对象" & vbNewLine & _
                        "       From All_Tab_Privs" & vbNewLine & _
                        "       Where Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE') And Grantable = 'YES' And" & vbNewLine & _
                        "             (Table_Schema, Table_Name) In" & vbNewLine & _
                        "             (Select S.所有者, B.表名 From Zlsystems s, Zlbasecode b Where B.系统 = S.编号 And S.所有者 = User)" & vbNewLine & _
                        "       Group By Table_Schema, Table_Name" & vbNewLine & _
                        "       Having Count(Privilege) = 4" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select User, Object_Name" & vbNewLine & _
                        "       From User_Objects" & vbNewLine & _
                        "       Where Object_Type = 'TABLE' And" & vbNewLine & _
                        "             Object_Name In (Select B.表名 From Zlsystems s, Zlbasecode b Where B.系统 = S.编号 And S.所有者 = User)) g," & vbNewLine & _
                        "     (Select Grantor As Owner, Table_Name" & vbNewLine & _
                        "       From All_Tab_Privs" & vbNewLine & _
                        "       Where Grantor = User And Grantee = [1] And Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
                        "       Group By Grantor, Table_Name" & vbNewLine & _
                        "       Having Count(Privilege) = 4) r" & vbNewLine & _
                        "Where T.所有者 = G.所有者 And T.表名 = G.对象 And T.所有者 = R.Owner(+) And T.表名 = R.Table_Name(+)"
'            "Select T.系统,T.系统名 ,T.所有者, T.表名,T.说明, Decode(R.Table_Name, Null, 0, 1) 授权否" & vbNewLine & _
'                            "From (Select S.名称 系统名,S.编号  系统, S.所有者, B.表名, B.说明" & vbNewLine & _
'                            "       From Zlsystems s, Zlbasecode b" & vbNewLine & _
'                            "       Where B.系统 = S.编号) t," & vbNewLine & _
'                            "     (Select 所有者, 对象" & vbNewLine & _
'                            "       From (Select Table_Schema As 所有者, Table_Name As 对象, Privilege As 权限 From All_Tab_Privs Where Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE') And Grantable = 'YES'  Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'DELETE' From All_Objects Where Owner = User And Object_Type = 'TABLE' Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'INSERT' From All_Objects Where Owner = User And Object_Type = 'TABLE' Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'SELECT' From All_Objects Where Owner = User And Object_Type = 'TABLE' Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'UPDATE' From All_Objects Where Owner = User And Object_Type = 'TABLE')" & vbNewLine & _
'                            "       Group By 所有者, 对象" & vbNewLine & _
'                            "       Having Count(权限) = 4) g," & vbNewLine & _
'                            "     (Select Grantor As Owner, Table_Name" & vbNewLine & _
'                            "       From All_Tab_Privs" & vbNewLine & _
'                            "       Where Grantor = User And Grantee =[1] And Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
'                            "       Group By Grantor, Table_Name" & vbNewLine & _
'                            "       Having Count(Privilege) = 4) r" & vbNewLine & _
'                            "Where T.所有者 = G.所有者 And T.表名 = G.对象 And T.所有者 = R.Owner(+) And T.表名 = R.Table_Name(+)"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "基础编码读取", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "表名,说明,系统,系统名,所有者,授权否,授权否 最初授权")
        Case RDT_Systems
            If gblnInIDE Then
                rsTmp.CursorLocation = adUseClient
                '显示可以所有的系统
                Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
                Set ReadData = rsTmp
            Else
                Set ReadData = zlGetRegSystems
            End If
        Case RDT_ModRelas
            strSql = "Select 系统, 模块, 功能, 相关系统, 相关模块, 标题, 缺省值, 相关类型, 相关信息, 类型" & vbNewLine & _
                        "From (With a As (Select A.系统, A.模块, A.功能, Nvl(A.相关系统, 0) 相关系统, A.相关模块, B.标题, A.缺省值, A.相关类型," & vbNewLine & _
                        "                        A.功能 || ',' || A.相关功能 || ',' || A.相关类型 || ',' || A.缺省值 相关信息" & vbNewLine & _
                        "                 From Zlmodulerelas a, Zlprograms b" & vbNewLine & _
                        "                 Where Nvl(A.系统, 0) = Nvl(B.系统, 0) And A.模块 = B.序号)" & vbNewLine & _
                        "       Select A.系统, A.模块, A.功能, A.相关系统, A.相关模块, A.标题, Decode(Sum(A.缺省值), 0, 0, 1) 缺省值, Decode(Sum(A.相关类型), 0, 0, 1) 相关类型," & vbNewLine & _
                        "              F_List2str(Cast(Collect(A.相关信息) As T_Strlist), ';') 相关信息, 0 类型" & vbNewLine & _
                        "       From a" & vbNewLine & _
                        "       Group By A.系统, A.模块, A.功能, A.相关系统, A.相关模块, A.标题" & vbNewLine & _
                        "       Union All" & vbNewLine & _
                        "       Select A.系统, A.模块, Null 功能, A.相关系统, A.相关模块, A.标题, Decode(Sum(A.缺省值), 0, 0, 1) 缺省值," & vbNewLine & _
                        "              Decode(Sum(A.相关类型), 0, 0, 1) 相关类型, F_List2str(Cast(Collect(A.相关信息) As T_Strlist), ';') 相关信息, 1 类型" & vbNewLine & _
                        "       From a" & vbNewLine & _
                        "       Group By A.系统, A.模块, A.相关系统, A.相关模块, A.标题)"
            Set ReadData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "模块关系读取")
    End Select
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "ReadData:" & err.Description, vbInformation, Me.Caption
End Function

Private Sub ClearDataAndVar(Optional ByVal blnFormUnload As Boolean)
'功能：清空一些变量
    Set mclsPrivilege = Nothing
    Set mrsModule = Nothing
    Set mrsTable = Nothing
    Set mrsFunction = Nothing
    Set mrsModsInfo = Nothing
    Set mrsTree = Nothing
    Set mcllHaveSys = Nothing
    Set mcllKeyModule = Nothing
    Set mcllTip = Nothing
    Set mclsPrivilege = Nothing
    Set mrsRelasTree = Nothing
    '模块关系，以及功能关系、系统列表可以不清空，以便下次使用
End Sub

Private Sub FillSystem()
    '显示可以所有的系统
    If mrsSys Is Nothing Then Set mrsSys = ReadData(RDT_Systems)
    If glngSysNo <> -1 Then mrsSys.Filter = "编号 = " & glngSysNo
    cmbSystem.Clear
    mrsSys.Sort = "编号"
    With mrsSys
        Do Until .EOF
            cmbSystem.addItem RPAD(!名称 & "（" & !编号 & "）", 25) & " v" & !版本号
            cmbSystem.ItemData(cmbSystem.NewIndex) = !编号 & ""
            If !所有者 & "" = UCase(gstrUserName) And cmbSystem.ListIndex < 0 Then
                cmbSystem.ListIndex = cmbSystem.NewIndex
            End If
            .MoveNext
        Loop
    End With
    '有两种系统是程序固定的
    If (gobjRegister.zlRegTool And 2) = 2 Then cmbSystem.addItem "自定义报表"
    cmbSystem.addItem "基础工具"
    cmbSystem.addItem "取数函数"
    cmbSystem.addItem "基础编码"
    If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0
End Sub

Private Function GetModuleInfo() As ADODB.Recordset
'功能:获取模块信息
    Dim rsReturn As ADODB.Recordset
    Dim strGrant As String, strDefault As String, intGrant As Integer, strAll As String
    Dim lngSys As Long, lng序号 As Long, str标题 As String
    
    Dim strPre As String
    
    On Error GoTo errH
    If mrsModule Is Nothing Then Set mrsModule = ReadData(RDT_Module)
    '模块类型=0:虚拟模块
    '              =1:实体模块
    '              =2:自定义报表模块或基础工具
    
    '模块相关=0=不存在相关模块不被关联
    '               1=不存在相关模块被关联
    '               2=存在相关模块不被关联
    '               3=存在相关模块被关联
    Set rsReturn = CopyNewRec(mrsModule, True, "序号,标题,系统,授权否", Array("授权改变", adInteger, 1, 0, "改变批次", adInteger, 1, 0, _
                                                        "授权功能", adVarChar, 2000, Empty, "默认功能", adVarChar, 2000, Empty, _
                                                        "所有功能", adVarChar, 2000, Empty, "模块类型", adInteger, 1, 0, "模块相关", adInteger, 1, 0))
    mrsModule.Filter = ""
    mrsModule.Sort = "系统,序号,功能"
    With mrsModule
        Do While Not .EOF
            If strPre <> !系统 & "_" & !序号 Then
                If strPre <> "" Then
                    rsReturn.AddNew Array("序号", "标题", "系统", "授权否", "授权功能", "默认功能", "所有功能", "模块类型", "授权改变", "改变批次", "模块相关"), _
                                                    Array(lng序号, str标题, lngSys, intGrant, Mid(strGrant, 2), Mid(strDefault, 2), Mid(strAll, 2), 0, 0, 0, 0)
                End If
                strGrant = "": strDefault = "": intGrant = 0: strAll = ""
                lngSys = !系统: str标题 = !标题 & "": lng序号 = !序号
                strPre = !系统 & "_" & !序号
            End If
            If !授权否 = 1 Then
                intGrant = 1
            End If
            If !功能 & "" <> "基本" Then
                If !授权否 = 1 Then strGrant = strGrant & "," & !功能
                If !缺省值 = 1 Then strDefault = strDefault & "," & !功能
                strAll = strAll & "," & !功能
            End If
            .MoveNext
        Loop
        '最后一个模块的加入
        If strPre <> "" Then
            rsReturn.AddNew Array("序号", "标题", "系统", "授权否", "授权功能", "默认功能", "所有功能", "模块类型", "授权改变", "改变批次", "模块相关"), _
                                            Array(lng序号, str标题, lngSys, intGrant, Mid(strGrant, 2), Mid(strDefault, 2), Mid(strAll, 2), 0, 0, 0, 0)
        End If
    End With
    Set GetModuleInfo = rsReturn
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "GetModuleInfo:" & err.Description, vbInformation, Me.Caption
End Function

Private Function SetVirtualVisual(ByVal blnVisual As Boolean, Optional ByRef strFirstNode As String) As Boolean
'功能：设置虚拟模块可见性
'参数：blnVisual=是否可见
'返回：第一个节点的Key
    Dim objNode As Node
    
    On Error GoTo errH
    If msftStyle <> SFT_应用系统 Then
        chkVirtual.Visible = False: Exit Function
    End If
    chkVirtual.Visible = True '回复可见,根据该系统状况，进行设置
    mblnVirtual = True
    '不存在实体模块，则只显示虚拟模块，并隐藏复选框
    mrsModsInfo.Filter = "系统=" & mlngSys & " And 模块类型=1" & IIf(chkOnlyShow.value = 1, " And 授权否=1", "")
    If mrsModsInfo.EOF Then
        tvwMenu(MT_模块).Nodes.Clear '清空所有节点
        tvwMenu(MT_模块).Tag = ""
        strFirstNode = ""
        blnVisual = True
        chkVirtual.value = 1
        chkVirtual.Visible = False
    End If
    '不存在虚拟模块，则隐藏复选框
    mrsModsInfo.Filter = "系统=" & mlngSys & " And 模块类型=0  " & IIf(chkOnlyShow.value = 1, " And 授权否=1", "")
    If mrsModsInfo.EOF Then
        blnVisual = False
        chkVirtual.value = 0
        chkVirtual.Visible = False
    End If
    mblnVirtual = False
    On Error Resume Next
    tvwMenu(MT_模块).Nodes.Remove "V_" & mlngSys
    err.Clear: On Error GoTo errH
    If blnVisual Then
        Set objNode = tvwMenu(MT_模块).Nodes.Add(, , "V_" & mlngSys, "虚拟模块", "分类", "分类_选中")
        If chkOnlyShow.value = 1 Then objNode.Checked = True
        With mrsModsInfo
            Do While Not .EOF
                Set objNode = tvwMenu(MT_模块).Nodes.Add("V_" & mlngSys, 4, "M_000000_" & !序号, "【" & Format(!序号, "000000") & "】" & !标题, "Module")
                objNode.Checked = !授权否 = 1
                '不默认勾选节点，防止速度较慢
'                '获取第一节点与第一勾选节点
'                If objNode.Checked And tvwMenu(MT_模块).Tag = "" Then
'                    tvwMenu(MT_模块).Tag = objNode.Key
'                End If
'                If strFirstNode = "" Then
'                    strFirstNode = objNode.Key
'                End If
                .MoveNext
            Loop
        End With
    End If
    SetVirtualVisual = True '返回第一个节点
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "SetVirtualVisual:" & err.Description, vbInformation, Me.Caption
End Function

Private Function SetOnlyShow(Optional ByVal strKey As String) As Boolean
'功能：只显示选中项
'参数：strKey=判断某个节点，为空则判断整个树形
'返回：True=有子级节点或没有子级节点且当前节点被勾选，False=其他情况
    Dim objNode As Node
    Dim strDelKeys As String, strTMp As String
    Dim arrTmp As Variant, i As Long
    
    If chkOnlyShow.value = 1 Then
        If tvwMenu(MT_模块).Nodes.Count = 0 Then Exit Function
        '获取初始节点，以便循环
        If strKey = "" Then
            If GetUpdateKey(tvwMenu(MT_模块).Tag) <> "" Then
                If Not tvwMenu(MT_模块).Nodes(tvwMenu(MT_模块).Tag).Checked Then
                    tvwMenu(MT_模块).Tag = ""
                End If
            Else
                tvwMenu(MT_模块).Tag = ""
            End If
            Set objNode = tvwMenu(MT_模块).Nodes(1)
        ElseIf tvwMenu(MT_模块).Nodes(strKey).Children <> 0 Then
            Set objNode = tvwMenu(MT_模块).Nodes(strKey).Child
        End If
        '获取可以删除的节点
        Do While Not objNode Is Nothing
            '若子级被选中，则父级选中
            objNode.Checked = SetOnlyShow(objNode.Key)
            If Not objNode.Checked Then
                strDelKeys = strDelKeys & "|" & objNode.Key
            Else
                If tvwMenu(MT_模块).Tag = "" And GetUpdateKey(objNode.Key) <> "" Then
                    tvwMenu(MT_模块).Tag = objNode.Key
                End If
            End If
            Set objNode = objNode.Next
        Loop
        '删除可以删除的节点
        arrTmp = Split(Mid(strDelKeys, 2), "|")
        For i = LBound(arrTmp) To UBound(arrTmp)
            tvwMenu(MT_模块).Nodes.Remove arrTmp(i)
        Next
        If strKey <> "" Then
            '若当前父级节点有未删除的子级节点，则不用保留父级节点
            If tvwMenu(MT_模块).Nodes(strKey).Children <> 0 Then
                Set objNode = tvwMenu(MT_模块).Nodes(strKey).Child
                SetOnlyShow = True
            Else '若父级节点没有子节点，判断节点是否被选择
                SetOnlyShow = tvwMenu(MT_模块).Nodes(strKey).Checked
            End If
        End If
        If strKey = "" Then
            Call SetVirtualVisual(chkVirtual.value <> 0)
            '展开标记节点
            If tvwMenu(MT_模块).Tag <> "" Then
                strTMp = tvwMenu(MT_模块).Tag: tvwMenu(MT_模块).Tag = ""
                Call SetNodeExpand(tvwMenu(MT_模块), strTMp) '展开节点
                Call tvwMenu_NodeClick(MT_模块, tvwMenu(MT_模块).Nodes(strTMp))
            End If
        End If
    Else
        cmbSystem.Tag = "-1"
        Call cmbSystem_Click
    End If
End Function

Private Function FindModule(Optional ByVal intCurPosition As Long, Optional ByVal blnSmart As Boolean) As Long
'功能：进行模块查找
'参数：intCurPosition=当前位置，<=1表示从头到尾开始查找，否则从当前位置开始查找
'          blnSmart=True-搜索框输入内容灵敏适应,False-不灵敏适应搜索框
'返回：匹配项目位置
    Dim i As Integer
    Dim blnFind As Boolean
    Dim strLike As String, strKeyLike As String
    Dim objNode As Node
    Dim strMsg As String
    Dim strName As String
    
    On Error Resume Next
    If intCurPosition < 0 Then FindModule = -1: Exit Function
    '起始位置处理
    If intCurPosition >= tvwMenu(MT_模块).Nodes.Count Then
        intCurPosition = 0
    End If
    '查找字符串解析
    strName = cmbSystem.Text
    If msftStyle = SFT_其他 Then
        strLike = "*" & mstrFind & "*"
        If cmbSystem.Text = "基础编码" Then
            strKeyLike = "T_*_*"
        Else '取数函数
            strKeyLike = "F_*_*"
        End If
    Else
        If msftStyle <> SFT_自定义报表 Then
            strName = "模块"
        End If
        If IsNumeric(mstrFind) Then '按编号查找
            strLike = "【*" & mstrFind & "*】*"
        Else '按名称查获
            strLike = "【*】*" & mstrFind & "*"
        End If
        strKeyLike = "M_*_*"
    End If
    '进行查找
    For i = intCurPosition + 1 To tvwMenu(MT_模块).Nodes.Count
        Set objNode = tvwMenu(MT_模块).Nodes(i)
        If objNode.Key Like strKeyLike Then
            If objNode.Text Like strLike Then
                objNode.Expanded = True
                objNode.Selected = True: blnFind = True
                Exit For
            End If
        End If
    Next
    '未查找到原因提示
    If Not blnFind Then
        If mblnReturn Then
            If mlngCurPos <= 1 Then
                If chkOnlyShow.value = 1 Then
                    strMsg = "你所查找的" & strName & "未勾选，请取消勾选""仅已授权""再进行查找！"
                ElseIf chkVirtual.Visible And chkVirtual.value = 0 Then
                    strMsg = "你所查找的" & strName & "可能是虚拟模块，请勾选""含虚拟模块""再进行查找！"
                Else
                    strMsg = "未找到匹配的" & strName & "！"
                End If
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, Me.Caption
                End If
                mlngCurPos = -1
                '提示是否从头开始查找
            Else
                If MsgBox("未找到匹配的" & strName & "，是否重新进行查找", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    mlngCurPos = 0
                    mlngCurPos = FindModule(mlngCurPos)
                    FindModule = mlngCurPos
                Else
                    FindModule = -1
                End If
            End If
        Else
            FindModule = -1
        End If
    Else
        FindModule = i
        Call tvwMenu_NodeClick(MT_模块, objNode)
    End If
End Function

Private Sub lvwFunc_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvwFunc(Index).SortKey = ColumnHeader.Index - 1 Then
        lvwFunc(Index).SortOrder = IIf(lvwFunc(Index).SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwFunc(Index).SortKey = ColumnHeader.Index - 1
        lvwFunc(Index).SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwFunc_DblClick(Index As Integer)
    Dim objList As ListItem
    On Error Resume Next
    Set objList = lvwFunc(Index).ListItems(lvwFunc(Index).Tag)
    err.Clear: On Error GoTo 0
    If Not objList Is Nothing Then
        objList.Checked = Not objList.Checked
        Call lvwFunc_ItemCheck(Index, objList)
    End If
End Sub

Private Sub lvwFunc_GotFocus(Index As Integer)
    mintActive = Index + 2
End Sub

Private Sub lvwFunc_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim arrTmp As Variant
    Dim lng系统 As Long, lng序号 As Long
    Dim strPrivs As String, strTMp As String
    Dim objItem As ListItem
    Dim blnChange As Boolean
    
    arrTmp = Split(tvwMenu(Index).Tag, "_")
    If Index = MT_模块 Then
        lng系统 = mlngSys
        lng序号 = Val(arrTmp(2))
    Else
        lng系统 = Val(arrTmp(3))
        lng序号 = Val(arrTmp(4))
    End If
    If mblnItem Then Exit Sub
    If Item.Checked Then
        For Each objItem In lvwFunc(Index).ListItems
            If objItem.Checked = True Then
                strPrivs = strPrivs & "," & objItem.Text
            End If
        Next
        If strPrivs <> "" Then strPrivs = Mid(strPrivs, 2)
        '互斥关系,在此处理
        mrsRelExcl.Filter = "系统 = " & lng系统 & " And 序号 = " & lng序号 & " And 功能 = '" & Item.Text & "'"
        If Not mrsRelExcl.EOF Then
            mrsGroup.Filter = "系统 = " & lng系统 & " And 序号 = " & lng序号 & " And 组号 = " & mrsRelExcl!组号
            If Not mrsGroup.EOF Then
                strPrivs = setExcl(mrsGroup, mrsRelExcl!功能 & "_" & mrsRelExcl!组号, Item.Checked, lvwFunc(Index), strPrivs)
            End If
        End If
    End If
    '主从关系,在此处理
    strTMp = strPrivs
    strPrivs = CheckFunc(lng系统, lng序号, lvwFunc(Index), strPrivs)
    blnChange = strTMp <> strPrivs
    '更新授权情况
    Call UpdateGrantState("M_" & lng系统 & "_" & lng序号, True, strPrivs, 1)
    '先一步同步勾选模块的授权，否则会发生如下情况
    '先取消模块固定关联模块在右边的对应模块，再直接取消该模块授权，则发现取消的固定关联模块又被勾选的
    Call lvwFunc_ItemClick(Index, Item)
    If Index = MT_模块 Then
        blnChange = blnChange Or CheckNode(tvwMenu(MT_关联模块))
    End If
    '发生授权改变，则同步记录集
    If blnChange Then Call SynchronizeState
End Sub

Private Sub lvwFunc_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim strKey As String, arrTmp As Variant
    
    If lvwFunc(Index).Tag <> Item.Key Or mstrCurRelas Like "*_*_" Then '上次选中模块
        lvwFunc(Index).Tag = Item.Key
        If Index = MT_模块 Then  '模块类型
            '加载关联模块
            strKey = mlngSys & "_" & Val(Split(tvwMenu(Index).Tag, "_")(2)) & "_" & Item.Text
            Call FillRelasModule(strKey)
        End If
    End If
    Item.Selected = True
End Sub

Private Sub lvwFunc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objItem As ListItem, strTip As String
    
    Set objItem = lvwFunc(Index).HitTest(X, Y)
    If mcllTip Is Nothing Then Call InitTips
    If Not objItem Is Nothing Then
        strTip = objItem.SubItems(2)
        If strTip = "" Then strTip = "<无说明信息>"
        mcllTip("T_" & Index).TipText = SwapText(strTip)
        mcllTip("T_" & Index).Title = objItem.Text
    Else
        mcllTip("T_" & Index).TipText = ""
        mcllTip("T_" & Index).Title = ""
    End If
End Sub

Private Function SwapText(ByVal strTxt As String) As String
    
    Dim strReturn As String, strTMp As String, i As Integer
    strReturn = strTxt
    If InStr(strTxt, ";") > 0 Then
        strReturn = SwapWord(strReturn, ";")
    End If
    If InStr(strTxt, "；") > 0 Then
        strReturn = SwapWord(strReturn, "；")
    End If
    If InStr(strTxt, ".") > 0 Then
        strReturn = SwapWord(strReturn, ".")
    End If
    If InStr(strTxt, "。") > 0 Then
        strReturn = SwapWord(strReturn, "。")
    End If
    
    If strReturn = strTxt Then
        strReturn = swapLine("　　" & strTxt)
    End If
    '--
    strReturn = Replace(strReturn, " ", "")
    strReturn = Replace(strReturn, "　", "")
    strReturn = Replace(strReturn, "[CR]。[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]；[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR];[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR].[CR]", "[CR]")
    
    '行内换行
    Dim varLine As Variant
    
    varLine = Split(strReturn, "[CR]")
    For i = 0 To UBound(varLine)
        strTMp = strTMp & swapLine("　　" & varLine(i)) & vbNewLine
    Next
    
    If strTMp <> "" Then
        strReturn = strTMp
    End If
    '--　清除多余的空行
    strReturn = ClearLine(strReturn)
    SwapText = strReturn
End Function

Private Function ClearLine(strTxt) As String
    Dim i As Integer, Y As Integer
    Dim varLine As Variant
    Dim strReturn As String
    varLine = Split(strTxt, vbNewLine)
    For i = 0 To UBound(varLine)
        If InStr(",.;?!])}%>，。；！？：）］｝、》％’”", Mid(varLine(i), 1, 1)) > 0 Then
            strReturn = Mid(strReturn, 1, Len(strReturn) - 4) & Mid(varLine(i), 1, 1) & "[CR]" & Mid(varLine(i), 2) & "[CR]"
        Else
            strReturn = strReturn & varLine(i) & "[CR]"
        End If
    Next
    
    strReturn = Replace(strReturn, "[CR]。[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]；[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR];[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR].[CR]", "[CR]")
    
    strReturn = Replace(strReturn, "[CR][CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]", vbNewLine)
    ClearLine = strReturn
End Function

Private Function SwapWord(ByVal strTxt As String, strWord As String) As String
    Dim varLine As Variant
    Dim strReturn As String
    Dim i As Integer
    Dim strTxtTmp As String
    
    strTxtTmp = strTxt
    If Mid(strTxt, Len(strTxt), 1) = strWord Then
        strTxtTmp = Mid(strTxt, 1, Len(strTxt) - 1)
    End If
    
    If InStr(strTxtTmp, strWord) > 0 Then
        varLine = Split(strTxtTmp, strWord)
        For i = 0 To UBound(varLine)
            If varLine(i) <> "" Then
                'varLine(i) = swapLine("　　" & varLine(i))
                If varLine(i) & strWord <> strWord Then
                    strReturn = strReturn & varLine(i) & strWord & "[CR]"
                End If
            End If
        Next
    End If
    'If Mid(strTxtTmp, Len(strTxtTmp), 1) <> strWord Then strReturn = Mid(strReturn, 1, Len(strReturn) - 1)
    If strReturn <> "" Then
        SwapWord = strReturn
    Else
        SwapWord = strTxt
    End If
End Function

Private Function swapLine(ByVal strTxt As String) As String
    Dim strTMp As String
    strTMp = strTxt
    
    If Len(strTxt) > 18 Then
        swapLine = Mid(strTMp, 1, 18) & vbNewLine
        strTMp = Mid(strTMp, 19)
        swapLine = swapLine & swapLine(strTMp)
    Else
        swapLine = strTxt
    End If
End Function

Private Sub lvwFunc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call PopupMenu(Me.mnuPopu, 2)
    End If
End Sub

Private Sub mclsPrivilege_AfterProgress()
    stbThis.Panels(2).Text = ""
    pgb.value = 0
    DoEvents
End Sub

Private Sub mclsPrivilege_BeforeProgress(ByVal Title As String, ByVal Max As Long)
    stbThis.Panels(2).Text = Title
    pgb.Max = Max
    DoEvents
End Sub

Private Sub mclsPrivilege_Progressing(ByVal Progress As Long)
    pgb.value = Progress
End Sub

Private Sub mnuPopuState_Click(Index As Integer)
    Dim i As Integer
    mnuPopuState(0).Checked = (Index = 0)
    mnuPopuState(1).Checked = (Index = 1)
    For i = 0 To 1
        lvwFunc(i).View = IIf(Index = 0, lvwSmallIcon, lvwReport)
    Next
End Sub

Private Sub tvwMenu_Click(Index As Integer)
    Dim blnDo As Boolean, objNode As Node
    '固定关联关系不能取消勾选
    If Index = MT_关联模块 And tvwMenu(MT_关联模块).Tag <> "" Then
        Set objNode = tvwMenu(MT_关联模块).Nodes(tvwMenu(MT_关联模块).Tag)
        blnDo = RelasCanSet(objNode.Key)
    Else
        blnDo = True
        If Index = MT_模块 Then
            On Error Resume Next
            Set objNode = tvwMenu(MT_模块).Nodes(tvwMenu(MT_模块).Tag)
            If err.Number = 0 Then
                On Error GoTo 0
                Call tvwMenu_NodeClick(Index, objNode)
            Else
                On Error GoTo 0
            End If
        End If
    End If
    If Not blnDo Then objNode.Checked = Not objNode.Checked
End Sub

Private Sub tvwMenu_Collapse(Index As Integer, ByVal Node As MSComctlLib.Node)
    If Val(cmdExp.Tag) = 1 Then Exit Sub
    Call tvwMenu_NodeClick(Index, Node)
End Sub

Private Sub tvwMenu_DblClick(Index As Integer)
    Dim objNode As Node
    Dim blnDo As Boolean
    On Error Resume Next
    Set objNode = tvwMenu(Index).Nodes(tvwMenu(Index).Tag)
    err.Clear: On Error GoTo 0
    If Not objNode Is Nothing Then
        '固定关联关系不能取消勾选
        If Index = MT_关联模块 Then
            blnDo = RelasCanSet(objNode.Key, True)
        Else
            blnDo = True
        End If
        If objNode.Children = 0 And blnDo Then
            objNode.Checked = Not objNode.Checked
            Call tvwMenu_NodeCheck(Index, objNode)
        End If
    End If
End Sub

Private Sub tvwMenu_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)
    Node.ExpandedImage = "分类_选中"
End Sub

Private Sub tvwMenu_GotFocus(Index As Integer)
    mintActive = Index
End Sub

Private Sub tvwMenu_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)
    Dim blnDo As Boolean
    Dim blnUpdate As Boolean, blnTmp As Boolean
    Dim arrTmp As Variant
    Dim strKey As String
    '设置节点点击状态
    If Index = MT_关联模块 Then
        blnDo = RelasCanSet(Node.Key)
    Else
        blnDo = True
    End If
    '先一步同步勾选模块的授权，否则会发生如下情况
    '先取消模块固定关联模块在右边的对应模块，再直接取消该模块授权，则发现取消的固定关联模块又被勾选的
    If blnDo Then
        strKey = GetUpdateKey(Node.Key)
        If strKey <> "" Then
            If strKey Like "M*" Then blnUpdate = True
            Call UpdateGrantState(strKey, Node.Checked)
        End If
    End If
    Call tvwMenu_NodeClick(Index, Node)
    '勾选节点的其他处理
    If blnDo Then
        '更新授权
        If Index = MT_模块 And Node.Key Like "M*" Then
            blnTmp = CheckNode(tvwMenu(MT_关联模块))
            blnTmp = blnTmp Or CheckNode(tvwMenu(Index), Node.Key)
        Else
            blnTmp = CheckNode(tvwMenu(Index), Node.Key)
        End If
        If blnUpdate Or blnTmp Then Call SynchronizeState
    End If
End Sub

Private Sub AddModuleNode(ByVal strKey As String, ByVal strName As String, ByVal intType As Integer)
    '入参：节点key值，节点名称，节点类型（1：实体模块，0：虚拟模块）
    '功能：添加指定节点
    Dim objNode As Node
    Dim colNodes As Collection
    Dim arrNodes() As String
    Dim strNodeKye As String
    Dim i As Long, j As Long, lngRelative As Long
    
    If intType = 0 Then
        '位于虚拟模块
        Set objNode = tvwMenu(MT_模块).Nodes.Add("V_" & mlngSys, tvwChild, "M_000000" & "_" & strKey, "【" & Format(strKey, "000000") & "】" & strName, "Module")
        objNode.Checked = True
    Else
        '位于非虚拟模块
        mrsTree.Filter = "模块 = " & strKey
        mrsTree.Filter = "编号 = " & mrsTree!上级
        With mrsTree
            On Error Resume Next
            Set objNode = tvwMenu(MT_模块).Nodes("K_" & Format(!编号, "000000"))
            If err.Number <> 0 Then
                err.Clear
                '查找节点树
                Set colNodes = New Collection
                Call FindModulePath(mrsTree!编号, colNodes)
                '添加节点树
                For i = colNodes.Count To 1 Step -1
                    arrNodes = Split(colNodes(i), "?")
                    On Error Resume Next
                    '遍历所有顶级节点并排序
                    strNodeKye = FindNodePosition(arrNodes, lngRelative)
                    Set objNode = tvwMenu(MT_模块).Nodes.Add(strNodeKye, lngRelative, "K_" & Format(arrNodes(0), "000000"), arrNodes(2), "分类", "分类_选中")
                    objNode.Checked = True
                    colNodes.Remove i
                    If err.Number <> 0 Then err.Clear
                Next
                Call AddModuleNode(strKey, "", 1)
            Else
                mrsTree.Filter = "模块 = " & strKey
                Set objNode = tvwMenu(MT_模块).Nodes.Add("K_" & Format(!上级, "000000"), tvwChild, "M_" & Format(!编号, "000000") & "_" & !模块, "【" & Format(!模块, "000000") & "】" & !标题, "Module")
                objNode.Checked = True
            End If
        End With
    End If
End Sub

Private Function FindNodePosition(arrNodes() As String, lngRelative As Long) As String
    '入参：存储节点信息的数组
    '       arrNodes():节点数组
    '       FindNodePosition:要插入的节点的相邻节点
    '       lngRelative:要插入的节点和相邻节点的相对位置
    '出参：返回要插入节点的相邻节点和他们的相对位置
    '功能：查找节点在树形中的相对位置
    Dim objNode As Node
    Dim j As Long
    
    If arrNodes(1) = 0 Then
        Set objNode = tvwMenu(MT_模块).Nodes(1).FirstSibling
    Else
        Set objNode = tvwMenu(MT_模块).Nodes("K_" & Format(arrNodes(1), "000000")).Child
        If objNode Is Nothing Then
            FindNodePosition = "K_" & Format(arrNodes(1), "000000")
            lngRelative = tvwChild
            Exit Function
        End If
    End If
    Do While Not objNode Is Nothing
        If arrNodes(0) < Val(Split(objNode.Key, "_")(1)) And Split(objNode.Key, "_")(0) = "K" Then
            FindNodePosition = objNode.Key
            lngRelative = tvwPrevious
            Exit Function
        End If
        If objNode.Next Is Nothing Or objNode.Next.Key = "V_" & mlngSys Then
            FindNodePosition = objNode.Key
            lngRelative = tvwNext
            Exit Function
        Else
            Set objNode = objNode.Next
        End If
    Loop
End Function

Private Sub FindModulePath(ByVal strNum As String, colNodes As Collection)
    '功能：获取一个节点到根节点的路径
    '入参：节点的编号，存储节点路径的集合对象
    Dim objNode As Node

    mrsTree.Filter = "编号 = " & strNum
    On Error Resume Next
    Set objNode = tvwMenu(MT_模块).Nodes("K_" & Format(strNum, "000000"))
    If err.Number <> 0 Then
        err.Clear
        colNodes.Add mrsTree!编号 & "?" & mrsTree!上级 & "?" & mrsTree!标题
        If mrsTree!上级 <> 0 Then
            Call FindModulePath(mrsTree!上级, colNodes)
        End If
    End If
End Sub

Private Sub tvwMenu_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
'主要设置获得焦点节点的状态
    Dim arrTmp As Variant
    Dim ctTmp As ClearType
    Dim strKey As String
    
    lblNotice.Caption = ""
    '如果上次是功能的相关模块，则重新加载相关模块
    If tvwMenu(Index).Tag <> Node.Key Or mstrCurRelas Like "*_*_?*" Then
        If tvwMenu(Index).Tag <> "" And Index = MT_模块 And chkOnlyShow.value = False Then
            tvwMenu(Index).Nodes(tvwMenu(Index).Tag).Bold = False
        End If
        ctTmp = -1
        If Node.Key Like "M_*" And Index = MT_模块 Then
            strKey = mlngSys & "_" & Val(Split(Node.Key, "_")(2)) & "_"
        End If
        If tvwMenu(Index).Tag <> Node.Key And Index = MT_模块 Then
            ctTmp = CT_功能
        ElseIf tvwMenu(Index).Tag <> Node.Key Then
            ctTmp = CT_关联功能
        ElseIf Index <> MT_关联模块 Then
            If strKey <> "" And mstrCurRelas Like strKey & "*" Then
            Else
                ctTmp = CT_关联模块
            End If
        End If
        tvwMenu(Index).Tag = Node.Key
        Call ClearFace(ctTmp) '清空界面
        If strKey <> "" Then
            Call FillRelasModule(strKey)
        End If
        If Node.Key Like "M_*" Then
            If ctTmp = CT_功能 Then
                Call FillFunc(lvwFunc(MT_模块))
            End If
            If ctTmp >= CT_关联模块 Then
            ElseIf ctTmp >= CT_关联功能 Then
                Call FillFunc(lvwFunc(MT_关联模块))
            End If
        End If
    End If
    If Index = MT_模块 Then
        lvwFunc(MT_模块).Enabled = Node.Checked
        If Not Node.Checked Then
            Set lvwFunc(MT_模块).SelectedItem = Nothing
        End If
        lvwFunc(MT_模块).BackColor = IIf(lvwFunc(MT_模块).Enabled, &H80000005, &H8000000F)
    ElseIf RelasCanSet(Node.Key) Then
         lvwFunc(MT_关联模块).Enabled = Node.Checked
         If Not Node.Checked Then
            Set lvwFunc(MT_关联模块).SelectedItem = Nothing
         End If
         lvwFunc(MT_关联模块).BackColor = IIf(lvwFunc(MT_关联模块).Enabled, &HEFF0E0, &H8000000F)
    End If
    If Index = MT_模块 Then Node.Bold = True
    Node.Selected = True
End Sub

Private Sub ClearFace(Optional ByVal ctInput As ClearType)
'功能：清空界面的部分区域
    'CT_Sys需要特别处理的内容
    If ctInput = CT_Sys Then
        tvwMenu(MT_模块).Nodes.Clear: tvwMenu(MT_模块).Tag = ""
        mblnClear = True: txtSearch.Text = "": mblnClear = False
        mblnExpanded = False: mintActive = 0
        cmdExp.Caption = IIf(mblnExpanded, "全部折叠&D)", "全部展开(&D)")
    End If
    'CT_Sys，CT_功能需要特别处理的内容
    If ctInput >= CT_功能 Then
        lvwFunc(MT_模块).ListItems.Clear: lvwFunc(MT_模块).Tag = ""
        mstrCurRelas = ""
    End If
    'CT_Sys，CT_关联模块，CT_功能，需要特别处理的内容
    If ctInput >= CT_关联模块 Then
        tvwMenu(MT_关联模块).Nodes.Clear: tvwMenu(MT_关联模块).Tag = ""
    End If
    'CT_Sys，CT_关联模块，CT_功能，CT_关联功能需要特别处理的内容
    If ctInput >= CT_关联功能 Then
        lvwFunc(MT_关联模块).ListItems.Clear: lvwFunc(MT_关联模块).Tag = ""
    End If
End Sub

Private Sub txtSearch_Change()
    mlngCurPos = 0
    mstrFind = txtSearch.Text
    mblnReturn = False
    If mstrFind <> "" And Not mblnClear Then
        mlngCurPos = FindModule(mlngCurPos)
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "%" Or Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "*" Or Chr(KeyAscii) = "_" Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        If Not mblnReturn Then
            mblnReturn = True
            mlngCurPos = 0
            mstrFind = txtSearch.Text
            mlngCurPos = FindModule(mlngCurPos, True)
        End If
    End If
End Sub

Private Function CheckNode(ByRef tvwInput As TreeView, Optional ByVal strKey As String, Optional ByVal lngLevel As Long, Optional ByRef lngCount As Long) As Boolean
'功能：设置模块菜单勾选状态
'返回：勾选状态是否发生变化
    Dim blnCheck As Boolean, blnChildCheck As Boolean
    Dim objNode As Node, objParent As Node
    Dim arrTmp As Variant
    Dim strKeyTmp As String
    Dim strGrant As String
    
    On Error GoTo errH
    With tvwInput
        If tvwInput.Index = MT_模块 Then
            If .Nodes.Count = 0 Then Exit Function
            If strKey = "" Then '初始设置勾选状态
                Set objNode = .Nodes(1)
                Do While Not objNode Is Nothing
                    blnCheck = CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount) '递归设置
                    strKeyTmp = GetUpdateKey(objNode.Key)
                    If objNode.Checked <> blnCheck And strKeyTmp <> "" Then
                        lngCount = lngCount + 1
                        '更新授权
                        Call UpdateGrantState(strKeyTmp, blnCheck)
                    End If
                    objNode.Checked = blnCheck
                    Set objNode = objNode.Next
                Loop
            Else
                Set objParent = .Nodes(strKey)
                If lngLevel <= 0 Then '手工勾选
                    '向下级开始
                    blnCheck = objParent.Checked
                    Set objNode = objParent.Child
                    Do While Not objNode Is Nothing
                        '勾选状态与父级不同，则更新授权记录集
                        strKeyTmp = GetUpdateKey(objNode.Key)
                        If strKeyTmp <> "" And objNode.Checked <> blnCheck Then
                            lngCount = lngCount + 1
                            '更新授权
                            Call UpdateGrantState(strKeyTmp, blnCheck)
                        End If
                        objNode.Checked = blnCheck '设置借点勾选状态
                        If objNode.Children <> 0 Then '如果有子节点，则递归
                            Call CheckNode(tvwInput, objNode.Key, -1, lngCount)
                        End If
                        Set objNode = objNode.Next
                    Loop
                    If lngLevel = 0 Then
                        '向上设置
                        Set objParent = .Nodes(strKey)
                        Do While Not objParent.Parent Is Nothing
                            Set objParent = objParent.Parent
                            blnChildCheck = True
                            Set objNode = objParent.Child
                            Do While Not objNode Is Nothing
                                If Not objNode.Checked Then
                                    blnChildCheck = False '子级有一个未勾选，父级不勾选
                                    Exit Do
                                End If
                                Set objNode = objNode.Next
                            Loop
                            '勾选状态与父级不同，则更新授权记录集
                            strKeyTmp = GetUpdateKey(objParent.Key)
                            If strKeyTmp <> "" And objParent.Checked <> blnChildCheck Then
                                lngCount = lngCount + 1
                                '更新授权
                                Call UpdateGrantState(strKeyTmp, blnCheck)
                            End If
                            objParent.Checked = blnChildCheck
                            '进入下一次循环
                        Loop
                    End If
                Else
                    Set objParent = .Nodes(strKey)
                    If objParent.Children <> 0 Then '有子级的才判断
                        Set objNode = objParent.Child
                        blnChildCheck = True
                        Do While Not objNode Is Nothing
                            If objNode.Children <> 0 Then
                                blnCheck = CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount)
                                strKeyTmp = GetUpdateKey(objNode.Key)
                                If strKeyTmp <> "" And objNode.Checked <> blnCheck Then
                                    lngCount = lngCount + 1
                                    Call UpdateGrantState(strKeyTmp, blnCheck, , 1)
                                End If
                                objNode.Checked = blnCheck
                            End If
                            blnChildCheck = blnChildCheck And objNode.Checked
                            '子级有一个未勾选，父级不勾选
                            If Not blnChildCheck Then Exit Do
                            Set objNode = objNode.Next
                        Loop
                        CheckNode = blnChildCheck
                    Else
                        CheckNode = objParent.Checked
                    End If
                End If
            End If
        Else '相关模块的设置
            If .Nodes.Count = 0 Then Exit Function
            If strKey = "" Then '初始设置勾选状态不需要，因为在加载时已经设置
                arrTmp = Split(mstrCurRelas, "_")
                If arrTmp(2) <> "" Then
                    blnCheck = lvwFunc(MT_模块).ListItems("F_" & arrTmp(2)).Checked
                Else
                    blnCheck = tvwMenu(MT_模块).Nodes(tvwMenu(MT_模块).Tag).Checked
                End If
                Set objNode = .Nodes(1)
                Do While Not objNode Is Nothing
                    arrTmp = Split(objNode.Key, "_")
                    mrsModsInfo.Filter = "系统=" & arrTmp(3) & " And 序号=" & arrTmp(4)
                    blnChildCheck = mrsModsInfo!授权否
                    If Not blnChildCheck Then
                        blnChildCheck = blnCheck And objNode.Tag Like "1^*"
                    End If
                    If objNode.Checked <> blnChildCheck Then
                        lngCount = lngCount + 1
                        If blnChildCheck And mrsModsInfo!授权否 = 0 Then '未授权，但是需要进行授权
                            strGrant = mrsModsInfo!默认功能
                            mrsModsInfo.Filter = "系统=" & arrTmp(1) & " And 序号=" & arrTmp(2)
                            strGrant = GetGrantByRelasInfo(mrsModsInfo!授权功能 & "", strGrant, Split(objNode.Tag, "^")(1))
                        End If
                        '更新授权
                        Call UpdateGrantState("M_" & arrTmp(3) & "_" & arrTmp(4), blnChildCheck, strGrant)
                    End If
                    objNode.Checked = blnChildCheck
                    If objNode.Children <> 0 Then
                        Call CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount) '递归设置
                    End If
                    Set objNode = objNode.Next
                Loop
            Else
                Set objParent = .Nodes(strKey)
                blnCheck = objParent.Checked
                Set objNode = objParent.Child
                Do While Not objNode Is Nothing
                    If objNode.Children <> 0 Then
                        Call CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount)
                    End If
                    arrTmp = Split(objNode.Key, "_")
                    mrsModsInfo.Filter = "系统=" & arrTmp(3) & " And 序号=" & arrTmp(4)
                    '判断子级勾选
                    If blnCheck Then
                        blnChildCheck = objNode.Tag Like "1^*"
                        If Not blnChildCheck Then
                            blnChildCheck = mrsModsInfo!授权否
                        End If
                    Else
                        blnChildCheck = blnCheck
                    End If
                    If objNode.Checked <> blnChildCheck Then
                        lngCount = lngCount + 1
                        If TypeName(arrTmp) <> "String()" Then
                            arrTmp = Split(objNode.Key, "_")
                            mrsModsInfo.Filter = "系统=" & arrTmp(3) & " And 序号=" & arrTmp(4)
                        End If
                        If blnChildCheck And mrsModsInfo!授权否 = 0 Then
                            strGrant = mrsModsInfo!默认功能
                            mrsModsInfo.Filter = "系统=" & arrTmp(1) & " And 序号=" & arrTmp(2)
                            strGrant = GetGrantByRelasInfo(mrsModsInfo!授权功能 & "", strGrant, Split(objNode.Tag, "^")(1))
                        End If
                        '更新授权
                        Call UpdateGrantState("M_" & arrTmp(3) & "_" & arrTmp(4), blnChildCheck, strGrant)
                    End If
                    objNode.Checked = blnChildCheck
                    Set objNode = objNode.Next
                Loop
            End If
        End If
    End With
    If lngLevel = 0 Then
        '#ADD#添加将授权结果同步到树形,更新树形记录集状态
        CheckNode = lngCount <> 0
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "CheckNode:" & err.Description, vbInformation, Me.Caption
End Function

Private Function GetUpdateKey(ByVal strKey As String) As String
'功能：获取数据缓存更新的Key
    Dim arrTmp As Variant
    
    If strKey Like "M*" Then
        arrTmp = Split(strKey, "_")
        If UBound(arrTmp) = 2 Then
            GetUpdateKey = "M_" & mlngSys & "_" & arrTmp(2)
        ElseIf UBound(arrTmp) = 4 Then
            GetUpdateKey = "M_" & arrTmp(3) & "_" & arrTmp(4)
        End If
    ElseIf strKey Like "T*" Then
        GetUpdateKey = strKey
    ElseIf strKey Like "F*" Then
        GetUpdateKey = strKey
    End If
End Function

Private Function RelasCanSet(ByVal strKey As String, Optional ByVal blnDblClick As Boolean) As Boolean
'功能：判断关联模块节点是否可以勾选以及取消勾选
    Dim objNode As Node
    Dim arrTmp As Variant, blnCheck As Boolean
    Dim strTMp As String
    RelasCanSet = True
    Set objNode = tvwMenu(MT_关联模块).Nodes(strKey)
    If Not objNode.Checked And Not blnDblClick Or blnDblClick And objNode.Checked Then
        If objNode.Tag Like "1^*" Then
            If Not objNode.Parent Is Nothing Then
                RelasCanSet = Not objNode.Parent.Checked
                If Not RelasCanSet Then lblNotice.Caption = objNode.Text & "与上级模块" & objNode.Parent.Text & "是固定关联关系，上级模块未取消授权，该不能取消授权！"
            Else
                arrTmp = Split(mstrCurRelas, "_")
                strTMp = tvwMenu(MT_模块).Nodes(tvwMenu(MT_模块).Tag).Text
                If arrTmp(2) <> "" Then
                    blnCheck = lvwFunc(MT_模块).ListItems("F_" & arrTmp(2)).Checked
                    strTMp = strTMp & "的“" & arrTmp(2) & "”功能"
                Else
                    blnCheck = tvwMenu(MT_模块).Nodes(tvwMenu(MT_模块).Tag).Checked
                End If
                RelasCanSet = Not blnCheck
                If Not RelasCanSet Then lblNotice.Caption = objNode.Text & "与上级模块" & strTMp & "是固定关联关系，上级" & IIf(arrTmp(2) <> "", "功能", "模块") & "未取消授权，该模块不能取消授权！"
            End If
        End If
    Else
        '上级未勾选，则不能勾选，添加该规则主要是由于死循环的原因
        If Not objNode.Parent Is Nothing Then
            RelasCanSet = objNode.Parent.Checked
            If Not RelasCanSet Then lblNotice.Caption = objNode.Text & "的上级模块" & objNode.Parent.Text & "未进行授权，上级模块未授权，该模块不能进行授权！"
        End If
    End If
End Function


Private Sub SynchronizeState(Optional ByVal blnFinaly As Boolean, Optional ByVal lngTimes As Long)
'功能：将授权记录集同步到树形记录集
'参数：lngChange=授权发生变化次数
'          blnFinaly=是否是最终授权
    Dim strKey As String
    Dim arrTmp As Variant, strTMp As String
    Dim i As Long
    Dim blnHaveChange As Boolean, blnCheck As Boolean
    Dim objNode As Node
    
    On Error GoTo errH
    '由于在加载树形结构时同时修正了授权了记录集，因此需要将授权整体信息更新到树形
    With mrsModsInfo
        .Filter = "改变批次=1"
        '不存在相关模块被相关
        '存在相关模块被关联
        Do While Not .EOF
            If Not mrsRelasTree Is Nothing Then
                If !模块相关 = 1 Or !模块相关 = 3 Then
                    strKey = !系统 & "_" & !序号
                    mrsRelasTree.Filter = "MainKey='" & strKey & "' And 授权否<>" & !授权否
                    Do While Not mrsRelasTree.EOF
                        If mrsRelasTree!TreeName & "" = mstrCurRelas Then
                            blnHaveChange = True
                        End If
                        mrsRelasTree.Update Array("授权否", "授权功能"), Array(!授权否, !授权功能)
                        mrsRelasTree.MoveNext
                    Loop
                End If
            End If
            '同步到菜单
            If !系统 = mlngSys Then
                If !模块类型 = 1 Then
                    strTMp = mcllKeyModule("K_" & mlngSys & "_" & !序号)
                    arrTmp = Split(strTMp, ",")
                    For i = LBound(arrTmp) To UBound(arrTmp) '可能菜单不存在，因此错误屏蔽
                        On Error Resume Next
                        Set objNode = tvwMenu(MT_模块).Nodes(arrTmp(i))
                        If err.Number = 0 Then
                            objNode.Checked = !授权否 = 1
                            '删除被取消授权的节点
                            If objNode.Checked = False And chkOnlyShow = 1 Then
                                Do While (Not objNode.Parent Is Nothing) And objNode.Previous Is Nothing And objNode.Next Is Nothing
                                    Set objNode = objNode.Parent
                                    tvwMenu(MT_模块).Nodes.Remove objNode.Child.Key
                                Loop
                                If objNode.Children = 0 Then
                                    tvwMenu(MT_模块).Nodes.Remove objNode.Key
                                End If
                            End If
                        Else
                            '添加该实体模块节点树
                            Call AddModuleNode(!序号, !标题, 1)
                        End If
                    Next
                ElseIf chkVirtual.value Then
                    '添加虚拟模块
                    On Error Resume Next
                    Set objNode = tvwMenu(MT_模块).Nodes("M_000000_" & !序号)
                    If err.Number = 0 Then
                        objNode.Checked = !授权否 = 1
                        '删除被取消授权的节点
                        If objNode.Checked = False And chkOnlyShow = 1 Then
                            tvwMenu(MT_模块).Nodes.Remove "M_000000_" & !序号
                        End If
                    Else
                        Call AddModuleNode(!序号, !标题, 0)
                    End If
                End If
                err.Clear: On Error GoTo errH
            End If
            '修改批次等待下一次判断
            .Update "改变批次", 0
            .MoveNext
        Loop
    End With
    '关联模块的同步
    If blnHaveChange And Not blnFinaly Then
        With mrsRelasTree
            .Filter = "TreeName='" & mstrCurRelas & "'"
            .Sort = "Level,序号"
            arrTmp = Split(mstrCurRelas, "_")
            If arrTmp(2) <> "" Then
                blnCheck = lvwFunc(MT_模块).ListItems("F_" & arrTmp(2)).Checked
            Else
                blnCheck = tvwMenu(MT_模块).Nodes(tvwMenu(MT_模块).Tag).Checked
            End If
            Do While Not .EOF
                Set objNode = tvwMenu(MT_关联模块).Nodes("M_" & !Key)
                If !Level = 1 Then
                   objNode.Checked = !授权否 = 1 Or !相关类型 = 1 And blnCheck
                Else
                   objNode.Checked = (!授权否 = 1 Or !相关类型 = 1) And objNode.Parent.Checked
                End If
                '将功能更新到授权记录集
                If !授权否 = 0 And objNode.Checked Then
                    Call UpdateGrantState("M_" & !MainKey, True, !默认功能) '此时发生授权变化只记录，不做界面同步
                End If
                .MoveNext
            Loop
            mrsModsInfo.Filter = "改变批次=1"
            If Not .EOF And lngTimes < 4 Then  '再次同步
                Call SynchronizeState(False, lngTimes + 1)
            End If
        End With
    End If
    Call RefreshState
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "SynchronizeState:" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub RefreshState()
    Dim strMsg As String, strTMp As String
    Dim lngModule As Long, lngBase As Long, lngReport As Long
    Dim lngFunc As Long, lngTable As Long
    Dim lngCurSys As Long, lngVirtual As Long, lngNotVirtual As Long
    
    If mblnUnRefresh Then Exit Sub
    If msftStyle <> SFT_其他 Then
        mrsModsInfo.Filter = "授权否=1 And 系统<>0"
        lngModule = mrsModsInfo.RecordCount
        If mlngSys <> 0 Then
            mrsModsInfo.Filter = "授权否=1 And 系统=" & mlngSys
            lngCurSys = mrsModsInfo.RecordCount
            mrsModsInfo.Filter = "授权否=1 And 系统=" & mlngSys & " And 模块类型=1"
            lngNotVirtual = mrsModsInfo.RecordCount
            lngVirtual = lngCurSys - lngNotVirtual
        End If
        mrsModsInfo.Filter = "授权否=1 And 系统=0 And 序号<100"
        lngBase = mrsModsInfo.RecordCount
        mrsModsInfo.Filter = "授权否=1 And 系统=0 And 序号>=100"
        lngReport = mrsModsInfo.RecordCount
        If lngModule <> 0 Then
            strMsg = "所有系统已经授权：" & lngModule
            If lngCurSys <> 0 Then
                strMsg = strMsg & "，当前系统已授权：" & lngCurSys
                If lngNotVirtual <> 0 Then
                    strTMp = "（实体模块：" & lngNotVirtual & IIf(lngVirtual = 0, "）", "")
                End If
                If lngVirtual <> 0 Then
                    If strTMp = "" Then
                        strTMp = "（虚拟模块：" & lngVirtual & "）"
                    Else
                        strTMp = strTMp & "，虚拟模块：" & lngVirtual & "）"
                    End If
                End If
                strMsg = strMsg & strTMp
            End If
        End If
        strTMp = ""
        If lngBase <> 0 Then
            strTMp = "基础工具已授权：" & lngBase
        End If
        If lngReport <> 0 Then
            strTMp = strTMp & IIf(strTMp = "", "", "，") & "自定义报表已授权：" & lngReport
        End If
    Else
        If cmbSystem.Text = "取数函数" Then
            If mrsFunction Is Nothing Then Set mrsFunction = ReadData(RDT_Function)
            mrsFunction.Filter = "授权否=1"
            lngFunc = mrsFunction.RecordCount
            If lngFunc <> 0 Then
                strMsg = "取数函数已授权：" & lngFunc
            End If
        Else
            If mrsTable Is Nothing Then Set mrsTable = ReadData(RDT_Table)
            mrsTable.Filter = "授权否=1"
            lngTable = mrsTable.RecordCount
            If lngTable <> 0 Then
                strMsg = "基础编码已授权：" & lngTable
            End If
        End If
    End If
    If strMsg <> "" Or strTMp <> "" Then
        strMsg = strMsg & IIf(strMsg = "", "", IIf(strTMp = "", "", "，")) & strTMp
    End If
    stbThis.Panels(2).Text = strMsg
End Sub

Private Sub SetNodeExpand(ByVal tvwInput As TreeView, ByVal strKey As String)
'功能：展开某个节点
    Dim objNode As Node
    Set objNode = tvwInput.Nodes(strKey)
    objNode.Expanded = True
    Do While Not objNode.Parent Is Nothing
        objNode.Parent.Expanded = True
        Set objNode = objNode.Parent
    Loop
    tvwInput.Nodes(strKey).EnsureVisible
End Sub

Private Sub FillFunc(ByVal lvwInput As ListView)
    Dim lng系统 As Long, lng序号 As Long
    Dim arrTmp As Variant
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim strAll As String, strGant As String, strDefault As String
    Dim blnAddALL As Boolean
    Dim objNode As Node

    If tvwMenu(lvwInput.Index).Tag = "" Then Exit Sub
    Set objNode = tvwMenu(lvwInput.Index).Nodes(tvwMenu(lvwInput.Index).Tag)
    arrTmp = Split(objNode.Key, "_")
    If lvwInput.Index = MT_模块 Then
        lng系统 = mlngSys
        lng序号 = Val(arrTmp(2))
    Else
        lng系统 = Val(arrTmp(3))
        lng序号 = Val(arrTmp(4))
    End If
    If lng系统 = 0 Then
        '自定义报表或工具
        strSql = "Select 功能, 排列, 说明 From zlProgFuncs Where 系统 Is Null And 序号 = " & lng序号 & " And 功能 <> '基本'"
        Set rsTmp = gcnOracle.Execute(strSql)
    Else
        '具体的应用系统
        If gblnInIDE Then
            strSql = "Select a.功能, to_char(Nvl(a.排列,999),'000') as 排列, a.说明 " & _
                     "         From zlProgFuncs A " & _
                     "         Where A.系统 = " & lng系统 & _
                     " And A.序号 = " & lng序号 & " And A.功能 <> '基本'" & _
                     " Order By to_char(a.排列,'000')"
            Set rsTmp = gcnOracle.Execute(strSql)
        Else
            strSql = "Select Distinct 功能,to_Char(Nvl(排列,999),'000') as 排列,说明 From (Select A.功能, A.排列, A.说明 " & _
                     "         From zlProgFuncs A, Zlregfunc B " & _
                     "         Where Trunc(A.系统 / 100) = B.系统 And A.序号 = B.序号 And A.功能 = B.功能 And A.系统 = " & lng系统 & " And A.序号 = " & lng序号 & " And " & _
                     "               A.功能 <> '基本' " & _
                     "         Union " & _
                     "         Select 功能, 排列, 说明 From zlProgFuncs  a Where 功能 <> '基本' And 序号 Between 10000 And 19999 And A.系统 = " & lng系统 & " And A.序号 = " & lng序号 & _
                     "         Union " & _
                     "         Select  A.功能, A.排列, A.说明 " & _
                     "         From zlProgFuncs A, zlRPTPuts B " & _
                     "         Where A.系统 = B.系统 And A.序号 = B.程序ID And A.功能 = B.功能 And A.系统 = " & lng系统 & " And A.序号 = " & lng序号 & ")" & _
                     "  Order By to_char(排列,'000')"
            Set rsTmp = gcnOracle.Execute(strSql)
        End If
    End If
    With rsTmp
        Do While Not .EOF
            strAll = strAll & "," & !功能
            Set objItem = lvwInput.ListItems.Add(, "F_" & !功能, !功能)
            objItem.SubItems(1) = IIf(IsNull(!排列), "", !排列)
            objItem.SubItems(2) = IIf(IsNull(!说明), "", !说明)
            .MoveNext
        Loop
        If strAll <> "" Then strAll = Mid(strAll, 2)
    End With
    Call UpdateAllFunc(lng系统, lng序号, strAll)
    lvwInput.SortKey = 1
    mblnItem = True
    mrsModsInfo.Filter = "系统=" & lng系统 & " And 序号=" & lng序号
    strGant = IIf(mrsModsInfo!授权否 = 1, mrsModsInfo!授权功能, mrsModsInfo!默认功能)
    strGant = CheckFunc(lng系统, lng序号, lvwInput, strGant, strAll)
    If mrsModsInfo!授权否 = 1 Then Call UpdateGrantState("M_" & lng系统 & "_" & lng序号, True, strGant, 1)
    mblnItem = False
End Sub

Private Function CheckFunc(ByVal lng系统 As Long, ByVal lng序号 As Long, Optional ByRef lvwFunInput As ListView, Optional ByVal strGrant As String, Optional ByVal strAll As String) As String
'功能：设置功能选择状态
'参数：lng系统=系统号
'           lng序号=模块号
'           lvwFunInput=进行功能设置的列表
'           strGrant=授权功能
'           strALl=所有功能
'           blnClick=功能勾选调用
'返回：lvwFunInput为空时，返回授权功能
'说明：若lvwFunInput为空，则必须有strALl
    Dim lst As ListItem, lstTmp As ListItem
    Dim arrTmp As Variant, i As Long
    Dim strTMp As String
    
    On Error GoTo errH
    '产生功能列表的备份,并设置原功能列表勾选状态
    lvwTmp.ListItems.Clear: lvwTmp.Tag = ""
    If Not lvwFunInput Is Nothing Then
        For Each lst In lvwFunInput.ListItems
            If strGrant <> "" Then '初始设置功能勾选状态
                If InStr("," & strGrant & ",", "," & lst.Text & ",") > 0 Then
                    lst.Checked = True
                End If
            End If
            Set lstTmp = lvwTmp.ListItems.Add(, lst.Key, lst.Text)
            lstTmp.Checked = lst.Checked
            If lst.Checked Then strTMp = strTMp & "," & lst.Text
        Next
        lvwTmp.Tag = Mid(strTMp, 2)
    Else
        arrTmp = Split(strAll, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If arrTmp(i) <> "基本" Then
                Set lstTmp = lvwTmp.ListItems.Add(, "F_" & arrTmp(i), arrTmp(i))
                If strGrant <> "" Then
                    If InStr("," & strGrant & ",", "," & lstTmp.Text & ",") > 0 Then
                        lstTmp.Checked = True
                    End If
                End If
            End If
        Next
        lvwTmp.Tag = strGrant
    End If
    Call CheckFuncRelas(lng系统, lng序号, lvwTmp)
    CheckFunc = lvwTmp.Tag
    '从备份状态回复原列表
    If Not lvwFunInput Is Nothing Then
        For Each lstTmp In lvwTmp.ListItems
            lvwFunInput.ListItems(lstTmp.Key).Checked = lstTmp.Checked
        Next
    End If
    Exit Function
errH:
    MsgBox "CheckFunc:" & err.Description, vbInformation, Me.Caption
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub CheckFuncRelas(ByVal lng系统 As Long, ByVal lng序号 As Long, ByRef lvwInput As ListView)
'功能：检查权限关系
'参数：lng系统=系统号
'           lng序号=模块号
'           lvwInput=进行设置的的列表
'返回：功能的授权情况
    Dim i As Integer, intUpdate As Integer
    Dim lst As ListItem
    
    On Error GoTo errHand
    mintUpdate = 1
    Do While mintUpdate >= 1
        intUpdate = 0
        For Each lst In lvwInput.ListItems
            '主从关系
            '检查每个选择的功能是否符合权限间的关系,只处理项目为主项的情况,项目为子项的,由其主项的状态来调整。
            mrsRelas.Filter = "系统 = " & lng系统 & " And 序号 = " & lng序号 & " And 功能 = '" & lst.Text & "'"
            Do Until mrsRelas.EOF
                mrsGroup.Filter = "系统 = " & lng系统 & " And 序号 = " & lng序号 & " And 组号 = " & mrsRelas!组号
                If Not mrsGroup.EOF Then
                    mintUpdate = 1
                    Call setState(mrsGroup, mrsRelas!功能 & "_" & mrsRelas!组号, lst.Checked, lvwInput)
                    If mintUpdate > 0 Then
                        intUpdate = intUpdate + 1
                    End If
                End If
                mrsRelas.MoveNext
            Loop
            '互斥关系(初次选模块时起作用),
            mrsRelExcl.Filter = "系统 = " & lng系统 & " And 序号 = " & lng序号 & " And 功能 = '" & lst.Text & "'"
            If Not mrsRelExcl.EOF Then
                mrsGroup.Filter = "系统 = " & lng系统 & " And 序号 = " & lng序号 & " And 组号 = " & mrsRelExcl!组号
                If Not mrsGroup.EOF Then
                    mintUpdate = 1
                    Call setState(mrsGroup, mrsRelExcl!功能 & "_" & mrsRelExcl!组号, lst.Checked, lvwInput)
                    If mintUpdate > 0 Then
                        intUpdate = intUpdate + 1
                    End If
                End If
            End If
        Next
        mintUpdate = intUpdate
    Loop
    Exit Sub
errHand:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbQuestion, gstrSysName
End Sub

Private Sub setState(ByVal rsTmp As ADODB.Recordset, ByVal strKey As String, ByVal blnCheck As Boolean, ByRef lvwInput As ListView)
    '调整一组主从权限的状态，
    Dim objRelas As clsRelas, lst As ListItem, intUpdate As Integer
    Dim objGroup As clsRGroup, strPrivs As String
    Set objGroup = New clsRGroup
    
    Do Until rsTmp.EOF
        Set objRelas = New clsRelas
        objRelas.分组 = rsTmp!组号
        objRelas.功能 = rsTmp!功能
        objRelas.关系 = rsTmp!关系
        objRelas.主功能 = rsTmp!主项
        objRelas.主功能关系 = rsTmp!主项关系
        objRelas.Checked = InStr("," & lvwInput.Tag & ",", "," & rsTmp!功能 & ",") > 0
        objRelas.Key = rsTmp!功能 & "_" & rsTmp!组号
        Call objGroup.Add(objRelas, objRelas.Key)
        rsTmp.MoveNext
    Loop

    Call objGroup.RelasCheck(strKey, blnCheck)
    intUpdate = mintUpdate
    For Each lst In lvwInput.ListItems
        For Each objRelas In objGroup
            If objRelas.功能 = lst.Text And lst.Checked <> objRelas.Checked Then
               If lst.Checked <> objRelas.Checked Then
                    lst.Checked = objRelas.Checked
                   mintUpdate = mintUpdate + 1
               End If
            End If
        Next
    Next
    
    If intUpdate = mintUpdate Then
        mintUpdate = mintUpdate - 1
        If mintUpdate < 0 Then mintUpdate = 0
    End If
    
    strPrivs = ""
    For Each lst In lvwInput.ListItems
        If lst.Checked = True Then
            strPrivs = strPrivs & "," & lst.Text
        End If
    Next
    If strPrivs <> "" Then strPrivs = Mid(strPrivs, 2)
    lvwInput.Tag = strPrivs
End Sub

Private Function setExcl(ByVal rsTmp As ADODB.Recordset, ByVal strKey As String, ByVal blnCheck As Boolean, ByRef lvwInput As ListView, ByVal strGrant As String) As String
    '调整一组互斥权限的状态，
    Dim objRelas As clsRelas, lst As ListItem, intUpdate As Integer
    Dim objGroup As clsRGroup, strPrivs As String
    Set objGroup = New clsRGroup
    
    Do Until rsTmp.EOF
        Set objRelas = New clsRelas
        objRelas.分组 = rsTmp!组号
        objRelas.功能 = rsTmp!功能
        objRelas.关系 = rsTmp!关系
        objRelas.主功能 = rsTmp!主项
        objRelas.主功能关系 = rsTmp!主项关系
        objRelas.Checked = InStr("," & strGrant & ",", "," & rsTmp!功能 & ",") > 0
        objRelas.Key = rsTmp!功能 & "_" & rsTmp!组号
        Call objGroup.Add(objRelas, objRelas.Key)
        rsTmp.MoveNext
    Loop

    Call objGroup.RelasCheck(strKey, blnCheck)
    For Each lst In lvwInput.ListItems
        For Each objRelas In objGroup
            If objRelas.功能 = lst.Text And lst.Checked <> objRelas.Checked Then
                lvwInput.ListItems(lst.Index).Checked = objRelas.Checked
            End If
        Next
    Next
    
    strPrivs = ""
    For Each lst In lvwInput.ListItems
        If lst.Checked = True Then
            strPrivs = strPrivs & "," & lst.Text
        End If
    Next
    If strPrivs <> "" Then strPrivs = Mid(strPrivs, 2)
    setExcl = strPrivs
End Function

Private Sub UpdateAllFunc(ByVal lngSys As Long, ByVal lngModule As Long, ByVal strFuncs As String)
'可能有些是报表功能，因此需要跟新所有功能
    mrsModsInfo.Filter = "系统=" & lngSys & " And 序号=" & lngModule
    If Not mrsModsInfo.EOF Then
        mrsModsInfo.Update "所有功能", strFuncs
    End If
End Sub

Private Sub UpdateGrantState(Optional ByVal strKey As String, Optional ByVal blnGrant As Boolean = True, Optional ByVal strFuncs As String, Optional ByVal intGrantType As Integer)
'功能：更新模块授权信息
'参数：strKey=系统_模块,为空时表示将模块授权统计信息更新到模块授权明细信息记录集中
'          strFuncs=授权功能
'          blnGrant=True-授权,False-取消授权
'          intGrantType=0-默认被动授权,1-主动授权
    Dim arrTmp As Variant
    Dim strGrant As String
    Dim blnChange As Boolean
    Dim strObj As String, i As Long
    
    On Error GoTo errH
    '将模块授权的整体信息更新到模块授权明细信息记录集中
    If strKey = "" Then
        Call AdjustRelasTree '固定关联自动授权处理
        With mrsModsInfo
            .Filter = "授权改变=1"
            .Sort = "系统,序号"
            Do While Not .EOF
                
                mrsModule.Filter = "系统=" & !系统 & " And 序号=" & !序号
                If !授权否 = 0 Then '取消授权
                    Do While Not mrsModule.EOF
                        mrsModule.Update "授权否", 0
                        mrsModule.MoveNext
                    Loop
                Else '修正授权
                    '权限关系检查
                    strGrant = CheckFunc(!系统, !序号, , !授权功能, !所有功能)
                    Do While Not mrsModule.EOF
                        '该模块授权，则基本功能一定授权
                        If mrsModule!功能 & "" = "基本" Then
                            mrsModule.Update "授权否", 1
                        ElseIf InStr("," & strGrant & ",", "," & mrsModule!功能 & ",") > 0 Then
                            mrsModule.Update "授权否", 1
                        Else
                            mrsModule.Update "授权否", 0
                        End If
                        mrsModule.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End With
        Exit Sub
    End If
    '对某一个模块表或函数取消授权或授权
    arrTmp = Split(strKey, "_")
    strObj = arrTmp(2)
    If UBound(arrTmp) > 2 Then
        strObj = Mid(strKey, Len(arrTmp(1)) + 4)
    End If
    Select Case arrTmp(0)
        Case "M"
            With mrsModsInfo
                .Filter = "系统=" & arrTmp(1) & " And 序号=" & strObj
                If .EOF Then Exit Sub '没有该模块，则退出
                If blnGrant Then
                    If !授权否 = 0 Then
                        strGrant = IIf(strFuncs = "" And intGrantType = 0, !默认功能 & "", strFuncs)
                        .Update Array("授权改变", "改变批次", "授权否", "授权功能"), Array(1, 1, 1, strGrant)
                    Else
                        '已经授权再次进行默认授权时则授予授权功能，否则授予当前功能
                        strGrant = IIf(strFuncs = "" And intGrantType = 0, !授权功能 & "", strFuncs)
                        blnChange = (strFuncs <> !授权功能 & "") '本次授权功能是否发生变化
                        .Update Array("授权改变", "改变批次", "授权否", "授权功能"), Array(IIf(blnChange Or !授权改变 <> 0, 1, 0), IIf(blnChange Or !改变批次 <> 0, 1, 0), 1, strGrant)
                    End If
                Else
                    If !授权否 = 1 Then '取消授权
                        .Update Array("授权改变", "改变批次", "授权否", "授权功能"), Array(1, 1, 0, "")
                    End If
                End If
            End With
        Case "T"
            mrsTable.Filter = "表名='" & strObj & "' And 系统=" & arrTmp(1)
            mrsTable.Update "授权否", IIf(blnGrant, 1, 0)
        Case "F"
            mrsFunction.Filter = "函数名='" & strObj & "' And 系统=" & arrTmp(1)
            mrsFunction.Update "授权否", IIf(blnGrant, 1, 0)
    End Select
    Call RefreshState
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "UpdateGrantState:" & err.Description, vbInformation, Me.Caption
End Sub

Private Function GetRelasTree(Optional ByVal blnLast As Boolean)
    Dim rsRelasTmp As ADODB.Recordset
    Dim rsRelasTree As ADODB.Recordset
    Dim blnDo As Boolean, blnGrant As Boolean
    Dim strPreTree As String
    Dim objNode As Node

    On Error GoTo errH
    If mrsModRelas Is Nothing Then Set mrsModRelas = ReadData(RDT_ModRelas)
    mrsModRelas.Filter = "": mrsModRelas.Sort = "系统 , 模块, 功能, 相关系统, 相关模块,相关信息"
    Set rsRelasTmp = CopyNewRec(mrsModRelas)
    Set rsRelasTree = CopyNewRec(mrsModRelas, True, "相关系统 系统,相关模块 模块,相关类型,相关信息,缺省值,标题,标题 上级标题", Array("TreeName", adVarChar, 100, Empty, _
                                                     "TreeMain", adVarChar, 100, Empty, "TreeFunc", adVarChar, 100, Empty, "PreKey", adVarChar, 100, Empty, "MainPreKey", adVarChar, 100, Empty, _
                                                     "Key", adVarChar, 100, Empty, "MainKey", adVarChar, 100, Empty, "Level", adInteger, 1, 0, "序号", adInteger, 1, 0, "授权否", adInteger, 1, 0, _
                                                     "授权功能", adVarChar, 2000, Empty, "所有功能", adVarChar, 2000, Empty, "默认功能", adVarChar, 2000, Empty))
    mrsModRelas.Filter = "": mrsModRelas.Sort = "系统 , 模块, 功能, 相关系统, 相关模块,相关信息"
    With mrsModRelas
        Do While Not .EOF
            blnDo = False
            mrsModsInfo.Filter = "系统=" & !相关系统 & " And 序号=" & !相关模块
            If Not mrsModsInfo.EOF Then
                mrsModsInfo.Filter = "系统=" & !系统 & " And 序号=" & !模块
                blnDo = Not mrsModsInfo.EOF And strPreTree <> !系统 & "_" & !模块 & "_" & !功能
            End If
            If blnDo Then
                strPreTree = !系统 & "_" & !模块 & "_" & !功能
                blnGrant = mrsModsInfo!授权否 = 1
                If blnGrant Then '判断树形是否可用
                    If !功能 & "" <> "" Then blnGrant = InStr("," & mrsModsInfo!授权功能 & ",", "," & !功能 & ",") > 0
                End If
                Set objNode = tvwModRelas.Nodes.Add(, , strPreTree, "【(" & Format(!模块 & "", "000000") & ")" & !标题 & "】" & !功能, "分类")
                objNode.Checked = blnGrant
                If FillRelasTreeRec(!系统 & "_" & !模块 & "_" & !功能, !系统 & "_" & !模块, !功能 & "", rsRelasTree, rsRelasTmp) Then
                    mrsModsInfo.Filter = "系统=" & !系统 & " And 序号=" & !模块
                    mrsModsInfo.Update "模块相关", GetCurType(Val(mrsModsInfo!模块相关 & ""), True)
                End If
            End If
            .MoveNext
        Loop
    End With
    Set mrsRelasTree = rsRelasTree
'    '记录集同步，将授权记录集同步到树形
    Call SynchronizeState(Not blnLast)
    Exit Function
errH:
    MsgBox "GetRelasTree:" & err.Description, vbInformation, Me.Caption
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function FillRelasTreeRec(ByVal strTreeName As String, ByVal strTreeMain As String, ByVal strTreeFunc As String, ByRef rsTree As ADODB.Recordset, _
                                                        ByVal rsReals As ADODB.Recordset, Optional ByRef cllNodes As Collection, _
                                                        Optional ByVal strKey As String, Optional ByVal lngLevel As Long = 1, Optional ByRef lng序号 As Long) As Boolean
'功能：填充相关模块
'参数：strKey=按格式进行特殊处理
'                       格式1：系统_模块_功能-加载该模块（功能为空时）或模块功能的所有相关模块
'                       格式2：系统_模块_相关系统_相关模块-加载该模块的子级相关模块
    Dim arrTmp As Variant
    Dim strFilter As String
    Dim strMainKey As String, strPreKey As String, strMainPreKey As String, strTMp As String
    Dim cllNodesKey As New Collection
    Dim Item As Variant
    Dim objNode As Node
    Dim intGant As Integer, strGrant As String, strPreGrant As String
    
    On Error GoTo errH
    If lngLevel = 5 Then Exit Function '只加载4级节点
    If lngLevel = 1 And cllNodes Is Nothing Then
        Set cllNodes = New Collection '用来判断节点是否存在
        lng序号 = 0
    End If
    arrTmp = Split(IIf(strKey = "", strTreeName, strKey), "_")
    If UBound(arrTmp) <= 2 Then '加载所有相关模块
        strFilter = "系统=" & arrTmp(0) & " And 模块=" & arrTmp(1) & IIf(arrTmp(2) <> "", " And 功能='" & arrTmp(2) & "' And 类型=0", " And 类型=1")
        strMainPreKey = arrTmp(0) & "_" & arrTmp(1)
        strPreKey = ""
    Else '加载子级相关模块
        strFilter = "系统=" & arrTmp(2) & " And 模块=" & arrTmp(3) & " And 类型=1"
        strMainPreKey = arrTmp(2) & "_" & arrTmp(3)
        strPreKey = strKey
    End If
    '获取上一级授权
    arrTmp = Split(strMainPreKey, "_")
    mrsModsInfo.Filter = "系统=" & arrTmp(0) & " And 序号=" & arrTmp(1)
    strPreGrant = mrsModsInfo!授权功能 & ""
    
    With rsReals
        .Filter = strFilter
        Do While Not .EOF
            mrsModsInfo.Filter = "系统=" & !相关系统 & " And 序号=" & !相关模块
            If Not mrsModsInfo.EOF Then '该模块可以授权
                strMainKey = !相关系统 & "_" & !相关模块
                strTMp = !系统 & "_" & !模块 & "_" & !相关系统 & "_" & !相关模块
                On Error Resume Next
                cllNodesKey.Add strTMp
                cllNodes.Add "1", strTMp '判断是否存在该节点
                If err.Number = 0 Then
                    On Error GoTo errH
                    lng序号 = lng序号 + 1
                    If lngLevel = 1 Then
                        Set objNode = tvwModRelas.Nodes.Add(strTreeName, 4, strTreeName & "M_" & strTMp, "【" & Format(!相关模块 & "", "000000") & "】" & mrsModsInfo!标题, IIf(!相关类型 = 1, "Fixed", "Optional"))
                    Else
                        Set objNode = tvwModRelas.Nodes.Add(strTreeName & "M_" & strKey, 4, strTreeName & "M_" & strTMp, "【" & Format(!相关模块 & "", "000000") & "】" & mrsModsInfo!标题, IIf(!相关类型 = 1, "Fixed", "Optional"))
                    End If
                    If lngLevel = 1 Then
                        objNode.Checked = mrsModsInfo!授权否 = 1 Or !相关类型 = 1 And objNode.Parent.Checked
                    Else
                        objNode.Checked = (mrsModsInfo!授权否 = 1 Or !相关类型 = 1) And objNode.Parent.Checked
                    End If
                    intGant = IIf(objNode.Checked, 1, 0)
                    If mrsModsInfo!授权否 = 0 And objNode.Checked Then
                        strGrant = GetGrantByRelasInfo(strPreGrant, mrsModsInfo!默认功能 & "", !相关信息)
                    ElseIf mrsModsInfo!授权否 = 1 Then
                        strGrant = mrsModsInfo!授权功能 & ""
                    End If
                    objNode.Tag = !相关类型 & "^" & !相关信息
                    rsTree.AddNew Array("TreeName", "TreeMain", "TreeFunc", "PreKey", "MainPreKey", "Key", "MainKey", "Level", "序号", "系统", "模块", "相关类型", "相关信息", "缺省值", "标题", "上级标题", "授权否", "授权功能", "所有功能", "默认功能"), _
                                    Array(strTreeName, strTreeMain, strTreeFunc, strPreKey, strMainPreKey, strTMp, strMainKey, lngLevel, lng序号, !相关系统, !相关模块, !相关类型, !相关信息, !缺省值, mrsModsInfo!标题, !标题, intGant, strGrant, mrsModsInfo!所有功能, mrsModsInfo!默认功能)
                    mrsModsInfo.Update "模块相关", GetCurType(Val(mrsModsInfo!模块相关 & ""))
                    If mrsModsInfo!授权否 = 0 And objNode.Checked Then
                        '更新授权
                        Call UpdateGrantState("M_" & strMainKey, True, strGrant)
                    End If
                    '位置1
                Else
                    err.Clear: On Error GoTo errH
                End If
            End If
            .MoveNext
        Loop
    End With
    '递归加载节点，该断代码不能移动到【位置1】，因为直接关联优先加载原则
    For Each Item In cllNodesKey
        Call FillRelasTreeRec(strTreeName, strTreeMain, strTreeFunc, rsTree, rsReals, cllNodes, Item, lngLevel + 1, lng序号)
    Next
    If lngLevel = 1 Then
        FillRelasTreeRec = cllNodes.Count <> 0
    End If
    Exit Function
errH:
    MsgBox "FillRelasTreeRec:" & err.Description, vbInformation, Me.Caption
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetCurType(ByVal intOldType As Integer, Optional ByVal blnHaveTree As Boolean) As Integer
'功能：获取一个模块的相关情况
    Dim intReturn As Integer
    '模块相关=0=不存在相关模块不被关联
    '               1=不存在相关模块被关联
    '               2=存在相关模块不被关联
    '               3=存在相关模块被关联
    Select Case intOldType
        Case 0
            intReturn = IIf(blnHaveTree, 2, 1)
        Case 1
            intReturn = IIf(blnHaveTree, 3, 1)
        Case 2
            intReturn = IIf(blnHaveTree, 2, 3)
        Case 3
            intReturn = 3
    End Select
    GetCurType = intReturn
End Function

Private Sub FillRelasModule(ByVal strTreeName As String)
'功能：填充相关模块
'参数：strTreeName=系统_模块_功能-加载该模块（功能为空时）或模块功能的所有相关模块
    Dim objNode As Node, strCaption As String
    Dim blnTreeGrant As Boolean
    Dim arrTmp As Variant
    Dim strTMp As String
    Dim blnUpdate As Boolean
    Dim blnNew As Boolean
    
    On Error GoTo errH
    If mrsModRelas Is Nothing Then Set mrsModRelas = ReadData(RDT_ModRelas)
    arrTmp = Split(strTreeName, "_")
    If UBound(arrTmp) <> 2 Then Exit Sub
    '模块没有相关模块，则不用判断
    If arrTmp(2) <> "" And tvwMenu(MT_关联模块).Nodes.Count = 0 Then
        mstrCurRelas = strTreeName
        Exit Sub
    ElseIf arrTmp(2) <> "" Then '先清空节点颜色
        For Each objNode In tvwMenu(MT_关联模块).Nodes
            objNode.BackColor = &HEFF0E0: objNode.Bold = False
        Next
    Else
        If Not mstrCurRelas Like arrTmp(0) & "_" & arrTmp(1) & "_*" Then
            blnNew = True
            Call ClearFace(CT_关联模块)
        End If
    End If
    mstrCurRelas = strTreeName
    If mrsRelasTree Is Nothing Then
        mrsModRelas.Filter = "系统=" & arrTmp(0) & " And 模块=" & arrTmp(1) & IIf(arrTmp(2) = "", "", " And 功能='" & arrTmp(2) & "'")
        If mrsModRelas.RecordCount = 0 Then mstrCurRelas = strTreeName: Exit Sub
        Call GetRelasTree
    End If
    With mrsRelasTree
        .Filter = "TreeName='" & strTreeName & "'"
        .Sort = "Level,序号"
        If blnNew Then
            If Not .EOF Then
                mrsModsInfo.Filter = "系统=" & arrTmp(0) & " And 序号=" & arrTmp(1)
                If mrsModsInfo.EOF Then Exit Sub
                blnTreeGrant = mrsModsInfo!授权否
                If arrTmp(2) <> "" Then
                    blnTreeGrant = InStr("," & mrsModsInfo!授权功能 & ",", "," & arrTmp(2) & ",")
                End If
            End If
            '模块关联模块加载关联模块
            Do While Not .EOF
                strCaption = "【" & Format(!模块 & "", "000000") & "】" & !标题
                If !Level = 1 Then
                    Set objNode = tvwMenu(MT_关联模块).Nodes.Add(, , "M_" & !Key, strCaption, IIf(!相关类型 = 1, "Fixed", "Optional"))
                    objNode.Checked = !授权否 = 1 Or !相关类型 = 1 And blnTreeGrant
                Else
                    Set objNode = tvwMenu(MT_关联模块).Nodes.Add("M_" & !PreKey, tvwChild, "M_" & !Key, strCaption, IIf(!相关类型 = 1, "Fixed", "Optional"))
                    objNode.Checked = (!授权否 = 1 Or !相关类型 = 1) And objNode.Parent.Checked
                End If
                objNode.Tag = !相关类型 & "^" & !相关信息
                objNode.BackColor = &HEFF0E0
                If strTMp = "" Then strTMp = objNode.Key
                If objNode.Checked And tvwMenu(MT_关联模块).Tag = "" Then tvwMenu(MT_关联模块).Tag = objNode.Key
                '将功能更新到授权记录集
                If !授权否 = 0 And objNode.Checked Then
                    Call UpdateGrantState("M_" & !MainKey, True, !默认功能)
                    blnUpdate = True
                End If
                .MoveNext
            Loop
            
            If tvwMenu(MT_关联模块).Tag = "" And strTMp <> "" Then
                tvwMenu(MT_关联模块).Tag = strTMp
            End If
            '获取第一个选择的节点（若没有选择，则选中第一个节点）
            If tvwMenu(MT_关联模块).Tag <> "" Then
                Set objNode = tvwMenu(MT_关联模块).Nodes(tvwMenu(MT_关联模块).Tag)
                Call SetNodeExpand(tvwMenu(MT_关联模块), objNode.Key)
                tvwMenu(MT_关联模块).Tag = "" '清空Tag,在下面事件中会重新设置，否则不会加载功能
                Call tvwMenu_NodeClick(MT_关联模块, objNode)
            End If
        '    '记录集同步，将授权记录集同步到树形
            If blnUpdate Then Call SynchronizeState
        Else
            Do While Not .EOF
                tvwMenu(MT_关联模块).Nodes("M_" & !Key).Bold = arrTmp(2) <> ""
                .MoveNext
            Loop
        End If
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "FillRelasModule:" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub AdjustRelasTree()
'功能：最后点击授权时使用，用来标注关联模块需要授权
    Dim lngTimes As Long
    Dim blnDo As Boolean, strCurTree As String
    Dim arrTmp As Variant, blnTreeGrant As Boolean
    Dim objNode As Node
    Dim blnUpdate As Boolean
    
    On Error GoTo errH
    '同步授权状态
    SynchronizeState (True)
    '将相关模块树形同步到树形控件
    '如果授权发生变化或者发生两次同步，则终止该同步
    lngTimes = 1: blnDo = True
    If mrsRelasTree Is Nothing Then Call GetRelasTree
    With mrsRelasTree
        Do While blnDo
            .Filter = ""
            .Sort = "TreeName,Level,序号"
            strCurTree = ""
            Do While Not .EOF
                If strCurTree <> !TreeName & "" Then
                    strCurTree = !TreeName & ""
                    arrTmp = Split(strCurTree, "_")
                    mrsModsInfo.Filter = "系统=" & arrTmp(0) & " And 序号=" & arrTmp(1)
                    blnTreeGrant = mrsModsInfo!授权否 = 1
                    If arrTmp(2) <> "" Then
                        blnTreeGrant = InStr("," & mrsModsInfo!授权功能 & ",", "," & arrTmp(2) & ",")
                    End If
                    Set objNode = tvwModRelas.Nodes(strCurTree)
                    objNode.Checked = blnTreeGrant
                End If
                Set objNode = tvwModRelas.Nodes(strCurTree & "M_" & !Key)
                objNode.Checked = (!授权否 = 1 Or !相关类型 = 1) And objNode.Parent.Checked
                If !授权否 = 0 And objNode.Checked Then
                    Call UpdateGrantState("M_" & !MainKey, True, !默认功能)
                End If
                .MoveNext
            Loop
            mrsModsInfo.Filter = "改变批次=1"
            blnDo = Not mrsModsInfo.EOF And lngTimes < 3
            If Not mrsModsInfo.EOF Then
                '同步授权状态
                SynchronizeState (True)
            End If
            lngTimes = lngTimes + 1
        Loop
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "AdjustRelasTree:" & err.Description, vbInformation, Me.Caption
End Sub

Private Function GetGrantByRelasInfo(ByVal strPreGrant As String, ByVal strGrtant As String, ByVal strRelasInfo As String, Optional ByVal blnGetDefault As Boolean) As String
'功能：根据当前授权以及相关信息获取应该授权的功能
'参数：
'         strPreGrant=上级授权
'         strGrtant=当前授权功能
'         strRelasInfo=当前节点与上级借点的相关信息
' 返回：修正之后的授权
    Dim arrRelasInfo As Variant, arrRelasTmp As Variant
    Dim i As Long, strReturn As String
    
    strReturn = strGrtant
    arrRelasInfo = Split(strRelasInfo, ";")
    For i = LBound(arrRelasInfo) To UBound(arrRelasInfo)
        arrRelasTmp = Split(arrRelasInfo(i), ",") '功能,相关功能,相关类型,缺省值
        If arrRelasTmp(0) = "" And arrRelasTmp(1) = "" Then '模块相关不做处理
        ElseIf arrRelasTmp(0) <> "" And arrRelasTmp(1) = "" Then '功能模块相关不做处理
        ElseIf arrRelasTmp(0) = "" And arrRelasTmp(1) <> "" Then '模块功能相关
            '固定关联，没有授权，需要授权
            If arrRelasTmp(2) = 1 And InStr("," & strGrtant & ",", "," & arrRelasTmp(1) & ",") = 0 Then
                    strReturn = IIf(strReturn = "", arrRelasTmp(1), strReturn & "," & arrRelasTmp(1))
            End If
        Else '功能功能相关需要进行判别
            If InStr("," & strPreGrant & ",", "," & arrRelasTmp(0) & ",") > 0 Then
                '固定关联，没有授权，需要授权
                If arrRelasTmp(2) = 1 And InStr("," & strGrtant & ",", "," & arrRelasTmp(1) & ",") = 0 Then
                        strReturn = IIf(strReturn = "", arrRelasTmp(1), strReturn & "," & arrRelasTmp(1))
                End If
            End If
        End If
    Next
    GetGrantByRelasInfo = strReturn
End Function





