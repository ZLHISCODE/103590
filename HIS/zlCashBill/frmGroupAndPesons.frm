VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "codejock.dockingpane.9600.ocx"
Begin VB.Form frmGroupAndPesons 
   Caption         =   "缴款人员分组"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmGroupAndPesons.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11745
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList igl16 
      Left            =   6075
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":058A
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":0B24
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":10BE
            Key             =   "Group"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPersons 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   375
      ScaleHeight     =   5685
      ScaleWidth      =   7185
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3360
      Width           =   7185
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   555
         TabIndex        =   13
         Top             =   435
         Width           =   3870
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   555
         TabIndex        =   11
         Top             =   75
         Width           =   3900
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "增加"
         Height          =   300
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   60
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "移除"
         Height          =   300
         Index           =   4
         Left            =   5175
         TabIndex        =   15
         Top             =   60
         Width           =   570
      End
      Begin MSComctlLib.ListView lvwPerson 
         Height          =   6435
         Left            =   0
         TabIndex        =   16
         Top             =   915
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   11351
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "igl32"
         SmallIcons      =   "igl16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "姓名"
            Object.Tag             =   "姓名"
            Text            =   "姓名"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "编号"
            Object.Tag             =   "编号"
            Text            =   "编号"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Key             =   "简码"
            Object.Tag             =   "简码"
            Text            =   "简码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "出生日期"
            Object.Tag             =   "出生日期"
            Text            =   "出生日期"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "性别"
            Object.Tag             =   "性别"
            Text            =   "性别"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "民族"
            Object.Tag             =   "民族"
            Text            =   "民族"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "办公室电话"
            Object.Tag             =   "办公室电话"
            Text            =   "办公室电话"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "电子邮件"
            Object.Tag             =   "电子邮件"
            Text            =   "电子邮件"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "管理职务"
            Object.Tag             =   "管理职务"
            Text            =   "管理职务"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原组"
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   12
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成员"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   135
         Width           =   360
      End
   End
   Begin VB.PictureBox picGroup 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   7860
      Left            =   465
      ScaleHeight     =   7860
      ScaleWidth      =   4935
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   4935
      Begin VB.CommandButton cmdFucn 
         Caption         =   "删除"
         Height          =   300
         Index           =   2
         Left            =   3900
         TabIndex        =   8
         Top             =   855
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "修改"
         Height          =   300
         Index           =   1
         Left            =   3285
         TabIndex        =   7
         Top             =   855
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "增加"
         Height          =   300
         Index           =   0
         Left            =   2685
         TabIndex        =   6
         Top             =   855
         Width           =   570
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   645
         TabIndex        =   3
         Top             =   465
         Width           =   3900
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   645
         TabIndex        =   5
         Top             =   855
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   645
         TabIndex        =   1
         Top             =   75
         Width           =   3900
      End
      Begin MSComctlLib.ListView lvwGroups 
         Height          =   6510
         Left            =   60
         TabIndex        =   9
         Top             =   1275
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   11483
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "igl32"
         SmallIcons      =   "igl16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "组名称"
            Object.Tag             =   "组名称"
            Text            =   "组名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "负责人"
            Object.Tag             =   "负责人"
            Text            =   "负责人"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "说明"
            Object.Tag             =   "说明"
            Text            =   "说明"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   2
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "负责人"
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   915
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "组名称"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   135
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   8085
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmGroupAndPesons.frx":1658
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
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
   Begin MSComctlLib.ImageList igl32 
      Left            =   6930
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":1EEC
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":27C6
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":30A0
            Key             =   "Group"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStructure 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   4260
      ScaleHeight     =   5685
      ScaleWidth      =   7185
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1215
      Width           =   7185
      Begin MSComctlLib.ListView lvwStructure 
         Height          =   6435
         Left            =   0
         TabIndex        =   24
         Top             =   555
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   11351
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "igl32"
         SmallIcons      =   "igl16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "姓名"
            Object.Tag             =   "姓名"
            Text            =   "姓名"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "编号"
            Object.Tag             =   "编号"
            Text            =   "编号"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Key             =   "简码"
            Object.Tag             =   "简码"
            Text            =   "简码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "出生日期"
            Object.Tag             =   "出生日期"
            Text            =   "出生日期"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "性别"
            Object.Tag             =   "性别"
            Text            =   "性别"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "民族"
            Object.Tag             =   "民族"
            Text            =   "民族"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "办公室电话"
            Object.Tag             =   "办公室电话"
            Text            =   "办公室电话"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "电子邮件"
            Object.Tag             =   "电子邮件"
            Text            =   "电子邮件"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "管理职务"
            Object.Tag             =   "管理职务"
            Text            =   "管理职务"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "移除"
         Height          =   300
         Index           =   6
         Left            =   5160
         TabIndex        =   23
         Top             =   45
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "增加"
         Height          =   300
         Index           =   5
         Left            =   4560
         TabIndex        =   22
         Top             =   60
         Width           =   570
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   555
         TabIndex        =   21
         Top             =   75
         Width           =   3900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "组长"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   135
         Width           =   360
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmGroupAndPesons.frx":397A
      Left            =   555
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmGroupAndPesons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************************************************************
'功能:缴款人员分组
'编制:刘兴洪
'日期:2010-11-23 15:42:13
'说明:
'    33633
'********************************************************************************************************************************************
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private WithEvents mfrmFilter As frmBillInFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mlngModul As Long, mstrPrivs As String
Private mblnFirst As Boolean  '第一次加载窗体
Private mstrKey As String, mstrPreGroupKey As String    '上一次的记录
Private Enum mPaneID
    Pane_Group = 1    '
    Pane_Persons = 3
    Pane_Structure = 2
End Enum
Private mblnItem As Boolean  '为真表示单击到ListView某一项上
Private mintSucess As Integer '>0表示只少更改了一项值的
Private mintGroupColumn As Integer, mintPersonColumn As Integer
Private mintStructureColumn As Integer
Private mblnEdit As Boolean '是否处于编辑状态
Private mblnStartDrop As Boolean '开始拖动
Private mstrSelect As String
Private mblnReSel As Boolean '是否重新选中
Private mblnItemClick As Boolean  '是否点到项目
Private Enum mTxtIdx
    idx_组名称 = 0
    idx_负责人 = 1
    idx_组说明 = 2
    idx_成员 = 3
    idx_原组 = 4
    idx_组长 = 6
End Enum
Private Enum mCmdIdx
    idx_组增加 = 0
    idx_组修改 = 1
    idx_组删除 = 2
    idx_成员增加 = 3
    idx_成员移出 = 4
    idx_组长增加 = 5
    idx_组长删除 = 6
End Enum
Private Sub cmdFucn_Click(Index As Integer)
    Select Case Index
    Case mCmdIdx.idx_组增加
        Call AddGroups(0)
    Case mCmdIdx.idx_组修改
        Call AddGroups(1)
    Case mCmdIdx.idx_组删除
        Call DeleteGroup
    Case mCmdIdx.idx_成员增加
        Call AddPerson
    Case mCmdIdx.idx_成员移出
        If lvwPerson.SelectedItem Is Nothing Then Exit Sub
        If lvwGroups.SelectedItem Is Nothing Then Exit Sub
        Call PersonFromGroupToOtherGroup(Mid(lvwPerson.SelectedItem.Key, 2), Trim(txtEdit(mTxtIdx.idx_成员)), _
            Val(txtEdit(mTxtIdx.idx_原组).Tag), Trim(txtEdit(mTxtIdx.idx_原组)))
    Case mCmdIdx.idx_组长增加
        Call AddStructure
    Case mCmdIdx.idx_组长删除
        If lvwStructure.SelectedItem Is Nothing Then Exit Sub
        If lvwGroups.SelectedItem Is Nothing Then Exit Sub
        Call DeleteStructure(Mid(lvwStructure.SelectedItem.Key, 2), Trim(lvwStructure.SelectedItem.Text))
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mstrPreGroupKey = ""
    Call LoadGroups
    RestoreListViewState lvwPerson, Me.Name, ""
    RestoreListViewState lvwGroups, Me.Name, ""
    RestoreListViewState lvwStructure, Me.Name, ""
End Sub
Public Function ShowGroups(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示或设置成员分组(程序入口)
    '入参:frmMain-窗体
    '       lngModule-模块号
    '       strPrivs-权限串
    '出参:
    '返回:如果操作了一项成功的,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-23 15:45:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModul = lngModule: mstrPrivs = strPrivs: mintSucess = 0
    Me.Show 1, frmMain
    ShowGroups = mintSucess > 0
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    
    mblnFirst = True
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars '初始菜单及工具栏
    Call InitPanel
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    mblnEdit = zlStr.IsHavePrivs(mstrPrivs, "成员分组")
    Call SetCtrlVisible
End Sub
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的visible属性
    '编制:刘兴洪
    '日期:2010-11-23 16:38:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Err = 0: On Error Resume Next
    For i = 0 To txtEdit.UBound
        txtEdit(i).Visible = mblnEdit
    Next
    For i = 0 To lbl.UBound
        lbl(i).Visible = mblnEdit
    Next
    For i = 0 To cmdFucn.UBound
        cmdFucn(i).Visible = mblnEdit
    Next
End Sub
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_View_LargeICO   '"大图标(&G)"
            Call SetIcoShow(0)
    Case conMenu_View_MinICO ' "小图标(&M)")
            Call SetIcoShow(1)
    Case conMenu_View_ListICO  '"列表(&L)"
            Call SetIcoShow(2)
    Case conMenu_View_DetailsICO '"详细资料(&D)"
            Call SetIcoShow(3)
    Case conMenu_View_Refresh   '刷新
        Call LoadGroups
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Function GetViewShow(ByVal bytShow As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取列表当前的显示方式
    '入参:bytShow(0-大图标;1-小图标;2-列表;3-详细资料
    '出参:
    '返回:如果显示方式与参数传入的方式一致,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-23 16:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.ActiveControl Is Me.lvwGroups Then
        GetViewShow = (lvwGroups.View = bytShow)
    Else
        GetViewShow = (lvwPerson.View = bytShow)
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = zlIsHaveData
    Case conMenu_View_LargeICO   '"大图标(&G)"
        Control.Checked = GetViewShow(0)
    Case conMenu_View_MinICO ' "小图标(&M)")
        Control.Checked = GetViewShow(1)
    Case conMenu_View_ListICO  '"列表(&L)"
        Control.Checked = GetViewShow(2)
    Case conMenu_View_DetailsICO '"详细资料(&
        Control.Checked = GetViewShow(3)
    Case conMenu_View_Refresh   '刷新
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
           ' Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1502" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1502"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.ID
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Parameter     '参数调用
        Case Else   '其他操作功能调用
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If Me.Visible = False Then Exit Sub

    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub
Private Sub Form_Initialize()
  Call InitCommonControls
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTemp As String
    Err = 0: On Error Resume Next
   SaveWinState Me, App.ProductName
   zlSaveDockPanceToReg Me, dkpMan, "区域"
   SaveListViewState lvwPerson, Me.Name, ""
   SaveListViewState lvwGroups, Me.Name, ""
   SaveListViewState lvwStructure, Me.Name, ""
End Sub

Private Sub lvwGroups_DragDrop(Source As Control, x As Single, y As Single)
    Dim objList As ListItem, strMoveID As String, strMoveName As String, i As Long
    If Source Is lvwPerson And Not lvwGroups.DropHighlight Is Nothing Then
        'Set lvwGroups.SelectedItem = lvwGroups.DropHighlight
        mblnStartDrop = False: mstrSelect = "": mblnReSel = False
        With lvwPerson
            strMoveID = "": strMoveName = "": i = 1
            For Each objList In .ListItems
                If objList.Selected Then
                    strMoveID = strMoveID & "," & Mid(objList.Key, 2)
                    If i > 3 Then
                       If i = 4 Then strMoveName = strMoveName & "..."
                    Else
                        strMoveName = strMoveName & "," & objList.Text
                    End If
                End If
            Next
            If strMoveName <> "" Then strMoveName = Mid(strMoveName, 2)
            If strMoveID <> "" Then strMoveID = Mid(strMoveID, 2)
            If strMoveID = "" Then Exit Sub
            Call PersonFromGroupToOtherGroup(strMoveID, strMoveName, _
                Mid(lvwGroups.SelectedItem.Key, 2), lvwGroups.SelectedItem.Text, _
                Mid(lvwGroups.DropHighlight.Key, 2), lvwGroups.DropHighlight.Text, False)
            Set lvwGroups.DropHighlight = Nothing
            lvwGroups.SelectedItem.EnsureVisible
            Call ClearDropVariable
        End With
    End If
End Sub

Private Sub lvwGroups_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim objOver As ListItem
    If Source Is lvwPerson Then
        Set objOver = lvwGroups.HitTest(x, y)
        If Not objOver Is Nothing Then
            If objOver.Key <> lvwGroups.SelectedItem.Key Then
                Set lvwGroups.DropHighlight = objOver
                lvwGroups.DropHighlight.EnsureVisible
            Else
                Set lvwGroups.DropHighlight = Nothing
            End If
        Else
            Set lvwGroups.DropHighlight = Nothing
        End If
    End If
End Sub

Private Sub lvwGroups_GotFocus()
    '
    Call SetGoupsEnable
End Sub

Private Sub lvwPerson_DragDrop(Source As Control, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False: mblnItemClick = False
End Sub

Private Sub lvwPerson_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim objList As ListItem
    If mblnReSel = True Then Exit Sub
    '多选时,要重新选中
    If Source Is lvwPerson Then
        mblnReSel = True
        If InStr(1, mstrSelect, ",") = 0 Then Exit Sub
        With lvwPerson
            For Each objList In .ListItems
                If InStr("," & mstrSelect & ",", "," & objList.Key & ",") > 0 And objList.Selected = False Then objList.Selected = True
            Next
        End With
    End If
End Sub

Private Sub lvwStructure_DragDrop(Source As Control, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False: mblnItemClick = False
End Sub

Private Sub lvwStructure_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim objList As ListItem
    If mblnReSel = True Then Exit Sub
    '多选时,要重新选中
    If Source Is lvwStructure Then
        mblnReSel = True
        If InStr(1, mstrSelect, ",") = 0 Then Exit Sub
        With lvwStructure
            For Each objList In .ListItems
                If InStr("," & mstrSelect & ",", "," & objList.Key & ",") > 0 And objList.Selected = False Then objList.Selected = True
            Next
        End With
    End If
End Sub

Private Sub lvwPerson_GotFocus()
    Call SetPersonEnable
End Sub

Private Sub lvwStructure_GotFocus()
    Call SetStructureEnable
End Sub

Private Sub lvwPerson_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtEdit(mTxtIdx.idx_成员).Text = Item.Text
    txtEdit(mTxtIdx.idx_成员).Tag = Mid(Item.Key, 2)
    txtEdit(mTxtIdx.idx_原组).Text = Me.lvwGroups.SelectedItem.Text
    txtEdit(mTxtIdx.idx_原组).Tag = Mid(Me.lvwGroups.SelectedItem.Key, 2)
    cmdFucn(mCmdIdx.idx_成员移出).Tag = 1
    Call SetPersonEnable
    mblnItemClick = True
End Sub

Private Sub lvwStructure_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtEdit(mTxtIdx.idx_组长).Text = Item.Text
    txtEdit(mTxtIdx.idx_组长).Tag = Mid(Item.Key, 2)
    cmdFucn(mCmdIdx.idx_组长删除).Tag = 1
    Call SetStructureEnable
    mblnItemClick = True
End Sub

Private Sub SetPersonEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置组成员的编辑属性
    '编制:刘兴洪
    '日期:2010-11-24 17:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDelete As Boolean
    blnDelete = Val(cmdFucn(mCmdIdx.idx_成员移出).Tag) > 0
    cmdFucn(mCmdIdx.idx_成员移出).Enabled = blnDelete
    cmdFucn(mCmdIdx.idx_成员增加).Enabled = Not blnDelete
    txtEdit(mTxtIdx.idx_原组).Enabled = False
    txtEdit(mTxtIdx.idx_原组).BackColor = Me.BackColor
End Sub
Private Sub SetGoupsEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置组的编辑属性
    '编制:刘兴洪
    '日期:2010-11-24 17:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnModify As Boolean
    blnModify = Not Me.lvwGroups.SelectedItem Is Nothing
    cmdFucn(mCmdIdx.idx_组删除).Enabled = blnModify
    cmdFucn(mCmdIdx.idx_组修改).Enabled = blnModify
    If Not blnModify Then
        cmdFucn(mCmdIdx.idx_组增加).Enabled = Trim(txtEdit(mTxtIdx.idx_组名称).Text) <> ""
    Else
        cmdFucn(mCmdIdx.idx_组增加).Enabled = Trim(txtEdit(mTxtIdx.idx_组名称).Text) <> "" And lvwGroups.SelectedItem.Text <> Trim(txtEdit(mTxtIdx.idx_组名称))
    End If
End Sub

Private Sub SetStructureEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置组的编辑属性
    '编制:刘兴洪
    '日期:2010-11-24 17:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnModify As Boolean
    blnModify = Not Me.lvwStructure.SelectedItem Is Nothing
    cmdFucn(mCmdIdx.idx_组长删除).Enabled = blnModify
    If Not blnModify Then
        cmdFucn(mCmdIdx.idx_组长增加).Enabled = Trim(txtEdit(mTxtIdx.idx_组长).Text) <> ""
    Else
        cmdFucn(mCmdIdx.idx_组长增加).Enabled = Trim(txtEdit(mTxtIdx.idx_组长).Text) <> "" And lvwStructure.SelectedItem.Text <> Trim(txtEdit(mTxtIdx.idx_组长))
    End If
End Sub

Private Sub lvwPerson_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False
End Sub

Private Sub lvwStructure_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False
End Sub

Private Sub lvwPerson_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objList As ListItem, i As Long
    If Button <> 1 Then Exit Sub
    If mblnEdit = False Then Exit Sub '不能编辑的,不能拖动
    If mblnStartDrop Then Exit Sub
    If mblnItemClick = False Then Exit Sub
    
    '拖动开始
    With lvwPerson
        If .ListItems.Count = 0 Then Exit Sub
        For Each objList In .ListItems
            If objList.Selected Then mstrSelect = mstrSelect & "," & objList.Key
        Next
    End With
    If mstrSelect <> "" Then mstrSelect = Mid(mstrSelect, 2)
    If InStr(1, mstrSelect, ",") > 0 Then
        Set lvwPerson.DragIcon = igl32.ListImages("Group").Picture
    Else
        Set lvwPerson.DragIcon = lvwPerson.SelectedItem.CreateDragImage
    End If
    lvwPerson.Drag 1
    mblnStartDrop = True: mblnReSel = False
End Sub

Private Sub lvwStructure_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objList As ListItem, i As Long
    If Button <> 1 Then Exit Sub
    If mblnEdit = False Then Exit Sub '不能编辑的,不能拖动
    If mblnStartDrop Then Exit Sub
    If mblnItemClick = False Then Exit Sub
    
    '拖动开始
    With lvwStructure
        If .ListItems.Count = 0 Then Exit Sub
        For Each objList In .ListItems
            If objList.Selected Then mstrSelect = mstrSelect & "," & objList.Key
        Next
    End With
    If mstrSelect <> "" Then mstrSelect = Mid(mstrSelect, 2)
    If InStr(1, mstrSelect, ",") > 0 Then
        Set lvwStructure.DragIcon = igl32.ListImages("Group").Picture
    Else
        Set lvwStructure.DragIcon = lvwStructure.SelectedItem.CreateDragImage
    End If
    lvwStructure.Drag 1
    mblnStartDrop = True: mblnReSel = False
End Sub

Private Sub picGroup_Resize()
    Dim sngLeft As Single
    Dim sngTop As Single
    Err = 0: On Error Resume Next
    With picGroup
        
        sngLeft = .ScaleWidth - (txtEdit(mTxtIdx.idx_负责人).Width + txtEdit(mTxtIdx.idx_负责人).Left)
        sngLeft = sngLeft - (cmdFucn(mCmdIdx.idx_组增加).Width + 10) * 3
        sngTop = txtEdit(mTxtIdx.idx_负责人).Top
        If sngLeft < 0 Then
            sngLeft = txtEdit(mTxtIdx.idx_负责人).Left
            sngTop = txtEdit(mTxtIdx.idx_负责人).Top + txtEdit(mTxtIdx.idx_负责人).Height + 100
        ElseIf sngLeft < (txtEdit(mTxtIdx.idx_负责人).Width + txtEdit(mTxtIdx.idx_负责人).Left) Or sngLeft > (txtEdit(mTxtIdx.idx_负责人).Width + txtEdit(mTxtIdx.idx_负责人).Left) Then
                sngLeft = (txtEdit(mTxtIdx.idx_负责人).Width + txtEdit(mTxtIdx.idx_负责人).Left) + 100
        End If
        
        cmdFucn(mCmdIdx.idx_组增加).Left = sngLeft
        cmdFucn(mCmdIdx.idx_组修改).Left = cmdFucn(mCmdIdx.idx_组增加).Left + cmdFucn(mCmdIdx.idx_组增加).Width + 10
        cmdFucn(mCmdIdx.idx_组删除).Left = cmdFucn(mCmdIdx.idx_组修改).Left + cmdFucn(mCmdIdx.idx_组修改).Width + 10
        cmdFucn(mCmdIdx.idx_组修改).Top = sngTop
        cmdFucn(mCmdIdx.idx_组删除).Top = sngTop
        cmdFucn(mCmdIdx.idx_组增加).Top = sngTop
        txtEdit(mTxtIdx.idx_组名称).Width = .ScaleWidth - txtEdit(mTxtIdx.idx_组名称).Left - 100
        txtEdit(mTxtIdx.idx_组说明).Width = .ScaleWidth - txtEdit(mTxtIdx.idx_组说明).Left - 100
        If mblnEdit = False Then
             lvwGroups.Top = .ScaleTop
        Else
            lvwGroups.Top = sngTop + cmdFucn(mCmdIdx.idx_组增加).Height + 50
        End If
        lvwGroups.Left = .ScaleLeft
        lvwGroups.Width = .ScaleWidth
        lvwGroups.Height = .ScaleHeight - lvwGroups.Top - 50
    End With
End Sub

Private Sub picPersons_Resize()
    Dim sngLeft As Single, sngTop As Single
    Err = 0: On Error Resume Next
    With picPersons
        If .ScaleWidth - txtEdit(mTxtIdx.idx_成员).Left > 3900 Then
            txtEdit(mTxtIdx.idx_成员).Width = 3900
        Else
            txtEdit(mTxtIdx.idx_成员).Width = .ScaleWidth - txtEdit(mTxtIdx.idx_成员).Left - 50
        End If
        txtEdit(mTxtIdx.idx_原组).Width = txtEdit(mTxtIdx.idx_成员).Width
        
        lvwPerson.Left = .ScaleLeft
        sngLeft = txtEdit(mTxtIdx.idx_原组).Left + txtEdit(mTxtIdx.idx_原组).Width + (cmdFucn(mCmdIdx.idx_成员移出).Width + 10) * 2
        sngLeft = .ScaleWidth - sngLeft
        If sngLeft < 0 Then
            sngTop = txtEdit(mTxtIdx.idx_原组).Top + txtEdit(mTxtIdx.idx_原组).Height + 50
            sngLeft = txtEdit(mTxtIdx.idx_原组).Left
        Else
            sngLeft = txtEdit(mTxtIdx.idx_原组).Left + txtEdit(mTxtIdx.idx_原组).Width + 50
            sngTop = txtEdit(mTxtIdx.idx_原组).Top
        End If
        cmdFucn(mCmdIdx.idx_成员增加).Left = sngLeft
        cmdFucn(mCmdIdx.idx_成员移出).Left = cmdFucn(mCmdIdx.idx_成员增加).Left + cmdFucn(mCmdIdx.idx_成员增加).Width + 10
        cmdFucn(mCmdIdx.idx_成员增加).Top = sngTop
        cmdFucn(mCmdIdx.idx_成员移出).Top = sngTop
        sngTop = sngTop + cmdFucn(mCmdIdx.idx_成员增加).Height + 50
        If mblnEdit Then
            lvwPerson.Top = sngTop
        Else
            lvwPerson.Top = .ScaleTop
        End If

        
        lvwPerson.Width = .ScaleWidth
        lvwPerson.Height = .ScaleHeight - lvwPerson.Top
    End With
End Sub
Private Sub SetIcoShow(ByVal bytShow As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置图标显示方式
    '入参:bytShow-(0-大图标;1- 小图标;2-列表;3-详细资料
    '编制:刘兴洪
    '日期:2010-11-23 16:54:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLvw As Object
    If Me.ActiveControl Is Me.lvwGroups Then
        Set objLvw = lvwGroups
    Else
        Set objLvw = lvwPerson
    End If
    With objLvw
        .View = bytShow
    End With
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '入参:
    '出参:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-15 11:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
        
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标(&G)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细资料(&D)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细资料")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2010-11-15 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, lngWidth As Long
    Dim lngHeight As Long
    With dkpMan
        Set objPane = .CreatePane(mPaneID.Pane_Group, 400, 400, DockLeftOf, Nothing)
        objPane.Title = "缴款人员分组信息": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picGroup.hWnd
        objPane.Tag = mPaneID.Pane_Group
        
        Set objPane = .CreatePane(mPaneID.Pane_Structure, 400, 200, DockRightOf)
        objPane.Title = "组长构成信息"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picStructure.hWnd
        objPane.Tag = mPaneID.Pane_Structure
        
        Set objPane = .CreatePane(mPaneID.Pane_Persons, 400, 400, DockBottomOf, objPane)
        objPane.Title = "组成员信息"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picPersons.hWnd
        objPane.Tag = mPaneID.Pane_Persons
        
        
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
'    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Function
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPaneID.Pane_Persons
        Item.Handle = picPersons.hWnd
    Case mPaneID.Pane_Group
        Item.Handle = picGroup.hWnd
    End Select
End Sub

Private Function LoadGroups() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载入库数据给网格
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-15 14:54:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, lngPreID As Long, objItem As ListItem
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    
    Err = 0: On Error GoTo errHandle:
    gstrSQL = "" & _
    "   Select A.Id, A.组名称,A.简码, A.说明, A.负责人id, A.删除日期,B.姓名 as 负责人  " & _
    "   From 财务缴款分组 A,人员表 B " & _
    "   Where A.负责人ID=B.Id(+) And (A.删除日期>Sysdate Or A.删除日期 Is Null)"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With lvwGroups
        .ListItems.Clear
        Do While Not rsTemp.EOF
            Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!ID), NVL(rsTemp!组名称), "Group", "Group")
            objItem.SubItems(1) = NVL(rsTemp!负责人)
            objItem.SubItems(2) = NVL(rsTemp!说明)
            objItem.Tag = NVL(rsTemp!负责人id)
            If mstrPreGroupKey = objItem.Key Then objItem.Selected = True: objItem.EnsureVisible
            rsTemp.MoveNext
        Loop
        If Not .SelectedItem Is Nothing Then
            Call lvwGroups_ItemClick(.SelectedItem)
        End If
    End With
    LoadGroups = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub LoadGroupStructure(ByVal lng组ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载组成员数据
    '入参:lng组ID-组ID
    '编制:刘兴洪
    '日期:2010-11-23 17:36:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey  As String, objItem As ListItem, i As Integer
    Dim strIco As String
    
    On Error GoTo ErrHand
    
    gstrSQL = " " & _
    "   Select A.组Id ,A.组长ID, B.编号,B.姓名,B.简码,b.出生日期,B.身份证号,B.性别,B.民族,B.办公室电话,B.电子邮件,B.管理职务 " & _
    "   From 财务组组长构成 A,人员表 B " & _
    "   Where A.组长ID=B.Id And A.组ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng组ID)
    
   With lvwStructure
        If Not .SelectedItem Is Nothing Then strKey = .SelectedItem.Key
        .ListItems.Clear
        Do While Not rsTemp.EOF
            If InStr(1, NVL(rsTemp!性别), "男") > 0 Then
                strIco = "Man"
            ElseIf InStr(1, NVL(rsTemp!性别), "女") > 0 Then
                strIco = "Woman"
            Else
                strIco = "Man" ' "Other"
            End If
            Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!组长ID), NVL(rsTemp!姓名), strIco, strIco)
            i = 1
            objItem.SubItems(i) = NVL(rsTemp!编号): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!简码): i = i + 1
            objItem.SubItems(i) = Format(rsTemp!出生日期, "yyyy-mm-dd"): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!性别): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!民族): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!办公室电话): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!电子邮件): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!管理职务): i = i + 1
            If strKey = objItem.Key Then
                objItem.Selected = True: objItem.EnsureVisible
            End If
            rsTemp.MoveNext
        Loop
        If Not .SelectedItem Is Nothing Then
            Call lvwStructure_ItemClick(.SelectedItem)
        End If
    End With
    mstrSelect = "": mblnItemClick = False: mblnStartDrop = False: mblnReSel = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub LoadGroupPersons(ByVal lng组ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载组成员数据
    '入参:lng组ID-组ID
    '编制:刘兴洪
    '日期:2010-11-23 17:36:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey  As String, objItem As ListItem, i As Integer
    Dim strIco As String
    
    On Error GoTo ErrHand
    
    gstrSQL = " " & _
    "   Select A.组Id ,A.成员ID, B.编号,B.姓名,B.简码,b.出生日期,B.身份证号,B.性别,B.民族,B.办公室电话,B.电子邮件,B.管理职务 " & _
    "   From 缴款成员组成 A,人员表 B " & _
    "   Where A.成员ID=B.Id And A.组ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng组ID)
    
   With lvwPerson
        If Not .SelectedItem Is Nothing Then strKey = .SelectedItem.Key
        .ListItems.Clear
        Do While Not rsTemp.EOF
            If InStr(1, NVL(rsTemp!性别), "男") > 0 Then
                strIco = "Man"
            ElseIf InStr(1, NVL(rsTemp!性别), "女") > 0 Then
                strIco = "Woman"
            Else
                strIco = "Man" ' "Other"
            End If
            Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!成员ID), NVL(rsTemp!姓名), strIco, strIco)
            i = 1
            objItem.SubItems(i) = NVL(rsTemp!编号): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!简码): i = i + 1
            objItem.SubItems(i) = Format(rsTemp!出生日期, "yyyy-mm-dd"): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!性别): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!民族): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!办公室电话): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!电子邮件): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!管理职务): i = i + 1
            If strKey = objItem.Key Then
                objItem.Selected = True: objItem.EnsureVisible
            End If
            rsTemp.MoveNext
        Loop
        If Not .SelectedItem Is Nothing Then
            Call lvwPerson_ItemClick(.SelectedItem)
        End If
    End With
    mstrSelect = "": mblnItemClick = False: mblnStartDrop = False: mblnReSel = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
 Private Sub zlRptPrint(ByVal bytFunc As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2010-11-23 17:55:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrint As Object, objLvw As Object
    Dim objRow As New zlTabAppRow
    Dim str单位 As String
        
    Set objPrint = New zlPrintLvw
    If Me.ActiveControl Is lvwGroups Then
        objPrint.Title.Text = GetUnitName & "分组清单"
        Set objLvw = lvwGroups
    Else
        If lvwGroups Is Nothing Then Exit Sub
        objPrint.Title.Text = GetUnitName & lvwGroups.SelectedItem.Text & "成员组成"
        Set objLvw = lvwPerson
    End If
    Set objPrint.Body.objData = objLvw
    objPrint.BelowAppItems.Add "打印人：" & UserInfo.姓名
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytFunc
    End If
End Sub
Private Sub DeleteGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除组
    '编制:刘兴洪
    '日期:2010-11-23 17:59:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTittle As String, intIndex As Integer
    Dim rsTemp As ADODB.Recordset
    With lvwGroups
        If .SelectedItem Is Nothing Then Exit Sub
        lngID = Val(Mid(.SelectedItem.Key, 2))
        strTittle = .SelectedItem.Text
    End With
    If lngID = 0 Then Exit Sub
    If MsgBox("你确认要删除组名称为『 " & strTittle & "』的分组吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    
    gstrSQL = "" & _
    "   Select Count(distinct A.成员ID) as 成员数,Sum(nvl(C.余额,0)) as 余额 " & _
    "   From  缴款成员组成 A,人员表 B,人员缴款余额 C " & _
    "   where A.成员ID=B.id and B.姓名=C.收款员(+) and C.性质(+)=1 and  A.组ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If Val(NVL(rsTemp!成员数)) <> 0 Then
        If Val(NVL(rsTemp!余额)) = 0 Then
            If MsgBox("组名称为『 " & strTittle & "』下还有" & Val(NVL(rsTemp!成员数)) & "个成员,你是否要解散该组？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Me.MousePointer = 0
                Exit Sub
            End If
        Else
            If MsgBox("组名称为『 " & strTittle & "』下还有" & Val(NVL(rsTemp!成员数)) & "个" & vbCrLf & "成员,并且还存在暂存金,你是否还要解散该组？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    'Zl_财务缴款分组_Delete(Id_In In 财务缴款分组.ID%Type) Is
    gstrSQL = "Zl_财务缴款分组_Delete(" & lngID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = 0
    With lvwGroups
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            lvwGroups_ItemClick .SelectedItem
        Else
            Call lvwGroups_GotFocus
        End If
    End With
    Call SetStructureEnable
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub



Private Sub PersonFromGroupToOtherGroup(ByVal str成员ID As String, ByVal str成员姓名 As String, _
    lng原组ID As Long, str原组名称 As String, _
    Optional lng新组ID As Long = -1, Optional str新组名称 As String = "", _
    Optional blnFromOtherGroupMoveCur As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:人员从一组移动到另一个组,或则移除某一组
    '入参:str成员ID-指定的成员(多个时,用逗号分离)
    '       lng原组ID-原组的ID
    '       lng新组ID-新组的ID(为-1表示移除)
    '       blnFromOtherGroupMoveCur-从其他组移动到当前组;否则从当前组移动到其他组
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-24 10:51:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTemp As String, intIndex As Integer, strID As String
    Dim rsTemp As ADODB.Recordset, cllPro As Collection
    Dim varData As Variant, i As Long
    
    If str成员ID = "" Then Exit Sub
    
    If lng新组ID < 0 Then
        If MsgBox("你确认要将人员『 " & str成员姓名 & "』从『 " & str原组名称 & "』中移出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("你确认要将人员『 " & str成员姓名 & "』从 『 " & str原组名称 & "』中" & vbCrLf & "移到 『 " & str新组名称 & "』吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    Set cllPro = New Collection
    
    ' Zl_缴款成员组成_Move
    '  成员id_In Varchar2,
    '  原组id_In In 缴款成员组成.组id%Type,
    '  新组id_In In 缴款成员组成.组id%Type := -1
     If Len(str成员ID) < 2000 Then
        gstrSQL = "Zl_缴款成员组成_Move('" & str成员ID & "'," & lng原组ID & "," & lng新组ID & ")"
        AddArray cllPro, gstrSQL
     Else
        varData = Split(str成员ID, ",")
        strTemp = ""
        For i = 0 To UBound(varData)
            If varData(i) <> "" Then
                If Len(strTemp) >= 1980 Then
                    strTemp = Mid(strTemp, 2)
                    gstrSQL = "Zl_缴款成员组成_Move('" & strTemp & "'," & lng原组ID & "," & lng新组ID & ")"
                    AddArray cllPro, gstrSQL
                    strTemp = ""
                End If
                strTemp = strTemp & "," & varData(i)
            End If
        Next
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            gstrSQL = "Zl_缴款成员组成_Move('" & strTemp & "'," & lng原组ID & "," & lng新组ID & ")"
            AddArray cllPro, gstrSQL
        End If
    End If
    Err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    
    Err = 0: On Error GoTo errHandle:
    Dim objItem As ListItem
    Me.MousePointer = 0
    
    With lvwPerson
        intIndex = .SelectedItem.Index
        If blnFromOtherGroupMoveCur = False Or lng新组ID <= 0 Then   '从当前组移到其他组时,需要移出数据
            varData = Split(str成员ID, ",")
            For i = 0 To UBound(varData)
                If varData(i) <> "" Then .ListItems.Remove "K" & varData(i)
            Next
        Else     '从其他组移动到当前组时,需要增加数据
            varData = Split(str成员ID, ",")
            For i = 0 To UBound(varData)
                LoadLocalPerson Val(varData(i))
            Next
        End If
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            lvwPerson_ItemClick .SelectedItem
        Else
            Call lvwGroups_GotFocus
        End If
    End With
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrHand:
    Me.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DeleteStructure(ByVal str成员ID As String, ByVal str成员姓名 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除组长
    '入参:str成员ID-指定的成员(多个时,用逗号分离)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-24 10:51:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTemp As String, intIndex As Integer, strID As String
    Dim rsTemp As ADODB.Recordset, cllPro As Collection
    Dim varData As Variant, i As Long
    
    If str成员ID = "" Then Exit Sub
    
    If MsgBox("你确认要将组长『 " & str成员姓名 & "』移出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    Set cllPro = New Collection
    
    ' Zl_缴款成员组成_Move
    '  成员id_In Varchar2,
    '  原组id_In In 缴款成员组成.组id%Type,
    '  新组id_In In 缴款成员组成.组id%Type := -1
     If Len(str成员ID) < 2000 Then
        gstrSQL = "Zl_缴款成员组成_Move('" & str成员ID & "'," & Mid(lvwGroups.SelectedItem.Key, 2) & ",Null" & ",1)"
        AddArray cllPro, gstrSQL
     Else
        varData = Split(str成员ID, ",")
        strTemp = ""
        For i = 0 To UBound(varData)
            If varData(i) <> "" Then
                If Len(strTemp) >= 1980 Then
                    strTemp = Mid(strTemp, 2)
                    gstrSQL = "Zl_缴款成员组成_Move('" & strTemp & "'," & Mid(lvwGroups.SelectedItem.Key, 2) & ",Null" & ",1)"
                    AddArray cllPro, gstrSQL
                    strTemp = ""
                End If
                strTemp = strTemp & "," & varData(i)
            End If
        Next
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            gstrSQL = "Zl_缴款成员组成_Move('" & strTemp & "'," & Mid(lvwGroups.SelectedItem.Key, 2) & ",Null" & ",1)"
            AddArray cllPro, gstrSQL
        End If
    End If
    Err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    
    Err = 0: On Error GoTo errHandle:
    Dim objItem As ListItem
    Me.MousePointer = 0
    
    With lvwStructure
        intIndex = .SelectedItem.Index

        varData = Split(str成员ID, ",")
        For i = 0 To UBound(varData)
            If varData(i) <> "" Then .ListItems.Remove "K" & varData(i)
        Next

        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            lvwStructure_ItemClick .SelectedItem
        Else
            Call lvwGroups_GotFocus
        End If
    End With
    Call SetStructureEnable
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrHand:
    Me.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ClearDropVariable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除拖动变量值
    '编制:刘兴洪
    '日期:2010-11-26 16:38:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnItemClick = False: mstrSelect = "": mblnStartDrop = False: mblnReSel = False
End Sub

Private Function CheckGroupInput(ByVal lngID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查组输入是否合法
    '入参:
    '出参:
    '返回:合法,返回true,否则返加False
    '编制:刘兴洪
    '日期:2010-11-24 16:15:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If Trim(txtEdit(mTxtIdx.idx_组名称)) = "" Then
        ShowMsgbox "组名称必须输入,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组名称): Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mTxtIdx.idx_组名称)), 50, 0, "组名称") = False Then
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组名称): Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mTxtIdx.idx_组说明)), 50, 0, "组说明") = False Then
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组说明): Exit Function
    End If
    If Val(txtEdit(mTxtIdx.idx_负责人).Tag) = 0 Then
        ShowMsgbox "负责人必须输入或输入不合法,请选择!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_负责人): Exit Function
    End If
    gstrSQL = "Select 1 From 财务缴款分组 where 组名称=[1] and 删除日期>=sysdate and ID+0<>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtEdit(mTxtIdx.idx_组名称)), lngID)
    If Not rsTemp.EOF Then
        ShowMsgbox "组名称已经存在,不能再增加!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组名称): Exit Function
    End If
    CheckGroupInput = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddGroups(bytType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加组信息
    '入参:bytType:0-增加;1-修改
    '编制:刘兴洪
    '日期:2010-11-24 16:04:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    Dim objList As ListItem
    
    If bytType <> 0 Then
        With lvwGroups
            If .SelectedItem Is Nothing Then Exit Sub
            lngID = Val(Mid(.SelectedItem.Key, 2))
        End With
    End If
    
    If CheckGroupInput(lngID) = False Then Exit Sub
    
    On Error GoTo errHandle
    If bytType = 0 Then
        lngID = zlDatabase.GetNextId("财务缴款分组")
    End If
   ' Zl_财务缴款分组_Update
   gstrSQL = "Zl_财务缴款分组_Update("
    '  Id_In       In 财务缴款分组.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  组名称_In   In 财务缴款分组.组名称%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_组名称).Text) & "',"
    '  简码_In     In 财务缴款分组.简码%Type,
    gstrSQL = gstrSQL & "'" & Left(Trim(zlCommFun.SpellCode(Trim(mTxtIdx.idx_组名称))), 20) & "',"
    '  说明_In     In 财务缴款分组.说明%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_组说明).Text) & "',"
    '  负责人id_In In 财务缴款分组.负责人id%Type,
    gstrSQL = gstrSQL & "" & Val(txtEdit(mTxtIdx.idx_负责人).Tag) & ","
    '  修改标志_In Integer:=0
    gstrSQL = gstrSQL & "" & bytType & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    If bytType = 0 Then
        Set objList = lvwGroups.ListItems.Add(, "K" & lngID, Trim(txtEdit(mTxtIdx.idx_组名称).Text), "Group", "Group")
    Else
        Set objList = lvwGroups.SelectedItem
    End If
    objList.Text = Trim(txtEdit(mTxtIdx.idx_组名称).Text)
    objList.Tag = Trim(txtEdit(mTxtIdx.idx_负责人).Tag)
    objList.SubItems(1) = Trim(txtEdit(mTxtIdx.idx_负责人).Text)
    objList.SubItems(2) = Trim(txtEdit(mTxtIdx.idx_组说明).Text)
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CheckPersonInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查组成员的输入是否合法
    '入参:
    '出参:
    '返回:合法,返回true,否则返加False
    '编制:刘兴洪
    '日期:2010-11-24 16:15:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem
    
    On Error GoTo errHandle
    If lvwGroups.SelectedItem Is Nothing Then Exit Function
    
    If Val(txtEdit(mTxtIdx.idx_成员).Tag) = 0 Then
        ShowMsgbox "组成员必须选择,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_成员): Exit Function
    End If
    If Val(Mid(lvwGroups.SelectedItem.Key, 2)) <> Val(txtEdit(mTxtIdx.idx_原组).Tag) Then
        If Val(txtEdit(mTxtIdx.idx_原组).Tag) <> 0 Then
            gstrSQL = "Select sum(nvl(余额,0) ) as 余额 From 人员缴款余额 Where 性质=1 and 收款员=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtEdit(mTxtIdx.idx_成员)))
            If Val(NVL(rsTemp!余额)) > 0 Then
                ShowMsgbox "人员在组名为" & txtEdit(mTxtIdx.idx_组名称) & "中还存在暂存金,不能移动到该组!"
                zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_成员): Exit Function
            End If
        End If
    Else
        For Each objItem In lvwPerson.ListItems
            If Val(txtEdit(mTxtIdx.idx_成员).Tag) = Val(Mid(objItem.Key, 2)) Then
                ShowMsgbox "人员" & txtEdit(mTxtIdx.idx_成员) & "已经在该组中存在,没必要再增加,请检查!"
                zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_成员): Exit Function
                Exit Function
            End If
        Next
    End If
    CheckPersonInput = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckStructureInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查组成员的输入是否合法
    '入参:
    '出参:
    '返回:合法,返回true,否则返加False
    '编制:刘兴洪
    '日期:2010-11-24 16:15:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem
    Dim strSQL As String
    
    On Error GoTo errHandle
    If lvwGroups.SelectedItem Is Nothing Then Exit Function
    
    If Val(txtEdit(mTxtIdx.idx_组长).Tag) = 0 Then
        ShowMsgbox "组长必须选择,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组长): Exit Function
    End If

    For Each objItem In lvwStructure.ListItems
        If Val(txtEdit(mTxtIdx.idx_组长).Tag) = Val(Mid(objItem.Key, 2)) Then
            ShowMsgbox "组长" & txtEdit(mTxtIdx.idx_组长) & "已经在该组中存在,没必要再增加,请检查!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组长): Exit Function
            Exit Function
        End If
    Next
    
    CheckStructureInput = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddPerson()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加组人员信息
    '编制:刘兴洪
    '日期:2010-11-24 16:04:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    If CheckPersonInput = False Then Exit Sub
    On Error GoTo errHandle
    If Val(txtEdit(mTxtIdx.idx_原组).Tag) <> 0 Then
        Call PersonFromGroupToOtherGroup(txtEdit(mTxtIdx.idx_成员).Tag, txtEdit(mTxtIdx.idx_成员), txtEdit(mTxtIdx.idx_原组).Tag, txtEdit(mTxtIdx.idx_原组), Mid(lvwGroups.SelectedItem.Key, 2), lvwGroups.SelectedItem.Text)
        'Call LoadLocalPerson(Val(txtEdit(mTxtIdx.idx_成员).Tag))
        Exit Sub
    End If
    'Zl_缴款成员组成_Insert
    gstrSQL = "Zl_缴款成员组成_Insert("
    '  组id_In   In 缴款成员组成.组id%Type,
    gstrSQL = gstrSQL & "" & Mid(lvwGroups.SelectedItem.Key, 2) & ","
    '  成员id_In In 缴款成员组成.成员id%Type
    gstrSQL = gstrSQL & "" & Val(txtEdit(mTxtIdx.idx_成员).Tag) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Call LoadLocalPerson(Val(txtEdit(mTxtIdx.idx_成员).Tag))
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_成员)
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AddStructure()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加组人员信息
    '编制:刘兴洪
    '日期:2010-11-24 16:04:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    If CheckStructureInput = False Then Exit Sub
    On Error GoTo errHandle
    'Zl_缴款成员组成_Insert
    gstrSQL = "Zl_缴款成员组成_Insert("
    '  组id_In   In 缴款成员组成.组id%Type,
    gstrSQL = gstrSQL & "" & Mid(lvwGroups.SelectedItem.Key, 2) & ","
    '  成员id_In In 缴款成员组成.成员id%Type
    gstrSQL = gstrSQL & "" & Val(txtEdit(mTxtIdx.idx_组长).Tag) & ",1)"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Call LoadLocalStructure(Val(txtEdit(mTxtIdx.idx_组长).Tag))
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组长)
    Call SetGoupsEnable
    Call SetPersonEnable
    Call SetStructureEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadLocalPerson(ByVal lng成员ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载指定人员的信息给ListView
    '编制:刘兴洪
    '日期:2010-11-24 16:52:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem, i As Long
    Dim strIco As String
    On Error GoTo errHandle
    
    gstrSQL = " " & _
    "   Select  B.ID  ,B.编号,B.姓名,B.简码,b.出生日期,B.身份证号,B.性别,B.民族,B.办公室电话,B.电子邮件,B.管理职务 " & _
    "   From  人员表 B " & _
    "   Where   B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng成员ID)
    If rsTemp.EOF Then Exit Sub
    
   With lvwPerson
        If InStr(1, NVL(rsTemp!性别), "男") > 0 Then
            strIco = "Man"
        ElseIf InStr(1, NVL(rsTemp!性别), "女") > 0 Then
            strIco = "Woman"
        Else
            strIco = "Man" '"Other"
        End If
        Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!ID), NVL(rsTemp!姓名), strIco, strIco)
        i = 1
        objItem.SubItems(i) = NVL(rsTemp!编号): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!简码): i = i + 1
        objItem.SubItems(i) = Format(rsTemp!出生日期, "yyyy-mm-dd"): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!性别): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!民族): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!办公室电话): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!电子邮件): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!管理职务): i = i + 1
        objItem.Selected = True: objItem.EnsureVisible
        Call lvwPerson_ItemClick(objItem)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadLocalStructure(ByVal lng成员ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载指定人员的信息给ListView
    '编制:刘兴洪
    '日期:2010-11-24 16:52:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem, i As Long
    Dim strIco As String
    On Error GoTo errHandle
    
    gstrSQL = " " & _
    "   Select  B.ID  ,B.编号,B.姓名,B.简码,b.出生日期,B.身份证号,B.性别,B.民族,B.办公室电话,B.电子邮件,B.管理职务 " & _
    "   From  人员表 B " & _
    "   Where   B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng成员ID)
    If rsTemp.EOF Then Exit Sub
    
   With lvwStructure
        If InStr(1, NVL(rsTemp!性别), "男") > 0 Then
            strIco = "Man"
        ElseIf InStr(1, NVL(rsTemp!性别), "女") > 0 Then
            strIco = "Woman"
        Else
            strIco = "Man" '"Other"
        End If
        Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!ID), NVL(rsTemp!姓名), strIco, strIco)
        i = 1
        objItem.SubItems(i) = NVL(rsTemp!编号): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!简码): i = i + 1
        objItem.SubItems(i) = Format(rsTemp!出生日期, "yyyy-mm-dd"): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!性别): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!民族): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!办公室电话): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!电子邮件): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!管理职务): i = i + 1
        objItem.Selected = True: objItem.EnsureVisible
        Call lvwStructure_ItemClick(objItem)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否存在数据
    '编制:刘兴洪
    '日期:2010-11-15 17:54:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.ActiveControl Is lvwGroups Then
        With lvwGroups
            zlIsHaveData = .ListItems.Count > 0: Exit Function
        End With
    Else
        With lvwPerson
            zlIsHaveData = .ListItems.Count > 0: Exit Function
        End With
    End If
End Function
 
Private Sub lvwGroups_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintGroupColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwGroups.SortOrder = IIf(lvwGroups.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintGroupColumn = ColumnHeader.Index - 1
        lvwGroups.SortKey = mintGroupColumn
        lvwGroups.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwGroups_DblClick()
    If mblnEdit Then
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_组名称)
    End If
End Sub
 
Private Sub lvwGroups_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub lvwGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
       Set objPopup = cbsThis.FindControl(, conMenu_ViewPopup, , True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
Private Sub lvwGroups_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrPreGroupKey = Item.Key Then Exit Sub
    mstrPreGroupKey = Item.Key
    txtEdit(mTxtIdx.idx_组名称).Text = Item.Text
    txtEdit(mTxtIdx.idx_负责人).Text = Item.SubItems(1)
    txtEdit(mTxtIdx.idx_负责人).Tag = Item.Tag
    txtEdit(mTxtIdx.idx_组说明).Text = Item.SubItems(2)
    Call LoadGroupPersons(Val(Mid(Item.Key, 2)))  '加载成员信息
    Call LoadGroupStructure(Val(Mid(Item.Key, 2)))
    Call SetGoupsEnable
End Sub


Private Sub lvwPerson_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintPersonColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwPerson.SortOrder = IIf(lvwPerson.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintPersonColumn = ColumnHeader.Index - 1
        lvwPerson.SortKey = mintPersonColumn
        lvwPerson.SortOrder = lvwAscending
    End If
End Sub

  
Private Sub lvwPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lvwStructure_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lvwPerson_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim objPopup As CommandBarPopup
    mblnStartDrop = False
    If Button = 2 Then
       Set objPopup = cbsThis.FindControl(, conMenu_ViewPopup, , True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
    mblnItemClick = False
End Sub

Private Sub lvwStructure_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintStructureColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwStructure.SortOrder = IIf(lvwStructure.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintStructureColumn = ColumnHeader.Index - 1
        lvwStructure.SortKey = mintStructureColumn
        lvwStructure.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwStructure_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim objPopup As CommandBarPopup
    mblnStartDrop = False
    If Button = 2 Then
       Set objPopup = cbsThis.FindControl(, conMenu_ViewPopup, , True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
    mblnItemClick = False
End Sub


Private Sub picStructure_Resize()
    With lvwStructure
        .Height = picStructure.ScaleHeight - .Top - 30
        .Width = picStructure.ScaleWidth
    End With
End Sub

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = "":
    If mTxtIdx.idx_成员 = Index Then
        cmdFucn(mCmdIdx.idx_成员移出).Tag = ""
        Call SetPersonEnable
    Else
        If Index = mTxtIdx.idx_组长 Then
            Call SetStructureEnable
        Else
            Call SetGoupsEnable
            Call SetPersonEnable
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = mTxtIdx.idx_成员 Or idx_负责人 = Index Then
        zlCommFun.OpenIme False
    Else
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lng人员ID As Long, rsTemp As ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errHandle
    
    Select Case Index
    Case mTxtIdx.idx_负责人
       If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index)), , lng人员ID) = False Then Exit Sub
       txtEdit(Index).Tag = lng人员ID
    Case mTxtIdx.idx_成员
       If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index)), , lng人员ID, , "", "门诊挂号员,门诊收费员,预交收款员,住院结帐员,入院登记员,发卡登记人") = False Then Exit Sub
       txtEdit(Index).Tag = lng人员ID
       '获取相关的组信息
        gstrSQL = "" & _
        "   Select A.Id, A.组名称 " & _
        "   From 财务缴款分组 A,缴款成员组成 B " & _
        "   Where A.ID=B.组Id and B.成员ID=[1] And (A.删除日期>Sysdate Or A.删除日期 Is Null)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng人员ID)
        If Not rsTemp.EOF Then
            txtEdit(mTxtIdx.idx_原组).Text = NVL(rsTemp!组名称)
            txtEdit(mTxtIdx.idx_原组).Tag = NVL(rsTemp!ID)
        Else
            txtEdit(mTxtIdx.idx_原组).Text = ""
        End If
    Case mTxtIdx.idx_组长
        If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index)), , lng人员ID) = False Then Exit Sub
        txtEdit(Index).Tag = lng人员ID
    Case Else
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode报表编号
    '编制:刘兴洪
    '日期:2010-11-15 17:11:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng组ID As Long
    If Not Me.lvwGroups.SelectedItem Is Nothing Then
        lng组ID = Val(Mid(Me.lvwGroups.SelectedItem.Key, 2))
    End If
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "组ID=" & lng组ID)
End Sub


Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub
