VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAuditItem 
   Caption         =   "病案审查标准"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   Icon            =   "frmAuditItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   11655
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   5520
      Index           =   2
      Left            =   315
      ScaleHeight     =   5520
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      Begin VB.PictureBox picTree 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   60
         ScaleHeight     =   1590
         ScaleWidth      =   2940
         TabIndex        =   14
         Top             =   240
         Width           =   2940
         Begin MSComctlLib.TreeView tvwAuditType 
            Height          =   1200
            Left            =   495
            TabIndex        =   16
            Top             =   420
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   2117
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "ils16"
            Appearance      =   0
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "审查标准"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   15
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox pic方案信息 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   45
         ScaleHeight     =   1695
         ScaleWidth      =   2790
         TabIndex        =   7
         Top             =   2565
         Width           =   2790
         Begin VB.PictureBox picFAXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2415
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   8
            Top             =   75
            Width           =   255
         End
         Begin VB.Label lbl分段线 
            BackStyle       =   0  'Transparent
            Caption         =   "分段线:"
            Height          =   195
            Left            =   225
            TabIndex        =   13
            Top             =   1035
            Width           =   2580
         End
         Begin VB.Label lbl总分 
            BackStyle       =   0  'Transparent
            Caption         =   "总分:"
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   705
            Width           =   2580
         End
         Begin VB.Label lbl启用时间 
            BackStyle       =   0  'Transparent
            Caption         =   "启用时间:"
            Height          =   195
            Left            =   225
            TabIndex        =   11
            Top             =   1365
            Width           =   2580
         End
         Begin VB.Label lbl方案名称 
            BackStyle       =   0  'Transparent
            Caption         =   "方案名称"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   10
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "方案信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   9
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   1
      Left            =   3240
      ScaleHeight     =   2865
      ScaleWidth      =   5880
      TabIndex        =   3
      Top             =   4215
      Width           =   5880
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   2025
         Left            =   690
         TabIndex        =   4
         Top             =   240
         Width           =   2700
         _Version        =   589884
         _ExtentX        =   4762
         _ExtentY        =   3572
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   0
      Left            =   4350
      ScaleHeight     =   2715
      ScaleWidth      =   5010
      TabIndex        =   2
      Top             =   825
      Width           =   5010
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditItem 
         Height          =   4695
         Left            =   105
         TabIndex        =   6
         Top             =   150
         Width           =   6270
         _cx             =   11060
         _cy             =   8281
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7875
      TabIndex        =   0
      ToolTipText     =   "快捷键：F3"
      Top             =   90
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15690
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditItem.frx":08CA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditItem.frx":171C
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   2670
      Picture         =   "frmAuditItem.frx":1990
      Top             =   11235
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5535
      Picture         =   "frmAuditItem.frx":19E5
      Top             =   11250
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   210
      Picture         =   "frmAuditItem.frx":1A34
      Top             =   11190
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   3060
      Picture         =   "frmAuditItem.frx":1BF4
      Top             =   11205
      Visible         =   0   'False
      Width           =   2790
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3135
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAuditItem.frx":1DB2
      Left            =   1185
      Top             =   195
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmAuditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long      '锁定控件，不刷新
Private mstrPrivs               As String               '权限串
Private mobjFindKey             As CommandBarPopup      '查询
Private mstrFindKey             As String               '查询串
Private mlngModule              As Long                 '模块号
Private mstrSaveKey             As String               '保存的上次的分类选择关键字
Private mRsAuditItem            As ADODB.Recordset      '数据集
Private mblnCheckAll            As Boolean              '是否显示下级
Private zlCheck                 As New clsCheck         '检测类
Private mzlPrintModeS           As gzlPrintModeS        '打印
Private mstrSortID              As String               '排序定位
Private mblnProgUsed            As Boolean              '方案是否已使用
Private mlngCurFAID             As Long                 '当前方案ID
Private mintTypeID              As Integer              '分类新增、修改、删除时的ID
Private mintItemID              As Integer              '项目新增、修改、删除时的ID
Private mDataChange             As Boolean              '数据是否被编辑过
Private menuEditMode            As 编辑模式
Private mblPopType              As Boolean
Private mcbrPopupBarType        As CommandBar           '弹出窗口【分类】
Private mcbrPopupBarItem        As CommandBar           '弹出窗口【项目】
Dim cbrPopupItem                As CommandBarControl    '弹出项
Private Const con_vsfField = "/*+ rule */ '' as 图标,a.id, a.分类id,a.编码,a.名称,a.简码,a.分值,a.分制,b.名称 as 分类,decode(a.适用对象,1,'住院医嘱',2,'住院病历',3,'护理病历',4,'护理记录',5,'首页记录',6,'医嘱报告',7,'疾病证明',8,'知情文件','未定义') as 适用对象,a.说明,a.审查依据,适用对象 as 适用编码,文件ID,适用环节,数据源"
Private Const conFieldFiles = "Select /*+ rule */ a.id as 文件ID,a.编号 as 文件编码,a.名称 as 文件名称,a.说明 as 文件说明" & vbCrLf & _
                         "from 病历文件列表 A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B " & vbCrLf & _
                         "where /*+ rule */a.id = b.COLUMN_VALUE And a.种类 = [2]"
Private Const conEmrField = "Select /*+ Rule*/ Rawtohex(b.Id) As 文件id, b.Code As 文件编码, b.Title As 文件名称, b.Note As 文件说明" & vbNewLine & _
                        "From (Select Hextoraw(Column_Value) As ID From Table(Zlcommunal.f_Str2list(:p0, ','))) A, Antetype_List B" & vbNewLine & _
                        "Where Hextoraw(a.Id) = b.Id And b.Kind = :p1" & vbNewLine & _
                        "Order By 文件编码"
Public Enum 编辑模式
    浏览 = 0
    新增 = 1
    修改 = 2
    复制新增 = 3
End Enum

'打印模式
Enum gzlPrintModeS
    zlPrint = 1         '打印
    zlView = 2          '查看
    zlExcel = 3         '输出到Excel
End Enum

'树节点定位
Dim nod                         As Node
Dim i                           As Long
Dim FirstKey                    As String
Dim v                           As Variant

'==============================================================================
'=功能： 初始菜单工具栏
'==============================================================================
Private Sub InitCommandBar()
    Dim objMenu         As CommandBarPopup
    Dim objBar          As CommandBar
    Dim objExtendedBar  As CommandBar
    Dim objPopup        As CommandBarPopup
    Dim objControl      As CommandBarControl
    Dim cbrCustom       As CommandBarControlCustom
    
    On Error GoTo ErrH

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewKind, "新增方案(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyKind, "修改方案(&F)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteKind, "删除方案(&L)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Show, "导入方案(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Hide, "导出方案(&C)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "选用方案(&S)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "增加(&A)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "修改(&E)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "删除(&X)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "增加项目(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改项目(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_CopyNewItem, "复制增加(&C)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除项目(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "保存更改(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消更改(&R)")
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_ShowAll, "所有下级(&A)", True)
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Find, "过滤(&F)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)

    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '主菜单右侧的查找
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = Trim(GetPara("定位依据", mlngModule, "名称", True))
    If mstrFindKey = "" Then mstrFindKey = "名称"

    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.名称", , , "名称")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.编码", , , "编码")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.简码", , , "简码")
    
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon

    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "保存", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '查找
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyE, conMenu_Edit_Modify          '修改
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete       '删除
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save     '保存
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add 0, vbKeyF4, conMenu_View_Option                '选择定位依据
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
    End With
    '------------------------------------------------------------------------------------------------------------------
    '弹出菜单分类
    Set mcbrPopupBarType = cbsMain.Add("弹出分类菜单", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_NewKind, "新增方案(&N)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_ModifyKind, "修改方案(&F)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_DeleteKind, "删除方案(&L)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_View_Show, "导入方案(&P)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_View_Hide, "导出方案(&C)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_Select, "选用方案(&S)")
    
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "增加(&A)...", True)
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改(&E)...")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除(&X)")
    
    Set mcbrPopupBarItem = cbsMain.Add("弹出项目菜单", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "增加(&N)", True)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "复制(&C)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存(&S)", True)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消(&R)")
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始区域划分
'==============================================================================
Private Sub InitDockPannel()
    Dim objPane As Pane

    On Error GoTo ErrH
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "分类"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 400, 100, DockRightOf, Nothing)
    objPane.Title = "事件"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 300, 100, DockBottomOf, objPane)
    objPane.Title = "详细"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    
    Call DockPannelInit(dkpMain)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 控件初始化
'==============================================================================
Private Sub InitControl()
    
    On Error GoTo ErrH
    
    Call InitVsflexGrid
    Call InitCommandBar
    Call InitDockPannel
    Call InitTabControl
    Call InitTreeView
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始化网格 VsflexGrid
'==============================================================================
Private Sub InitVsflexGrid()
    Dim strField        As String
    Dim strFieldWidth   As String
    Dim varField        As Variant
    Dim varFieldWidth   As Variant
    Dim i               As Integer
    On Error GoTo ErrH
    vsfAuditItem.FocusRect = flexFocusNone
    vsfAuditItem.ExtendLastCol = True
    vsfAuditItem.ExplorerBar = flexExSortShowAndMove
    vsfAuditItem.AutoResize = False
    gstrSQL = "" & _
        "Select " & con_vsfField & vbCrLf & _
        "From 病案审查目录 a,(SELECT /*+ rule */ id,名称 FROM 病案审查分类 START WITH id=[1] CONNECT BY PRIOR ID = 上级ID)b " & vbCrLf & _
        "Where a.分类id = b.ID and 1=0"
    Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, -1)
    Set vsfAuditItem.DataSource = mRsAuditItem
    With vsfAuditItem
        .ColWidth(0) = 250
        .MergeCol(.ColIndex("分类id")) = True
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColWidth(.ColIndex("图标")) = 450
        .ColWidth(.ColIndex("适用对象")) = 2000
        .ColWidth(.ColIndex("分值")) = 500
        .ColWidthMin = 450
        
'        .FrozenCols = 3
        If GetPersonSet Then
            '使用个性化设置【调已保存的格式】
            strField = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name & "\VSFlexGrid", vsfAuditItem.Name & "名称", "")
            strFieldWidth = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name & "\VSFlexGrid", vsfAuditItem.Name & "宽度", "")
            varField = Split(strField, ",")
            varFieldWidth = Split(strFieldWidth, ",")
            For i = 0 To UBound(varField)
                If varField(i) <> "" And Val(varFieldWidth(i)) <> 0 Then
                    .ColPosition(.ColIndex(varField(i))) = i
                    .ColWidth(i) = Val(varFieldWidth(i))
                End If
            Next
        End If
        .ColWidth(.ColIndex("ID")) = 0: .ColHidden(.ColIndex("ID")) = True
        .ColWidth(.ColIndex("分类id")) = 0: .ColHidden(.ColIndex("分类id")) = True
        .ColWidth(.ColIndex("适用编码")) = 0: .ColHidden(.ColIndex("适用编码")) = True
        .ColWidth(.ColIndex("审查依据")) = 0: .ColHidden(.ColIndex("审查依据")) = True
        .ColWidth(.ColIndex("文件ID")) = 0: .ColHidden(.ColIndex("文件ID")) = True
        .ColWidth(.ColIndex("适用环节")) = 0: .ColHidden(.ColIndex("适用环节")) = True
        .ColWidth(.ColIndex("分制")) = 0: .ColHidden(.ColIndex("分制")) = True
        .ColWidth(.ColIndex("数据源")) = 0: .ColHidden(.ColIndex("数据源")) = True
    End With
    DoEvents
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 病案审查分类
'==============================================================================
Private Sub InitTreeView()
    Dim rsTree      As ADODB.Recordset
    Dim intStartid As Integer
    On Error GoTo ErrH

    'Tree的初始化
    Set tvwAuditType.ImageList = GetImageList(16)
    tvwAuditType.Nodes.Clear
    
    gstrSQL = "Select ID,名称,启用时间 From 病案审查方案"
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    Do Until rsTree.EOF
        If zlCommFun.NVL(rsTree!启用时间) <> "" Then
            intStartid = rsTree!ID
        End If
        Set nod = tvwAuditType.Nodes.Add(, , "Root" & rsTree!ID, zlCommFun.NVL(rsTree!名称, "默认方案"), 20, 20)
        nod.Expanded = True
            
        rsTree.MoveNext
    Loop
    
'    '添加根节点
'    Set nod = tvwAuditType.Nodes.Add(, , "Root", "分类", 20, 20)
'    nod.Expanded = True

    gstrSQL = "SELECT /*+ rule */ id,上级ID,方案ID,编码,名称 FROM 病案审查分类 START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID"
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    If rsTree.RecordCount = 0 Then Exit Sub
    rsTree.Sort = "编码"
    i = 1
    Do Until rsTree.EOF
        '添加子节点
        Set nod = tvwAuditType.Nodes.Add(IIf("" & rsTree("上级ID") = "", "Root" & rsTree("方案ID"), "A" & rsTree("上级ID")), tvwChild, "A" & rsTree("ID"), "【" + "" & rsTree("编码") + "】" + "" & rsTree("名称"), 23, 24)
        If i = 1 Then FirstKey = nod.Key
        If FirstKey = nod.Key Then i = 2
        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
        rsTree.MoveNext
    Loop
    FirstKey = "A" & mintTypeID
    For Each v In tvwAuditType.Nodes
        If v.Key = FirstKey Then
            '设置选中
            v.Selected = True
            v.EnsureVisible
        End If
    Next
    If tvwAuditType.SelectedItem Is Nothing Then
        tvwAuditType.Nodes("Root" & intStartid).Selected = True
        tvwAuditType.Nodes("Root" & intStartid).Bold = True
        tvwAuditType.Nodes("Root" & intStartid).Tag = 1
    End If
    DoEvents
    tvwAuditType_NodeClick tvwAuditType.SelectedItem
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Err.Clear
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始Tab控件
'==============================================================================
Private Function InitTabControl() As Boolean
    
    On Error GoTo ErrH
    
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With
        Set .Icons = zlCommFun.GetPubIcons
        .InsertItem 0, " 基本信息 ", frmAuditItemEdit.hWnd, 0
        .Item(0).Selected = True
    End With

    InitTabControl = True

    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 位置设置
'==============================================================================
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


'==============================================================================
'=功能： 界面分割
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo ErrH
    
    Select Case Item.ID
        Case 1
            Item.Handle = picPane(2).hWnd
        Case 2
            Item.Handle = picPane(0).hWnd
        Case 3
            Item.Handle = picPane(1).hWnd
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 排序后定位记录 vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
    lngRow = vsfAuditItem.FindRow(mstrSortID, -1, vsfAuditItem.ColIndex("ID"), False, True)
    If lngRow > 0 Then vsfAuditItem.Row = lngRow
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 某列不能移动位置 vsfAuditItem[图标]
'==============================================================================
Private Sub vsfAuditItem_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = vsfAuditItem.ColIndex("图标") Then
        Position = -1
    Else
        If Position <= vsfAuditItem.ColIndex("图标") Then Position = Col
    End If
End Sub

'==============================================================================
'=功能： 排序前记录ID vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_BeforeSort(ByVal Col As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSortID = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 某列不能拖动大小 vsfAuditItem[图标]
'==============================================================================
Private Sub vsfAuditItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfAuditItem.ColIndex("图标") Then Cancel = True
End Sub

'==============================================================================
'=功能： 双击完成修改功能 vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_DblClick()
    On Error GoTo ErrH
    If vsfAuditItem.MouseRow <= 0 Then Exit Sub
    Call ExecuteCommand("修改项目")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 右键菜单 vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '弹出菜单处理
        
            Call SendLMouseButton(vsfAuditItem.hWnd, X, Y)

            mcbrPopupBarItem.ShowPopup
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
'==============================================================================
'=功能：行列变换时
'==============================================================================
Private Sub vsfAuditItem_RowColChange()
    Dim rsTemp          As ADODB.Recordset
    Dim varPos          As Variant, strReturn As String
    On Error GoTo ErrH
    DoEvents
    If vsfAuditItem.Rows = 1 Then
        With frmAuditItemEdit
            .txtTypeID.Tag = "-1"
            .txtTypeID.Text = ""
            .txtName.Text = ""
            .txtCode.Text = ""
            .txtMnemonicCode.Text = ""
            .cboUsed.ListIndex = -1
            .cboLink.ListIndex = -1
            .txtDescription.Text = ""
            .txtAudit_NotCheck.Text = ""
            .txtNumValue = ""
            .CboPalValue.ListIndex = -1
            .blnProgUsed = False
            Set .vsfFiles.DataSource = Nothing
        End With
        stbThis.Panels(2) = "当前显示有 0 个项目。"
        frmAuditItemEdit.vsfFiles.Rows = 1
        Exit Sub
    End If
    If vsfAuditItem.ColIndex("ID") <= 0 Then Exit Sub
    If Val(vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))) <= 0 Then
        frmAuditItemEdit.vsfFiles.Rows = 1
        Exit Sub
    End If
    With frmAuditItemEdit
        
        .txtTypeID.Tag = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
        
        gstrSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, CStr(Val("" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("分类ID")))))
        If Not zlCheck.Connection_ChkRsState(rsTemp) Then
            .txtTypeID.Tag = "" & rsTemp!ID
            .txtTypeID.Text = "[" + rsTemp!编码 + "]" & rsTemp!名称
        Else
            .txtTypeID.Tag = "-1"
            .txtTypeID.Text = "[全部]分类"
        End If
        
        If vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("数据源")) = "0" Then
            .optSource(0).Value = True:             .optSource(1).Value = False
        Else
            .optSource(0).Value = False:             .optSource(1).Value = True
        End If
        .txtName.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("名称"))
        .txtCode.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("编码"))
        .txtMnemonicCode.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("简码"))
        .cboUsed.ListIndex = zlCheck.Cmb_EditIndex(.cboUsed, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("适用编码")))
        .cboLink.ListIndex = zlCheck.Cmb_EditIndex(.cboLink, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("适用环节")))
        .txtDescription.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("说明"))
        .txtAudit_NotCheck.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("审查依据"))
        .txtFileID.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("文件ID"))
        .txtNumValue.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("分值"))
        .CboPalValue.ListIndex = IIf(vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("分制")) = "", 0, vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("分制")))
        
        .blnProgUsed = mblnProgUsed
        If .optSource(0).Value Then
            gstrSQL = conFieldFiles
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .txtFileID.Text, AuditFileTran(zlCheck.Cmb_ID(.cboUsed), 0))
        Else
            gstrSQL = conEmrField
            strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, .txtFileID.Text & "^" & DbType.T_String & "^p0|" & AuditFileTran(zlCheck.Cmb_ID(.cboUsed), 1) & "^" & DbType.T_String & "^p1", rsTemp)
            If strReturn <> "" Then
                zlCheck.Msg_OK strReturn
                Exit Sub
            End If
        End If
        Set .vsfFiles.DataSource = rsTemp
    End With
    stbThis.Panels(2) = "当前显示有 " & vsfAuditItem.Rows - 1 & " 个项目。"
    varPos = zlCheck.Connection_GetBookMark(mRsAuditItem, "ID=" & CStr("" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))))
    DoEvents
    If Not IsNull(varPos) Then
        If Val(varPos) > 0 Then mRsAuditItem.Bookmark = varPos
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格数据加载 vsfAuditItem
'==============================================================================
Private Sub DataAuditItem(Optional strWhere As String)
    Dim strKey      As String
    Dim i           As Long
    Dim nTmpNode As Node
    
    On Error GoTo ErrH
    If strWhere = "" Then
        If ObjPtr(tvwAuditType.SelectedItem) = 0 Then Exit Sub
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            If tvwAuditType.SelectedItem.Tag = "1" Then
                mblnProgUsed = True
            Else
                mblnProgUsed = False
            End If
            Exit Sub
        End If
        
        
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            mlngCurFAID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
            If tvwAuditType.SelectedItem.Tag = "1" Then
                mblnProgUsed = True
            Else
                mblnProgUsed = False
            End If
        Else
            Set nTmpNode = tvwAuditType.SelectedItem
            While Not nTmpNode.Parent Is Nothing
                Set nTmpNode = nTmpNode.Parent
            Wend
            
            If InStrRev(nTmpNode.Key, "Root") > 0 Then
                mlngCurFAID = Replace(nTmpNode.Key, "Root", "")
                If nTmpNode.Tag = "1" Then
                    mblnProgUsed = True
                Else
                    mblnProgUsed = False
                End If
            End If
        End If
        
        
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            strKey = Mid(tvwAuditType.SelectedItem.Key, 5)
        Else
            strKey = Mid(tvwAuditType.SelectedItem.Key, 2)
        End If
        
        If mblnCheckAll Then
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From 病案审查目录 a,(SELECT /*+ rule */ id,名称 FROM 病案审查分类 START WITH id=[1] CONNECT BY PRIOR ID = 上级ID)b " & vbCrLf & _
                    "Where a.分类id = b.ID"
        Else
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From 病案审查目录 a,病案审查分类 b" & vbCrLf & _
                    "Where a.分类id = b.ID and a.分类id=[1]"
        End If
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, strKey)
    Else
        gstrSQL = "" & _
                "Select " & con_vsfField & vbCrLf & _
                "From 病案审查目录 a,病案审查分类 b" & vbCrLf & _
                "Where a.分类id = b.ID And" & vbCrLf & strWhere
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    End If
    Set vsfAuditItem.DataSource = mRsAuditItem
       
    With vsfAuditItem
        If .Rows > 1 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpPictureAlignment, i, .ColIndex("图标")) = flexPicAlignCenterCenter
                Select Case .Cell(flexcpText, i, .ColIndex("适用编码"))
                    Case "1"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(2).Picture
                    Case "2"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(15).Picture
                    Case "3"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(16).Picture
                    Case "4"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(17).Picture
                    Case "5"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(18).Picture
                    Case "6"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(6).Picture
                    Case "7"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(3).Picture
                    Case "8"
                        .Cell(flexcpPicture, i, .ColIndex("图标")) = frmPubResource.ils16.ListImages(20).Picture
                End Select
            Next i
            .Row = 1
        End If
    End With
    Call DataUpdate
    Call vsfAuditItem_RowColChange
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 添加分类数据 TypeInsert
'==============================================================================
Public Sub TypeInsert()

    Dim nTmpNode As Node
    
    On Error GoTo ErrH
    
    With frmAuditItemTypeEdit
        .EditMode = 新增
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            .strID = "-1"
            .lngProjectID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
            .strProjectName = tvwAuditType.SelectedItem.Text
        Else
            .strID = Mid(tvwAuditType.SelectedItem.Key, 2)
            
            Set nTmpNode = tvwAuditType.SelectedItem
            While Not nTmpNode.Parent Is Nothing
                Set nTmpNode = nTmpNode.Parent
            Wend
            
            .lngProjectID = Replace(nTmpNode.Key, "Root", "")
            .strProjectName = nTmpNode.Text
        End If
        
        .Show vbModal
        If .blnCancel Then Set frmAuditItemTypeEdit = Nothing: Exit Sub
        mintTypeID = .strID
    End With
    Set frmAuditItemTypeEdit = Nothing
    '刷新树
    Call InitTreeView
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 修改分类数据 TypeUpdate
'==============================================================================
Public Sub TypeUpdate()

    Dim nTmpNode As Node
    
    On Error GoTo ErrH
    If InStrRev(tvwAuditType.SelectedItem.Key, "Root") > 0 Then
        MsgBox "根目录,不能修改。", vbInformation, "中联提示"
        Exit Sub
    End If
    With frmAuditItemTypeEdit
        .EditMode = 修改
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            .strID = Mid(tvwAuditType.SelectedItem.Key, 5)
            .lngProjectID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
            .strProjectName = tvwAuditType.SelectedItem.Text
        Else
            .strID = Mid(tvwAuditType.SelectedItem.Key, 2)
            
            Set nTmpNode = tvwAuditType.SelectedItem
            While Not nTmpNode.Parent Is Nothing
                Set nTmpNode = nTmpNode.Parent
            Wend

            .lngProjectID = Replace(nTmpNode.Key, "Root", "")
            .strProjectName = nTmpNode.Text
        End If
        
        .Show vbModal
        If .blnCancel Then Set frmAuditItemTypeEdit = Nothing: Exit Sub
    End With
    Set frmAuditItemTypeEdit = Nothing
    If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
        mintTypeID = Mid(tvwAuditType.SelectedItem.Key, 5)
    Else
        mintTypeID = Mid(tvwAuditType.SelectedItem.Key, 2)
    End If

    '刷新树
    Call InitTreeView
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 添加分类数据 TypeUpdate
'==============================================================================
Public Sub TypeDelete()
    Dim strKey      As String
    
    On Error GoTo ErrH
    If InStrRev(tvwAuditType.SelectedItem.Key, "Root") > 0 Then
        MsgBox "根目录，不能删除！", vbInformation, "中联提示"
        Exit Sub
    Else
        If MsgBox("确认删除分类""" & tvwAuditType.SelectedItem.Text & """及其所属项目吗？", vbOKCancel + vbDefaultButton2, "中联提示") <> vbOK Then Exit Sub
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            mintTypeID = Val(Mid(tvwAuditType.SelectedItem.Key, 5))
        Else
            mintTypeID = Val(Mid(tvwAuditType.SelectedItem.Key, 2))
        End If
    End If
    frmAuditItemTypeEdit.strID = CStr(mintTypeID)
    frmAuditItemTypeEdit.AuditItemTypeDelete
    Set frmAuditItemTypeEdit = Nothing
    tvwAuditType.Nodes.Remove tvwAuditType.SelectedItem.Index
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 打印 ItemPrint
'==============================================================================
Private Sub ItemPrint()
    On Error GoTo ErrH
    subPrint (mzlPrintModeS)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim lngLoop         As Long
    Dim objControl      As Object
    Dim objPrint        As New zlPrint1Grd
    Dim objAppRow       As zlTabAppRow
    
    If vsfAuditItem Is Nothing Then Exit Sub
    LockWindowUpdate vsfAuditItem.hWnd
    vsfAuditItem.ColHidden(vsfAuditItem.ColIndex("图标")) = True
    Call SearchPrintData(vsfAuditItem, frmPubResource.msfPrint)
    vsfAuditItem.ColHidden(vsfAuditItem.ColIndex("图标")) = False
    LockWindowUpdate 0
    '调用打印部件处理
    Set objPrint.Body = frmPubResource.msfPrint
    objPrint.Title.Text = Me.Caption
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人：" & UserInfo.姓名)
    Call objAppRow.Add("打印时间：" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    zlPrintOrView1Grd objPrint, bytMode
End Sub

'==============================================================================
'=功能： 查询 ItemFind
'==============================================================================
Private Sub ItemFind()
    Dim strWhere        As String
    On Error GoTo ErrH
    With frmAuditItemFind
        .Show vbModal
        If .blnCancel Then Set frmAuditItemFind = Nothing: Exit Sub
        strWhere = .strWhere
    End With
    Set frmAuditItemFind = Nothing
    DataAuditItem (strWhere)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrH
    mstrPrivs = UserInfo.模块权限
    mlngModule = ParamInfo.模块号
    Call ExecuteCommand("初始控件")
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)
        
    menuEditMode = 浏览
    Call ExecuteCommand("读取病案审查项目") '    mblnProgUsed = False
    frmAuditItemEdit.WinLock
    picFAXX.Picture = imgClose.Picture
    
    Call ExecuteCommand("读注册表")
        
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrH
    menuEditMode = 浏览
    Call SetPaneRange(dkpMain, 1, 100, 60, 450, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 3, 15, 300, Me.ScaleWidth, 350)
    
    dkpMain.RecalcLayout
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo ErrH
    
    Call ExecuteCommand("写注册表")
    Call SaveWinState(Me, App.ProductName)
    Set mobjFindKey = Nothing
    Set frmAuditItemEdit = Nothing
    
    SaveFlexState vsfAuditItem, Me.Name
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsfAuditItem.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 1
        tbcPage.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 2
'        tvwAuditType.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        
        On Error Resume Next
        pic方案信息.Move 0, picPane(2).ScaleHeight - pic方案信息.Height, picPane(2).ScaleWidth
        With picTree
            .Move 0, 0, pic方案信息.Width, picPane(2).Height - pic方案信息.Height
            .Cls
            .PaintPicture imgBGBlue.Picture, 0, 0, picTree.Width, 360, 0, 0, imgBGBlue.Width, 360
            .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, picTree.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, picTree.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picTree.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, 0, picTree.ScaleHeight - Screen.TwipsPerPixelY, picTree.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
        End With
        
        tvwAuditType.Move Screen.TwipsPerPixelX * 4, 390, Abs(picTree.ScaleWidth - 8 * Screen.TwipsPerPixelX), Abs(picTree.ScaleHeight - 390 - Screen.TwipsPerPixelY * 4)
        With pic方案信息
            .Cls
            .PaintPicture imgBGBlue.Picture, 0, 0, pic方案信息.Width, 360, 0, 0, imgBGBlue.Width, 360
            .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, pic方案信息.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, 0, pic方案信息.ScaleHeight - Screen.TwipsPerPixelY, pic方案信息.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
        End With
        picFAXX.Move pic方案信息.ScaleWidth - picFAXX.Width - 80
        Refresh
        
    End Select
End Sub

Private Sub tvwAuditType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '弹出菜单处理
        
            Call SendLMouseButton(tvwAuditType.hWnd, X, Y)

            mcbrPopupBarType.ShowPopup
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwAuditType_NodeClick(ByVal Node As MSComctlLib.Node)
        
    If mstrSaveKey = Node.Key Then Exit Sub
    If Left(Node.Key, 4) = "Root" Then
        vsfAuditItem.Rows = 1
        mstrSaveKey = Node.Key
        mlngCurFAID = Replace(mstrSaveKey, "Root", "")
        If Node.Tag = "1" Then
            mblnProgUsed = True
        Else
            mblnProgUsed = False
        End If
        Call DataUpdate
        Exit Sub
    End If
    mstrSaveKey = Node.Key
    
    Call ExecuteCommand("读取病案审查项目")
    
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo ErrH
    
    lngRow = 0
    If txtLocation.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '读取大于当前行的记录数据
        For lngLoop = vsfAuditItem.Row + 1 To vsfAuditItem.Rows - 1
            If InStr(UCase(vsfAuditItem.TextMatrix(lngLoop, vsfAuditItem.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '读取小于当前行的记录数据
        If lngRow = 0 Then
            For lngLoop = 0 To vsfAuditItem.Row
                If InStr(UCase(vsfAuditItem.TextMatrix(lngLoop, vsfAuditItem.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If vsfAuditItem.Rows > 1 And lngRow >= 1 Then vsfAuditItem.Row = lngRow
        
        Call LocationObj(txtLocation)
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
    
    On Error GoTo ErrH
    Dim strF As String
    Dim strTvwName As String
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        Call InitControl
    Case "读取病案审查项目"
        Call DataAuditItem
    Case "增加分类"
        Call TypeInsert
    Case "修改分类"
        Call TypeUpdate
    Case "删除分类"
        Call TypeDelete
    Case "增加项目"
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            frmAuditItemEdit.lngItemTypeID = Mid(tvwAuditType.SelectedItem.Key, 5)
        Else
            frmAuditItemEdit.lngItemTypeID = Mid(tvwAuditType.SelectedItem.Key, 2)
        End If
        Call frmAuditItemEdit.ItemInsert
    Case "修改项目"
        frmAuditItemEdit.lngItemID = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
        Call frmAuditItemEdit.ItemUpdate
    Case "删除项目"
        Dim varPos      As Variant
        frmAuditItemEdit.lngItemID = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
        frmAuditItemEdit.strItemCode = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("编码"))
        frmAuditItemEdit.strItemName = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("名称"))
        
        Call frmAuditItemEdit.ItemDelete
        
        
        varPos = vsfAuditItem.Row
        Call DataAuditItem
        Call vsfAuditItem_RowColChange
        If varPos <= vsfAuditItem.Rows - 1 Then vsfAuditItem.Row = varPos
        
    Case "保存项目"
        If frmAuditItemEdit.ItemSave Then: Exit Function
        mstrSaveKey = "A" & frmAuditItemEdit.lngItemTypeID
        Dim lngRow      As Long
        '设置选择分类
        FirstKey = "A" & CStr(frmAuditItemEdit.lngItemTypeID)
        For Each v In tvwAuditType.Nodes
            If v.Key = FirstKey Then
                '设置选中
                v.Selected = True
                v.EnsureVisible
            End If
        Next
        frmAuditItemEdit.WinLock
        Call DataAuditItem
        DoEvents
        lngRow = vsfAuditItem.FindRow(CStr(frmAuditItemEdit.lngItemID), -1, vsfAuditItem.ColIndex("ID"), False, True)
        If lngRow > 0 Then vsfAuditItem.Row = lngRow

    Case "取消项目"
        If frmAuditItemEdit.ItemCancel Then Exit Function
        
        Call vsfAuditItem_RowColChange
        
    Case "复制增加项目"
        Call frmAuditItemEdit.ItemCopy
    Case "导入方案"
         '从XML文件导入
        Dim strXML As String
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        
        On Error GoTo err1
        dlgThis.ShowOpen
        
err1:
        If Err.Number = cdlCancel Then
           Err.Clear
           strXML = ""
           Exit Function
        End If
        
        strXML = dlgThis.FileName
        If gobjFSO.FileExists(strXML) Then
            If ImportFromXMLFile(Me.tvwAuditType, strXML) Then
                '刷新
                Call InitTreeView
            End If
        End If
        
    Case "导出方案"
        If ObjPtr(tvwAuditType.SelectedItem) = 0 Then Exit Function
        strTvwName = tvwAuditType.SelectedItem.Text
        
        dlgThis.FileName = "病案审查_" & strTvwName & "_方案.xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Function
        strF = dlgThis.FileName
        On Error GoTo ErrH
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Function
        End If
        
        If ExportToXMLFile(tvwAuditType, strF) Then
            DoEvents
            MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
        End If
        
    Case "前一条"
        With vsfAuditItem
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
    Case "后一条"
        With vsfAuditItem
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            End If
        End With
    Case "读注册表"
        mblnCheckAll = (Val(GetPara("所有下级", mlngModule, "0", False)) = 1)
        If GetPersonSet Then
            '使用个性化设置
            dkpMain.LoadStateFromString GetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, "")
            mstrFindKey = Trim(GetPara("定位依据", mlngModule, "名称", True))
            Call RestoreWinState(Me, App.ProductName)
        End If
    Case "写注册表"
        '使用个性化设置
        Call SetPara("定位依据", mstrFindKey, mlngModule)
        Call SetPara("所有下级", IIf(mblnCheckAll, 1, 0), mlngModule)
        Call SetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SaveWinState(Me, App.ProductName)
    End Select
    ExecuteCommand = True
    
    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_Edit_NewKind                       '新增方案
            Call ProgAdd
        Case conMenu_Edit_ModifyKind                    '修改方案
            Call ProgEdit
        Case conMenu_Edit_DeleteKind                    '删除方案
            Call ProgDel
        Case conMenu_View_Show                          '导入方案
            Call ExecuteCommand("导入方案")
        Case conMenu_View_Hide                          '导出方案
            Call ExecuteCommand("导出方案")
        Case conMenu_Edit_Select                        '选用方案
            Call ProgSele
        Case conMenu_Edit_NewParent                     '增加分类
            Call ExecuteCommand("增加分类")
        Case conMenu_Edit_ModifyParent                  '修改分类
            Call ExecuteCommand("修改分类")
        Case conMenu_Edit_DeleteParent                  '删除分类
            Call ExecuteCommand("删除分类")
        Case conMenu_Edit_NewItem                       '增加项目
            vsfAuditItem.Rows = vsfAuditItem.Rows + 1
            vsfAuditItem.Row = vsfAuditItem.Rows - 1
            Call ExecuteCommand("增加项目")
        Case conMenu_Edit_Modify                        '修改项目
            Call ExecuteCommand("修改项目")
        Case conMenu_Edit_CopyNewItem                   '复制增加项目
            vsfAuditItem.Rows = vsfAuditItem.Rows + 1
            vsfAuditItem.Row = vsfAuditItem.Rows - 1
            Call ExecuteCommand("复制增加项目")
        Case conMenu_Edit_Delete                        '删除项目
            Call ExecuteCommand("删除项目")
        Case conMenu_Edit_Transf_Save                   '保存项目
            Call ExecuteCommand("保存项目")
        Case conMenu_Edit_Transf_Cancle                  '取消项目
            blnNewCancel = frmAuditItemEdit.EditMode = 新增 Or frmAuditItemEdit.EditMode = 复制新增
            
            Call ExecuteCommand("取消项目")
            If blnNewCancel And frmAuditItemEdit.EditMode = 浏览 Then
                vsfAuditItem.Rows = vsfAuditItem.Rows - 1
                vsfAuditItem.Row = vsfAuditItem.Rows - 1
            End If
            
        Case conMenu_View_ShowAll                       '所有下级
            mblnCheckAll = Not mblnCheckAll
            Control.Checked = mblnCheckAll
            DataAuditItem
        Case conMenu_View_Find                          '搜索查找
            Call ItemFind
        Case conMenu_File_Preview   '预览
            mzlPrintModeS = zlView
            Call ItemPrint
        Case conMenu_File_Print   '打印
            mzlPrintModeS = zlPrint
            Call ItemPrint
        Case conMenu_File_Excel '输出到&Excel
            mzlPrintModeS = zlExcel
            Call ItemPrint
        Case conMenu_View_Forward
            Call ExecuteCommand("前一条")
        Case conMenu_View_Backward
            Call ExecuteCommand("后一条")
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txtLocation
        Case conMenu_View_Refresh               '刷新
            Dim lngRow      As Long
            If vsfAuditItem.Rows = 1 Then Exit Sub
            mintItemID = vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
            Call DataAuditItem
            lngRow = vsfAuditItem.FindRow(mintItemID, -1, vsfAuditItem.ColIndex("ID"), False, True)
            If lngRow > 0 Then vsfAuditItem.Row = lngRow
        Case Else
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '与业务无关的功能，公共的功能
                Call CommandBarExecutePublic(Control, Me, vsfAuditItem, "病案审查目录清单")
            End If
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo ErrH

    picPane(0).Enabled = (frmAuditItemEdit.EditMode = 浏览)
    picPane(2).Enabled = (frmAuditItemEdit.EditMode = 浏览)
    txtLocation.Locked = (frmAuditItemEdit.EditMode <> 浏览)
    With vsfAuditItem
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
                Control.Enabled = ((vsfAuditItem.Rows > 1) And IsPrivs(mstrPrivs, "增删改"))
            Case conMenu_EditPopup
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
            Case conMenu_Edit_NewParent
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Control.Visible And frmAuditItemEdit.EditMode = 浏览)
                End If
            Case conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent    '
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And frmAuditItemEdit.EditMode = 浏览)
                End If
            Case conMenu_Edit_NewItem                    '增加项目
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And frmAuditItemEdit.EditMode = 浏览)
                End If
            Case conMenu_Edit_Modify                        '修改项目
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And (Not zlCheck.Connection_ChkRsState(mRsAuditItem)) And frmAuditItemEdit.EditMode = 浏览)
                End If
            Case conMenu_Edit_Delete                  '删除
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And (Not zlCheck.Connection_ChkRsState(mRsAuditItem)) And frmAuditItemEdit.EditMode = 浏览)
                End If
            Case conMenu_Edit_CopyNewItem                   '复制增加
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And (Not zlCheck.Connection_ChkRsState(mRsAuditItem)) And frmAuditItemEdit.EditMode = 浏览)
                End If
            Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle '保存或取消保存
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = ((frmAuditItemEdit.EditMode <> 浏览) And Control.Visible)
            Case conMenu_View_Forward
                Control.Enabled = (frmAuditItemEdit.EditMode = 浏览) And vsfAuditItem.Row > 1
            Case conMenu_View_Backward
                Control.Enabled = (frmAuditItemEdit.EditMode = 浏览) And vsfAuditItem.Row + 1 < vsfAuditItem.Rows
            Case conMenu_View_Find, conMenu_View_Refresh
                Control.Enabled = (frmAuditItemEdit.EditMode = 浏览)
            Case conMenu_View_ShowAll                       '所有下级
                Control.Checked = mblnCheckAll
                Control.Enabled = (frmAuditItemEdit.EditMode = 浏览)
            Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
                If InStr(Control.Caption, mstrFindKey) > 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If
                
            Case conMenu_Edit_NewKind                       '新增方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_ModifyKind                    '修改方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvwAuditType.Nodes.count > 0
                
            Case conMenu_Edit_DeleteKind                    '删除方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvwAuditType.Nodes.count > 0
            Case conMenu_View_Show                          '导入方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvwAuditType.SelectedItem.Children = 0 And Left(tvwAuditType.SelectedItem.Key, 4) = "Root"
                End If
            Case conMenu_View_Hide                          '导出方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = Control.Visible And tvwAuditType.SelectedItem.Children > 0 And Left(tvwAuditType.SelectedItem.Key, 4) = "Root"
                End If
            Case conMenu_Edit_Select                        '选用方案
                Control.Enabled = tvwAuditType.Nodes.count > 0
            Case Else
                Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
End Sub

'==============================================================================
'=功能： 方案信息关闭或显示
'==============================================================================
Private Sub picFAXX_Click()
    On Error GoTo ErrH
    
    If picFAXX.Tag = "" Then
        picFAXX.Tag = "Opened"
        picFAXX.Picture = imgOpen.Picture
        pic方案信息.Height = 340
    Else
        picFAXX.Tag = ""
        picFAXX.Picture = imgClose.Picture
        pic方案信息.Height = 1695
    End If
    picFAXX.Refresh
    Call picPane_Resize(2)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 方案信息焦点变色
'==============================================================================
Private Sub picFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If X >= 0 And X <= picFAXX.ScaleWidth And Y >= 0 And Y <= picFAXX.ScaleHeight Then
        SetCapture picFAXX.hWnd
        '鼠标移入！！！
        picFAXX.Line (0, 0)-(picFAXX.ScaleWidth - Screen.TwipsPerPixelX, picFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '鼠标移出！！！
        picFAXX.Cls
        ReleaseCapture
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=增加审查方案
'==============================================================================
Private Sub ProgAdd()
    Dim f As New frm审查方案编辑
    On Error GoTo ErrH
    f.ShowForm   '新增
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '填充列表
'        Call DataLoad
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=修改审查方案
'==============================================================================
Private Sub ProgEdit()
    Dim f               As New frm审查方案编辑
    Dim lng总分         As Double
    On Error GoTo ErrH
    
'    mlngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If mlngCurFAID < 1 Then Exit Sub
    f.ShowForm mlngCurFAID   '修改，传入ID
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '填充列表
        Call DataAuditItem
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=删除审查方案
'==============================================================================
Private Sub ProgDel()
    Dim intIndex        As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrH
    
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If mlngCurFAID < 1 Then Exit Sub
    If MsgBox("你确认要删除该条方案吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '检查是否在病案反馈记录中存在
    Set rsTmp = gclsPackage.GetProjectUse(mlngCurFAID)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!条数 > 0 Then
            Call MsgBox("该方案已经被使用过,暂不能删除!", vbInformation, gstrSysName)
            Exit Sub
        End If
    End If
    gstrSQL = "ZL_病案审查方案_Delete(" & CStr(mlngCurFAID) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call InitTreeView
    Call DataAuditItem
    Call SetMenu
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=选用方案
'==============================================================================
Private Sub ProgSele()
    Dim intIndex        As Long
    Dim bln已使用       As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lngDefaultID As Long
    
    On Error GoTo ErrH
    
    If mlngCurFAID < 1 Then Exit Sub
    If MsgBox("注意：审查分案的选用是一件非常慎重的事情，通常不要随意更改！" & vbCrLf & "请确认选用本审查方案吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    For i = 1 To tvwAuditType.Nodes.count
        If tvwAuditType.Nodes(i).Bold Then
            lngDefaultID = Replace(tvwAuditType.Nodes(i).Key, "Root", "")
        End If
    Next
    
    Dim rsTemp As New ADODB.Recordset
    Set rsTmp = gclsPackage.GetProjectUse(lngDefaultID)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!条数 > 0 Then
            If MsgBox("注意：系统默认审查分案正在使用当中，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    gstrSQL = "ZL_病案审查方案_选用(" & CStr(mlngCurFAID) & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call InitTreeView
    Call SetMenu
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 数据统计
'==============================================================================
Private Sub DataUpdate()
    Dim rs              As ADODB.Recordset
    Dim lng总分         As Double
    On Error GoTo ErrH
    gstrSQL = "Select 名称,总分,分段线,启用时间,停用时间,说明 From 病案审查方案 where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurFAID)
    If Not rs.EOF Then
        lbl方案名称.Caption = rs("名称")
        lbl分段线.Caption = "分段线:" & rs("分段线")
        lbl启用时间.Caption = "启用时间:" & zlCommFun.NVL(rs("启用时间"))
        lbl总分.Caption = "总分:" & rs("总分")
        lng总分 = rs("总分")
    Else
        lbl方案名称.Caption = ""
        lbl分段线.Caption = ""
        lbl启用时间.Caption = ""
        lbl总分.Caption = ""
    End If
    
'''    gstrSQL = "select sum(标准分值) from 病案评分标准 where 上级ID is null and 方案ID = [1]"
'''    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
'''    If Not rs.EOF Then
'''        If Abs(lng总分 - rs.Fields(0)) > 0.01 Then
'''            lbl总分 = lbl总分 + "，项目分数和为:" & rs.Fields(0)
'''            lbl总分.ForeColor = vbRed
'''        Else
'''            lbl总分.ForeColor = vbBlack
'''        End If
'''    Else
'''        lbl总分.ForeColor = vbRed
'''    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 显示记录数信息
'==============================================================================
Private Sub SetMenu()
    On Error GoTo ErrH
    stbThis.Panels(2).Text = "列表中共显示有" & vsfAuditItem.Rows - 1 & "行数据。"
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
