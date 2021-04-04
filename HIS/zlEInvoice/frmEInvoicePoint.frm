VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEInvoicePoint 
   BorderStyle     =   0  'None
   Caption         =   "电子票据开票点"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   2868
      Left            =   840
      ScaleHeight     =   2865
      ScaleWidth      =   4665
      TabIndex        =   4
      Top             =   1350
      Width           =   4668
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3225
         Left            =   2400
         MousePointer    =   9  'Size W E
         ScaleHeight     =   3225
         ScaleMode       =   0  'User
         ScaleWidth      =   22.5
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   35
      End
      Begin VB.PictureBox picFun 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   2520
         ScaleHeight     =   2655
         ScaleWidth      =   2415
         TabIndex        =   9
         Top             =   120
         Width           =   2415
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   615
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   855
            _Version        =   589884
            _ExtentX        =   1508
            _ExtentY        =   1085
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picTree 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   2055
         TabIndex        =   7
         Top             =   120
         Width           =   2055
         Begin MSComctlLib.TreeView tvw开票点 
            Height          =   1485
            Left            =   -360
            TabIndex        =   8
            Top             =   480
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   2619
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "ils16"
            Appearance      =   1
         End
      End
   End
   Begin VB.PictureBox pic电子票据对码 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   6264
      ScaleHeight     =   1935
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
      Begin VSFlex8Ctl.VSFlexGrid vs对码明细 
         Height          =   1080
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "开票点对码明细"
         Top             =   240
         Width           =   1995
         _cx             =   3519
         _cy             =   1905
         Appearance      =   0
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEInvoicePoint.frx":0000
         ScrollTrack     =   0   'False
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
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
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
   Begin VB.PictureBox pic电子票据设置 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   6360
      ScaleHeight     =   2175
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
      Begin VB.PictureBox picSplitH 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   3000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3000
      End
      Begin VSFlex8Ctl.VSFlexGrid vs对照 
         Height          =   1080
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "开票点对码"
         Top             =   1080
         Width           =   1995
         _cx             =   3519
         _cy             =   1905
         Appearance      =   0
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEInvoicePoint.frx":016D
         ScrollTrack     =   0   'False
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
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vs开票点 
         Height          =   1080
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "开票点对码"
         Top             =   0
         Width           =   1995
         _cx             =   3519
         _cy             =   1905
         Appearance      =   0
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEInvoicePoint.frx":02DA
         ScrollTrack     =   0   'False
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
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEInvoicePoint.frx":0467
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEInvoicePoint.frx":1DF9
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   3720
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
            Picture         =   "frmEInvoicePoint.frx":378B
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEInvoicePoint.frx":419D
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1032
      Left            =   0
      Top             =   648
      Width           =   528
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "基础数据设置"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmEInvoicePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng编码长度 As Long
Private mintColumn As Integer
Private mstrKey As String       '前一个树节点的关键值
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mint焦点 As Integer
Private mblnShowStop As Boolean  '显示停用
Private mblnShowAll As Boolean  '显示所有下级
Public mint对码方式 As Integer  '0-按客户端对,1-按收费员对;2-按收费员+客户端对
Private Enum mFocus
    Focus_None = 0
    Focus_开票点分类 = 1
    Focus_开票点 = 2
    Focus_开票点对照 = 3
    Focus_对照明细 = 4
End Enum
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long
Dim sngStartX As Single, sngStartY As Single    '移动前鼠标的位置

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo ErrHandler
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        With cbrMenuBar.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        End With
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "停用(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Add, "新增对码(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Modify, "修改对码(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Delete, "删除对码(&D")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示已停用开票点(&P)", cbrControl.index)
        cbrControl.Checked = mblnShowStop
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "显示所有下级(&H)", cbrControl.index)
        cbrControl.Checked = mblnShowAll
        cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&E)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&R)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "停用(&P)", cbrControl.index + 1): cbrControl.BeginGroup = True
        .Item(cbrControl.index + 1).BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Add, "新增对码(&A)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Modify, "修改对码(&U)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Delete, "删除对码(&D)", cbrControl.index + 1)
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("E"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add FCONTROL, Asc("P"), conMenu_Edit_Pause
        .Add FCONTROL, Asc("A"), conMenu_Edit_Audit_Add
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit_Modify
        .Add FCONTROL, Asc("D"), conMenu_Edit_Audit_Delete
    End With
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID

    Case conMenu_File_Parameter '参数设置
        Set objfrmEInvoiceParaSet = New frmEInvoiceParaSet
        Call objfrmEInvoiceParaSet.ShowMe(Me, mlngSys, 1145)
        
         mint对码方式 = zlDatabase.GetPara("开票点对码方式", mlngSys, 1145, 1)
        Call Load开票点对照(0)
        Call Load开票点对码明细(0)
        
    Case conMenu_Edit_NewItem '新增
        Call AddNewEInvoicePoint
    Case conMenu_Edit_Modify  '修改
        Call ModifyEInvoicePoint
    Case conMenu_Edit_Delete '删除
        Call DeleteEInvoicePoint
    Case conMenu_Edit_Reuse '启用
        Call StartEInvoicePoint
    Case conMenu_Edit_Pause '停用
        Call StopEInvoicePoint
    Case conMenu_Edit_Audit_Add '新增对码
        Call Set开票点对照
    Case conMenu_Edit_Audit_Modify '修改对码
        Call Set开票点对照(True)
    Case conMenu_Edit_Audit_Delete '删除对码
        Call Delete开票点对照
    Case conMenu_View_ShowStoped '显示停用的
         Control.Checked = Not Control.Checked
         mblnShowStop = Control.Checked
         Call load开票点分类
    Case conMenu_View_ShowAll '显示所有下级
         Control.Checked = Not Control.Checked
         mblnShowAll = Control.Checked
         Call load开票点分类
    Case conMenu_View_Refresh '刷新数据
        Call load开票点分类
    Case Else
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnable As Boolean
    Dim int焦点 As Integer
    Dim bln停用 As Boolean, bln对码 As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub

    If mint焦点 = Focus_开票点 Then
        blnEnable = Val(vs开票点.TextMatrix(vs开票点.Row, vs开票点.ColIndex("开票点ID"))) > 0
        bln停用 = Val(vs开票点.TextMatrix(vs开票点.Row, vs开票点.ColIndex("停用"))) = "1"
    End If
    
    If mint焦点 = Focus_开票点对照 Then
        bln对码 = Val(vs对照.TextMatrix(vs对照.Row, vs对照.ColIndex("ID"))) > 0
    End If
    If mint焦点 = Focus_对照明细 Then
        bln对码 = Val(vs对码明细.TextMatrix(vs对码明细.Row, vs对码明细.ColIndex("ID"))) > 0
    End If
    
    Select Case Control.ID
     Case conMenu_Edit_NewItem
        Control.Enabled = mint焦点 = Focus_开票点分类 Or mint焦点 = Focus_开票点
    Case conMenu_Edit_Modify
        If mint焦点 = Focus_开票点 Then
            Control.Enabled = blnEnable
        ElseIf mint焦点 = Focus_开票点分类 Then
            Control.Enabled = tvw开票点.SelectedItem.Key <> "Root"
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Delete
        If mint焦点 = Focus_开票点 Then
            Control.Enabled = blnEnable
        ElseIf mint焦点 = Focus_开票点分类 Then
            Control.Enabled = tvw开票点.SelectedItem.Image <> "Root"
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Reuse
        If mint焦点 = Focus_开票点 Then
            If Not blnEnable Then
                Control.Enabled = False
            Else
                Control.Enabled = bln停用
            End If
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Pause
        If mint焦点 = Focus_开票点 Then
            If Not blnEnable Then
                Control.Enabled = False
            Else
                Control.Enabled = Not bln停用
            End If
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Audit_Add
        Control.Enabled = mint焦点 = Focus_开票点对照 Or mint焦点 = Focus_对照明细
    Case conMenu_Edit_Audit_Modify, conMenu_Edit_Audit_Delete
        Control.Enabled = (mint焦点 = Focus_开票点对照 Or mint焦点 = Focus_对照明细) And bln对码
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        Control.Enabled = False
    Case Else
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    Call InitPage
    mblnShowStop = GetSetting("ZLSOFT", "私有模块\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示停用开票点", 0)
    mblnShowAll = GetSetting("ZLSOFT", "私有模块\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示所有下级开票点", 0)
    mint对码方式 = zlDatabase.GetPara("开票点对码方式", 100, 1145, 1)

    Call load开票点分类
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Public Sub AddNewEInvoicePoint()
    '新增
    Dim strKey As String
    Dim bln末级 As Boolean, blnRefresh As Boolean
    Dim frmEdit As New frmEInvoicePointSet

    On Error GoTo errHandle
    If tvw开票点.SelectedItem Is Nothing Then Exit Sub
    strKey = Mid(tvw开票点.SelectedItem.Key, 2)
    bln末级 = mint焦点 = Focus_开票点
    Call frmEdit.Init开票点设置("", strKey, bln末级, blnRefresh)
    If blnRefresh Then Call load开票点分类
  
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DeleteEInvoicePoint()
    '删除
    On Error GoTo errHandle
    Dim strKey As String, strSQL As String
    Dim intIndex As Long
    Dim strTemp As String
    
    If mint焦点 = Focus_开票点分类 Then
        If tvw开票点.SelectedItem Is Nothing Then Exit Sub
        strKey = tvw开票点.SelectedItem.Key
        If strKey = "Root" Then Exit Sub
        strTemp = Val(Mid(tvw开票点.SelectedItem.Key, 2))
    
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "该电子票据开票点下级存在其他开票点，不能删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你确认要删除名称为“" & tvw开票点.SelectedItem.Text & "”的电子票据开票点吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        MousePointer = 11
        strSQL = "zl_电子票据开票点_DELETE(" & strTemp & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        MousePointer = 0
    Else
        With vs开票点
            If .Row = 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("开票点ID"))) = 0 Then Exit Sub
            strTemp = Val(.TextMatrix(.Row, .ColIndex("开票点ID")))
            If CheckExistDepPres(strTemp) = True Then
                MsgBox "该电子票据开票点下级存在其他开票点，不能删除。", vbInformation, gstrSysName
                Exit Sub
            End If
            If MsgBox("你确认要删除名称为“" & vs开票点.TextMatrix(vs开票点.Row, vs开票点.ColIndex("名称")) & "”的电子票据开票点吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Me.MousePointer = 11
            strSQL = "Zl_电子票据开票点_DELETE(" & strTemp & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Me.MousePointer = 0
        End With
    End If
    Call load开票点分类
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Public Sub ModifyEInvoicePoint(Optional ByVal strID As String)
    '修改电子票据开票点
    'strID-开票点id
    Dim strKey As String
    Dim rsTemp As ADODB.Recordset
    Dim str上级编码 As String, blnRefresh As Boolean
    Dim strTemp As String, bln末级 As Boolean
    Dim frmEdit As New frmEInvoicePointSet
    
    On Error Resume Next
    bln末级 = mint焦点 = Focus_开票点
    
    If Val(strID) = 0 Then
        If mint焦点 = Focus_开票点 Then
            strID = vs开票点.TextMatrix(vs开票点.Row, vs开票点.ColIndex("开票点ID"))
        ElseIf mint焦点 = Focus_开票点分类 Then
            strID = Val(Mid(tvw开票点.SelectedItem.Key, 2))
            If strID = 0 Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    Call frmEdit.Init开票点设置(strID, , bln末级, blnRefresh)
    If blnRefresh Then Call load开票点分类
End Sub

Public Sub StartEInvoicePoint()
    '启用
    On Error GoTo errHandle
    Dim strKey As String, strSQL As String
    Dim intIndex As Long
    Dim strTemp As String
    
    With vs开票点
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("开票点ID"))) = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("停用")) = "" Then Exit Sub
        strTemp = Val(.TextMatrix(.Row, .ColIndex("开票点ID")))
        Me.MousePointer = 11
        strSQL = "zl_电子票据开票点_Start(" & strTemp & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Me.MousePointer = 0
    End With
    
     Call load开票点分类
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Public Sub StopEInvoicePoint()
    '停用
    On Error GoTo errHandle
    Dim strKey As String, strSQL As String
    Dim intIndex As Long
    Dim strTemp As String

    With vs开票点
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("开票点ID"))) = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("停用")) = "1" Then Exit Sub
        strTemp = Val(.TextMatrix(.Row, .ColIndex("开票点ID")))
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "该电子票据开票点下级存在其他开票点，不能停用。", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("你确认要停用名称为“" & .TextMatrix(.Row, .ColIndex("名称")) & "”的电子票据开票点吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            strSQL = "zl_电子票据开票点_Stop(" & strTemp & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Me.MousePointer = 0
        End If
    End With

     Call load开票点分类
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "私有模块\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示停用开票点", mblnShowStop
    SaveSetting "ZLSOFT", "私有模块\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示所有下级开票点", mblnShowAll
     
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
End Sub

Private Sub picFun_Resize()
    On Error Resume Next
    With picFun
        tbPage.Left = 0
        tbPage.Top = 0
        tbPage.Height = .ScaleHeight
        tbPage.Width = .ScaleWidth
    End With
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    With picMain
        picTree.Left = 0
        picTree.Top = 0
        picTree.Height = .ScaleHeight
        picTree.Width = .ScaleWidth * 0.2
        picFun.Left = picTree.Width
        picFun.Top = 0
        picFun.Height = .ScaleHeight
        picFun.Width = .ScaleWidth * 0.8
        picSplit.Left = picTree.Width
        picSplit.Height = .ScaleHeight
    End With
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartX = X
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - sngStartX
        If sngTemp > 1000 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 1000 Then
            picSplit.Left = sngTemp
            picTree.Width = picSplit.Left
            picFun.Left = picSplit.Left + picSplit.Width
            picFun.Width = picMain.ScaleWidth - picFun.Left
        End If
        zlcontrol.ControlSetFocus tvw开票点
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    On Error Resume Next

    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - sngStartY
        If sngTemp - vs开票点.Top > 2500 And Me.ScaleHeight - (sngTemp + picSplitH.Height) > 1500 Then
            picSplitH.Top = sngTemp
            vs开票点.Height = picSplitH.Top
            vs对照.Top = picSplitH.Top + picSplitH.Height
            vs对照.Height = pic电子票据设置.ScaleHeight - vs对照.Top
        End If
        zlcontrol.ControlSetFocus vs开票点
    End If
End Sub

Private Sub picTree_Resize()
    On Error Resume Next
    With picTree
        tvw开票点.Left = 0
        tvw开票点.Top = 0
        tvw开票点.Height = .ScaleHeight
        tvw开票点.Width = .ScaleWidth
    End With
End Sub

Private Sub pic电子票据对码_Resize()
    On Error Resume Next
    With pic电子票据对码
        vs对码明细.Left = 0
        vs对码明细.Top = 0
        vs对码明细.Width = .ScaleWidth
        vs对码明细.Height = .ScaleHeight
    End With
End Sub

Private Sub pic电子票据设置_Resize()
    On Error Resume Next
    With pic电子票据设置
        vs开票点.Left = 0
        vs开票点.Top = 0
        vs开票点.Width = .ScaleWidth
        vs开票点.Height = 0.6 * .ScaleHeight
        vs对照.Left = 0
        vs对照.Top = vs开票点.Height
        vs对照.Width = .ScaleWidth
        vs对照.Height = 0.4 * .ScaleHeight
        picSplitH.Left = 0
        picSplitH.Top = vs开票点.Top + vs开票点.Height
        picSplitH.Width = .ScaleWidth
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tvw开票点.SelectedItem Is Nothing Then Exit Sub
    If Item.Caption = "票据对码明细" Then
        Call Load开票点对码明细(Val(Mid(tvw开票点.SelectedItem.Key, 2)))
        picSplitH.Visible = False
    Else
        Call load开票点(tvw开票点.SelectedItem.Key)
        picSplitH.Visible = True
    End If
End Sub

Private Sub tvw开票点_DblClick()
    If tvw开票点.SelectedItem Is Nothing Then Exit Sub
    Call ModifyEInvoicePoint(Val(Mid(tvw开票点.SelectedItem.Key, 2)))
End Sub

Private Sub tvw开票点_GotFocus()
    mint焦点 = mFocus.Focus_开票点分类
End Sub

Private Sub tvw开票点_LostFocus()
    mint焦点 = mFocus.Focus_None
End Sub

Private Sub tvw开票点_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint焦点 = mFocus.Focus_开票点分类
    Call ShowPopup
End Sub


Public Sub tvw开票点_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errHandle
    If Node Is Nothing Then Exit Sub
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    If tbPage.Selected.Caption = "票据开票点" Then
        Call load开票点(mstrKey)
    Else
        Call Load开票点对码明细(Val(Mid(mstrKey, 2)))
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub load开票点分类()
'功能:加载开票点分类
    Dim strSQL As String
    Dim strKey As String
    Dim rs开票点 As ADODB.Recordset
    Dim i As Integer
    Dim nod As Node
    
    mstrKey = ""
    On Error GoTo errHandle

    If Not tvw开票点.SelectedItem Is Nothing Then
        strKey = tvw开票点.SelectedItem.Key
    End If
            
    strSQL = " Select ID, 上级id, 编码, 名称, 简码, 院区, 客户端, 位置, 末级, 建档时间, 撤档时间 From 电子票据开票点 " & _
                  " Where  Nvl(末级, 0) = 0 " & _
                  " Start with 上级id is null connect by prior id=上级id"
    Set rs开票点 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    tvw开票点.Nodes.Clear
    tvw开票点.Nodes.Add , , "Root", "所有开票点", "Root", "Root"
    tvw开票点.Nodes("Root").Sorted = True
    
    Do Until rs开票点.EOF
            
        If IsNull(rs开票点("上级id")) Then
            tvw开票点.Nodes.Add "Root", tvwChild, "_" & rs开票点("id"), "【" & rs开票点("编码") & "】" & rs开票点("名称"), "Dept", "Dept"
        Else
            tvw开票点.Nodes.Add "_" & rs开票点("上级id"), tvwChild, "_" & rs开票点("id"), "【" & rs开票点("编码") & "】" & rs开票点("名称"), "Dept", "Dept"
        End If
        tvw开票点.Nodes("_" & rs开票点("id")).Sorted = True
        rs开票点.MoveNext
    Loop

    On Error Resume Next
    Set nod = tvw开票点.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvw开票点.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
        tvw开票点_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvw开票点_NodeClick nod
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub load开票点(ByVal str上级ID As String)
'功能:加载开票点
'参数:str上级ID :电子票据开票点.上级id
    Dim strSQL As String
    Dim rs开票点 As ADODB.Recordset
    Dim bln停用 As Boolean
    Dim lng开票点id As Long, i As Integer
    Dim str停用 As String, intRow As Integer
    
    On Error GoTo errHandle
    Call Load开票点对照(0)
    
    If Not vs开票点.Row = 0 Then
        '保留原有键值
        lng开票点id = Val(vs开票点.TextMatrix(vs开票点.Row, 0))
    End If
    
    If Not mblnShowStop Then
        str停用 = " And (A.撤档时间 is null or A.撤档时间 = to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    If mblnShowAll Then
        strSQL = "Select a.*, b.名称 As 部门" & vbNewLine & _
                      "From (Select a.Id, a.上级id, b.名称 As 上级, a.名称, a.编码, a.简码, a.位置, a.客户端, To_Char(a.建档时间, 'YYYY-MM-DD') As 建档时间," & vbNewLine & _
                      "              To_Char(a.撤档时间, 'YYYY-MM-DD') As 撤档时间, a.部门id, a.院区" & vbNewLine & _
                      "       From 电子票据开票点 A, 电子票据开票点 B" & vbNewLine & _
                      "       Where a.上级id = b.Id(+) And Nvl(a.末级, 0) = 1" & str停用 & vbNewLine & _
                      "       Connect By Prior a.Id = a.上级id start with " & IIf(str上级ID = "Root", "A.上级ID is null ", "A.上级ID = [1]") & _
                      "       ) A, 部门表 B" & vbNewLine & _
                      "Where a.部门id = b.Id(+)"
    Else
         strSQL = "Select a.Id, a.上级id, c.名称 As 上级, a.名称, a.编码, a.简码, a.位置, a.客户端, To_Char(a.建档时间, 'YYYY-MM-DD') As 建档时间," & vbNewLine & _
                      "              To_Char(a.撤档时间, 'YYYY-MM-DD') As 撤档时间, a.部门id, a.院区,b.名称 As 部门" & vbNewLine & _
                      "       From 电子票据开票点 A, 部门表 B,电子票据开票点 C" & vbNewLine & _
                      "       Where Nvl(a.末级, 0) = 1 And a.部门id = b.Id(+) And a.上级id = c.Id(+) " & str停用 & IIf(str上级ID = "Root", " And A.上级ID is null ", " And A.上级ID = [1]")
    End If
    Set rs开票点 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(str上级ID, 2)))
    vs开票点.Clear 1: vs开票点.Rows = 2
    If rs开票点.EOF Then Exit Sub
    With vs开票点
        .Rows = rs开票点.RecordCount + 1
        For i = 1 To rs开票点.RecordCount
            .TextMatrix(i, .ColIndex("开票点ID")) = rs开票点!ID
            .TextMatrix(i, .ColIndex("名称")) = Nvl(rs开票点!名称)
            .TextMatrix(i, .ColIndex("编码")) = Nvl(rs开票点!编码)
            .TextMatrix(i, .ColIndex("简码")) = Nvl(rs开票点!简码)
            .TextMatrix(i, .ColIndex("客户端")) = Nvl(rs开票点!客户端)
            .TextMatrix(i, .ColIndex("位置")) = Nvl(rs开票点!位置)
            .TextMatrix(i, .ColIndex("建档时间")) = Nvl(rs开票点!建档时间)
            .TextMatrix(i, .ColIndex("撤档时间")) = Nvl(rs开票点!撤档时间)
            .TextMatrix(i, .ColIndex("上级")) = Nvl(rs开票点!上级)
            .TextMatrix(i, .ColIndex("部门")) = Nvl(rs开票点!部门)
            .TextMatrix(i, .ColIndex("院区")) = Nvl(rs开票点!院区)
             If Not CDate(IIf(IsNull(rs开票点("撤档时间")), CDate("3000/1/1"), rs开票点("撤档时间"))) = CDate("3000/1/1") Then
                .Cell(flexcpForeColor, i, .ColIndex("名称"), i, .ColIndex("院区")) = RGB(255, 0, 0)
                .TextMatrix(i, .ColIndex("停用")) = "1"
            End If
            rs开票点.MoveNext
        Next
        intRow = .FindRow(lng开票点id, 0, .ColIndex("开票点ID"), , True)
        If intRow > 0 Then .Row = intRow
        If Val(.TextMatrix(1, .ColIndex("开票点ID"))) > 0 Then
            Call Load开票点对照(Val(.TextMatrix(.Row, .ColIndex("开票点ID"))))
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckExistDepPres(ByVal lng上级id As Long) As Boolean
    '检查该电子票据开票点下是否存在其他开票点
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select 1 From 电子票据开票点 " & _
        " Where 上级id =[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "电子票据开票点", lng上级id)
    
    If rsTemp.RecordCount > 0 Then
        CheckExistDepPres = True
        Exit Function
    End If
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Public Sub Load开票点对照(ByVal lng开票点id As Long)
    '根据开票点id加载开票点对照信息
    Dim strSQL As String, i As Integer
    Dim rs开票点对照 As New ADODB.Recordset
    
    With vs对照
        .ColHidden(.ColIndex("收费员")) = mint对码方式 = 0
        .ColHidden(.ColIndex("收费员编号")) = mint对码方式 = 0
        .ColHidden(.ColIndex("收费员所属部门")) = mint对码方式 = 0
        .ColHidden(.ColIndex("客户端")) = mint对码方式 = 1
        .ColHidden(.ColIndex("部门")) = mint对码方式 = 1
        .ColHidden(.ColIndex("用途")) = mint对码方式 = 1
    End With
    vs对照.Clear 1: vs对照.Rows = 2
    If lng开票点id = 0 Then Exit Sub
    strSQL = "Select a.Id As 开票点id, b.id,a.名称 As 开票点, b.人员id, c.姓名 As 收费员, c.编号 as 收费员编号,g.名称 As 收费员所属部门, b.客户端, e.部门, e.用途" & vbNewLine & _
                    "From 电子票据开票点 A, 票据开票点对照 B, 人员表 C, 部门人员 D, zlClients E, 部门表 G" & vbNewLine & _
                    "Where a.Id = b.开票点id(+) And b.人员id = c.Id(+) And b.客户端 = e.工作站(+) And c.Id = d.人员id(+) And d.缺省(+) = 1 And" & vbNewLine & _
                    "      d.部门id = g.Id(+) And a.Id = [1] "
    Set rs开票点对照 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng开票点id)
    If rs开票点对照.EOF Then Exit Sub
    With vs对照
        .Rows = rs开票点对照.RecordCount + 1
        For i = 1 To rs开票点对照.RecordCount
            .TextMatrix(i, .ColIndex("开票点ID")) = rs开票点对照!开票点id
            .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(rs开票点对照!ID))
            .TextMatrix(i, .ColIndex("开票点")) = Nvl(rs开票点对照!开票点)
            .TextMatrix(i, .ColIndex("人员id")) = Nvl(rs开票点对照!人员id)
            .TextMatrix(i, .ColIndex("收费员")) = Nvl(rs开票点对照!收费员)
            .TextMatrix(i, .ColIndex("收费员编号")) = Nvl(rs开票点对照!收费员编号)
            .TextMatrix(i, .ColIndex("收费员所属部门")) = Nvl(rs开票点对照!收费员所属部门)
            .TextMatrix(i, .ColIndex("客户端")) = Nvl(rs开票点对照!客户端)
            .TextMatrix(i, .ColIndex("部门")) = Nvl(rs开票点对照!部门)
            .TextMatrix(i, .ColIndex("用途")) = Nvl(rs开票点对照!用途)
            rs开票点对照.MoveNext
        Next
    End With
End Sub

Public Sub Load开票点对码明细(ByVal lng上级id As Long)
    '根据开票点id加载开票点对照信息
    Dim strSQL As String, i As Integer
    Dim rs开票点对照 As New ADODB.Recordset
    
    vs对码明细.Clear 1: vs对码明细.Rows = 2
    With vs对码明细
        .ColHidden(.ColIndex("收费员")) = mint对码方式 = 0
        .ColHidden(.ColIndex("收费员编号")) = mint对码方式 = 0
        .ColHidden(.ColIndex("收费员所属部门")) = mint对码方式 = 0
        .ColHidden(.ColIndex("客户端")) = mint对码方式 = 1
        .ColHidden(.ColIndex("部门")) = mint对码方式 = 1
        .ColHidden(.ColIndex("用途")) = mint对码方式 = 1
    End With
    If mblnShowAll Then
        strSQL = "Select  a.开票点id, a.开票点, f.Id, f.人员id, f.客户端, b.姓名 As 收费员, b.编号 As 收费员编号, e.名称 As 收费员所属部门, d.部门, d.用途 " & vbNewLine & _
                        "From(Select a.Id As 开票点id, a.名称 As 开票点 From 电子票据开票点 A Where a.末级 = 1 Connect By Prior a.Id = a.上级id" & vbNewLine & _
                        "Start With " & IIf(Val(lng上级id) = 0, "A.上级ID is null )", "A.上级ID = [1])") & "A, 人员表 B, 部门人员 C, zlClients D, 部门表 E, 票据开票点对照 F " & vbNewLine & _
                        "Where f.人员id = b.Id(+) And f.客户端 = d.工作站(+) And f.Id = c.人员id(+) And c.缺省(+) = 1  And " & vbNewLine & _
                        "      c.部门id = e.Id(+) And a.开票点id = f.开票点id(+)"
    Else
        strSQL = "Select a.开票点id, a.开票点, f.Id, f.人员id, f.客户端, b.姓名 As 收费员, b.编号 As 收费员编号, e.名称 As 收费员所属部门, d.部门, d.用途 " & vbNewLine & _
                        "From(Select a.Id As 开票点id, a.名称 As 开票点 From 电子票据开票点 A Where  a.末级 = 1 " & vbNewLine & _
                        "And " & IIf(Val(lng上级id) = 0, "A.上级ID is null )", "A.上级ID = [1])") & "A, 人员表 B, 部门人员 C, zlClients D, 部门表 E, 票据开票点对照 F" & vbNewLine & _
                        "Where f.人员id = b.Id(+) And f.客户端 = d.工作站(+) And f.Id = c.人员id(+) And c.缺省(+) = 1  And  " & vbNewLine & _
                        "      c.部门id = e.Id(+) And a.开票点id = f.开票点id(+)"
    End If
    Set rs开票点对照 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng上级id)
    If rs开票点对照.EOF Then Exit Sub
    With vs对码明细
        .Rows = rs开票点对照.RecordCount + 1
        For i = 1 To rs开票点对照.RecordCount
            .TextMatrix(i, .ColIndex("开票点ID")) = rs开票点对照!开票点id
            .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(rs开票点对照!ID))
            .TextMatrix(i, .ColIndex("开票点")) = Nvl(rs开票点对照!开票点)
            .TextMatrix(i, .ColIndex("人员id")) = Nvl(rs开票点对照!人员id)
            .TextMatrix(i, .ColIndex("收费员")) = Nvl(rs开票点对照!收费员)
            .TextMatrix(i, .ColIndex("收费员编号")) = Nvl(rs开票点对照!收费员编号)
            .TextMatrix(i, .ColIndex("收费员所属部门")) = Nvl(rs开票点对照!收费员所属部门)
            .TextMatrix(i, .ColIndex("客户端")) = Nvl(rs开票点对照!客户端)
            .TextMatrix(i, .ColIndex("部门")) = Nvl(rs开票点对照!部门)
            .TextMatrix(i, .ColIndex("用途")) = Nvl(rs开票点对照!用途)
            rs开票点对照.MoveNext
        Next
    End With
End Sub

Private Sub Set开票点对照(Optional ByVal blnModify As Boolean)
    '根据开票点id设置开票点对照信息
    Dim lng开票点id As Long, lngID As Long
    Dim frmEdit As New frmEInvoicePointSet
    Dim bln对照 As Boolean, blnRefresh As Boolean
    
    If mint焦点 <> Focus_开票点对照 And mint焦点 <> Focus_对照明细 Then Exit Sub
    bln对照 = mint焦点 = Focus_开票点对照
    If bln对照 Then
        lng开票点id = Val(vs对照.TextMatrix(vs对照.Row, vs对照.ColIndex("开票点id")))
        lngID = Val(vs对照.TextMatrix(vs对照.Row, vs对照.ColIndex("id")))
    Else
        lng开票点id = Val(vs对码明细.TextMatrix(vs对码明细.Row, vs对码明细.ColIndex("开票点id")))
        lngID = Val(vs对码明细.TextMatrix(vs对码明细.Row, vs对码明细.ColIndex("ID")))
    End If
    If lng开票点id = 0 Then Exit Sub
    If Not blnModify Then
        lngID = 0
    Else
        If lngID = 0 Then Exit Sub
    End If
    Call frmEdit.Init开票点对码(mint对码方式, lng开票点id, lngID, blnRefresh)
    If Not blnRefresh Then Exit Sub
    If bln对照 Then
        Call Load开票点对照(lng开票点id)
    Else
        Call Load开票点对码明细(Val(Mid(mstrKey, 2)))
    End If
End Sub

Private Sub Delete开票点对照()
    '删除开票点对照信息
    Dim lngID As Long, lng开票点id As Long
    Dim strSQL As String, bln对照 As Boolean
    
    If mint焦点 <> Focus_开票点对照 And mint焦点 <> Focus_对照明细 Then Exit Sub
    bln对照 = mint焦点 = Focus_开票点对照
    If bln对照 Then
        lngID = Val(vs对照.TextMatrix(vs对照.Row, vs对照.ColIndex("id")))
        lng开票点id = Val(vs对照.TextMatrix(vs对照.Row, vs对照.ColIndex("开票点id")))
    Else
        lngID = Val(vs对码明细.TextMatrix(vs对码明细.Row, vs对码明细.ColIndex("id")))
        lng开票点id = Val(vs对码明细.TextMatrix(vs对码明细.Row, vs对码明细.ColIndex("开票点id")))
    End If
    If lng开票点id = 0 Then Exit Sub
    If lngID = 0 Then Exit Sub
    strSQL = "Zl_票据开票点对照_Update(2," & lngID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "票据开票点对照")
    If bln对照 Then
        Call Load开票点对照(lng开票点id)
    Else
        Call Load开票点对码明细(Val(Mid(mstrKey, 2)))
    End If
End Sub

Private Sub vs对码明细_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs对码明细, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs对码明细_DblClick()
    '对码
    Dim blnModify As Boolean
    If vs对码明细.Row = 0 Then Exit Sub
    blnModify = Val(vs对码明细.TextMatrix(vs对码明细.Row, vs对码明细.ColIndex("id"))) > 0
    Call Set开票点对照(blnModify)
End Sub

Private Sub vs对码明细_GotFocus()
    mint焦点 = Focus_对照明细
    If vs对码明细.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs对码明细, &HFFEBD7
End Sub

Private Sub vs对码明细_LostFocus()
    mint焦点 = Focus_None
    If vs对码明细.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs对码明细
    OS.OpenIme False
End Sub

Private Sub vs对码明细_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint焦点 = mFocus.Focus_对照明细
    Call ShowPopup
End Sub

Private Sub vs对照_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs对照, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs对照_DblClick()
    '对码
    Dim blnModify As Boolean
    If vs对照.Row = 0 Then Exit Sub
    blnModify = Val(vs对照.TextMatrix(1, vs对照.ColIndex("id"))) > 0
    Call Set开票点对照(blnModify)
End Sub

Private Sub vs对照_GotFocus()
    mint焦点 = mFocus.Focus_开票点对照
    If vs对照.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs对照, &HFFEBD7
End Sub

Private Sub vs对照_LostFocus()
    mint焦点 = mFocus.Focus_None
    If vs对照.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs对照
    OS.OpenIme False
End Sub

Private Sub vs对照_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint焦点 = mFocus.Focus_开票点对照
    Call ShowPopup
End Sub

Private Sub ShowPopup()
    '显示弹出菜单
    Dim objPopup As CommandBarPopup
    Err = 0: On Error GoTo ErrHandler
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    Dim objItem As TabControlItem
    With tbPage
        Set objItem = .InsertItem(1, "票据开票点", pic电子票据设置.hWnd, 0)
        objItem.Tag = 1
        Set objItem = .InsertItem(2, "票据对码明细", pic电子票据对码.hWnd, 0)
        objItem.Tag = 2
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    tbPage.Item(0).Selected = True
End Sub

Private Sub vs开票点_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Not (NewRow = 0 Or OldRow = 0) Then zl_VsGridRowChange vs开票点, OldRow, NewRow, OldCol, NewCol
    With vs开票点
        If NewRow = 0 Then Exit Sub
        If Val(.TextMatrix(NewRow, .ColIndex("开票点id"))) = 0 Then Exit Sub
        If .TextMatrix(NewRow, .ColIndex("停用")) = "1" Then
            .Cell(flexcpForeColor, NewRow, .ColIndex("名称"), NewRow, .ColIndex("院区")) = RGB(255, 0, 0)
        Else
            .Cell(flexcpForeColor, NewRow, .ColIndex("名称"), NewRow, .ColIndex("院区")) = &H80000008
        End If
        Call Load开票点对照(Val(.TextMatrix(NewRow, .ColIndex("开票点id"))))
    End With
End Sub

Private Sub vs开票点_DblClick()
    With vs开票点
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("开票点id"))) = 0 Then Exit Sub
        Call ModifyEInvoicePoint(Val(.TextMatrix(.Row, .ColIndex("开票点id"))))
    End With
End Sub

Private Sub vs开票点_GotFocus()
    mint焦点 = mFocus.Focus_开票点
    If vs开票点.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs开票点, &HFFEBD7
End Sub

Private Sub vs开票点_LostFocus()
    mint焦点 = mFocus.Focus_None
    If vs开票点.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs开票点
    OS.OpenIme False
End Sub

Private Sub vs开票点_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint焦点 = mFocus.Focus_开票点
    Call ShowPopup
End Sub


