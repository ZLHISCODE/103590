VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAuditItemShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病案审查项目"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   Icon            =   "frmAuditItemShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9900
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   3690
      ScaleHeight     =   450
      ScaleWidth      =   5010
      TabIndex        =   13
      Top             =   4980
      Width           =   5010
      Begin VB.CommandButton cmdSearch 
         Height          =   360
         Left            =   2610
         Picture         =   "frmAuditItemShow.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   375
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   930
         TabIndex        =   17
         Top             =   60
         Width           =   1650
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "所有(&A)"
         Height          =   360
         Left            =   0
         TabIndex        =   16
         Top             =   30
         Width           =   885
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   4995
         TabIndex        =   15
         Top             =   30
         Width           =   885
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   360
         Left            =   4050
         TabIndex        =   14
         Top             =   30
         Width           =   885
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3075
         TabIndex        =   19
         Top             =   135
         Width           =   90
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   0
      Left            =   4260
      ScaleHeight     =   2715
      ScaleWidth      =   5010
      TabIndex        =   11
      Top             =   1875
      Width           =   5010
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditItem 
         Height          =   4695
         Left            =   105
         TabIndex        =   12
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   5520
      Index           =   2
      Left            =   90
      ScaleHeight     =   5520
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   300
      Width           =   3015
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
         TabIndex        =   4
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
            TabIndex        =   5
            Top             =   75
            Width           =   255
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
            TabIndex        =   10
            Top             =   90
            Width           =   1095
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
            TabIndex        =   9
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label lbl启用时间 
            BackStyle       =   0  'Transparent
            Caption         =   "启用时间:"
            Height          =   195
            Left            =   225
            TabIndex        =   8
            Top             =   1365
            Width           =   2580
         End
         Begin VB.Label lbl总分 
            BackStyle       =   0  'Transparent
            Caption         =   "总分:"
            Height          =   195
            Left            =   225
            TabIndex        =   7
            Top             =   705
            Width           =   2580
         End
         Begin VB.Label lbl分段线 
            BackStyle       =   0  'Transparent
            Caption         =   "分段线:"
            Height          =   195
            Left            =   225
            TabIndex        =   6
            Top             =   1035
            Width           =   2580
         End
      End
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
         TabIndex        =   1
         Top             =   240
         Width           =   2940
         Begin MSComctlLib.TreeView tvwAuditType 
            Height          =   1200
            Left            =   495
            TabIndex        =   2
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
            TabIndex        =   3
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   990
      Top             =   60
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
            Picture         =   "frmAuditItemShow.frx":D0A4
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditItemShow.frx":DEF6
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   7125
      Picture         =   "frmAuditItemShow.frx":E16A
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   4245
      Picture         =   "frmAuditItemShow.frx":E1B9
      Top             =   7410
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   495
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAuditItemShow.frx":E20E
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   4635
      Picture         =   "frmAuditItemShow.frx":E222
      Top             =   7380
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   1770
      Picture         =   "frmAuditItemShow.frx":E3E0
      Top             =   7365
      Visible         =   0   'False
      Width           =   2790
   End
End
Attribute VB_Name = "frmAuditItemShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public mlngObject As Long                               '浏览对象
Public mstrSummitID As String                              '提交ID
Public mRetMain As ADODB.Recordset                                  '传入记录集

Public mlngOk As Long
Private mstrSaveKey             As String               '保存的上次的分类选择关键字
Private mintTypeID              As Integer              '分类新增、修改、删除时的ID
Private mintItemID              As Integer              '项目新增、修改、删除时的ID
Private mRsAuditItem            As ADODB.Recordset      '数据集
Private mlngCurFAID             As Long                 '当前方案ID
Private mblnProgUsed            As Boolean              '方案是否已使用
Private mblnCheckAll            As Boolean              '是否显示下级
Private zlCheck                 As New clsCheck         '检测类
Public mstr分制  As String
Public mstr分值  As String
Public mstr名称 As String
Public mstrID As String
Private mlng分类ID As Long

Private Const con_vsfField = "/*+ rule */ '' as 图标,a.id, a.分类id,a.编码,a.名称,a.简码,a.分值,a.分制,b.名称 as 分类,decode(a.适用对象,1,'住院医嘱',2,'住院病历',3,'护理病历',4,'护理记录',5,'首页记录',6,'医嘱报告',7,'疾病证明',8,'知情文件','未定义') as 适用对象,a.说明,a.审查依据,适用对象 as 适用编码,文件ID,适用环节"
Private Const conFieldFiles = "Select /*+ rule */ a.id as 文件ID,a.编号 as 文件编码,a.名称 as 文件名称,a.说明 as 文件说明" & vbCrLf & _
                         "from 病历文件列表 A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B " & vbCrLf & _
                         "where /*+ rule */a.id = b.COLUMN_VALUE And a.种类 = [2]"


'树节点定位
Dim nod                         As Node
Dim i                           As Long
Dim FirstKey                    As String
Dim v                           As Variant

Public Function zlInitData(ByVal RetMain As ADODB.Recordset, ByVal lngObject As Long, ByVal strSummitID As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mRetMain = RetMain
    mlngObject = lngObject
    mstrSummitID = strSummitID
End Function

Private Sub cmdAll_Click()
    Call DataFill
End Sub

Private Sub cmdCancel_Click()
    mlngOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mlngOk = True
    With vsfAuditItem
        
        If .Row > 0 Then
            mstr名称 = .TextMatrix(.Row, .ColIndex("名称"))
            mstrID = .RowData(.Row)
            mstr分制 = .TextMatrix(.Row, .ColIndex("分制"))
            mstr分值 = .TextMatrix(.Row, .ColIndex("分值"))
        End If
    End With
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text = "" Then Exit Sub
    Call GetAuditItem(mlngObject, mstrSummitID, txtSearch.Text)
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

Private Sub Form_Load()
    On Error GoTo ErrH
    mblnCheckAll = True
    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("加载数据")
    picFAXX.Picture = imgClose.Picture
    Call RestoreWinState(Me, App.ProductName)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrH
    Call SetPaneRange(dkpMain, 1, 200, 100, 200, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 2, 150, 100, 150, Me.ScaleWidth)
    Call SetPaneRange(dkpMain, 3, 350, 30, 350, 30)
    
    dkpMain.RecalcLayout
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
    Case "加载数据"
        Call DataFill
    End Select
    ExecuteCommand = True
    
    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 控件初始化
'==============================================================================
Private Sub InitControl()
    
    On Error GoTo ErrH
    
    Call InitVsflexGrid
    Call InitCommandBar
    Call InitDockPannel
    Call InitTreeView
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
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
    
    Set objPane = dkpMain.CreatePane(2, 200, 100, DockRightOf, Nothing)
    objPane.Title = "详细"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 300, 30, DockBottomOf, Nothing)
    objPane.Title = "控制"
    objPane.Options = PaneNoCaption

    dkpMain.SetCommandBars cbsMain
    
    Call DockPannelInit(dkpMain)
    
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
        .ColWidth(.ColIndex("名称")) = 3000
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
    
    gstrSQL = "Select ID,名称,启用时间 From 病案审查方案 where 启用时间 is not null"
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

    gstrSQL = "SELECT /*+ rule */ id,上级ID,方案ID,编码,名称 FROM 病案审查分类 where 方案ID=" & intStartid & " START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID "
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
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
'=功能： 初始菜单工具栏
'==============================================================================
Private Sub InitCommandBar()

    
    On Error GoTo ErrH

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Visible = False
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
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
'=功能： 网格数据加载 vsfAuditItem
'==============================================================================
Private Sub DataAuditItem(Optional strWhere As String)
    Dim strKey      As String
    Dim i           As Long
    Dim nTmpNode As Node
    Dim strWhere1 As String
    
    On Error GoTo ErrH
    If strWhere = "" Then
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
        
        If mstrSummitID = "" Then
            strWhere1 = " And A.文件ID is Null"
        Else
            strWhere1 = " And (A.文件ID is null or instr(','|| A.文件ID || ',' , ','|| '" & mstrSummitID & "' || ',')>0 )"
        End If
        
        If mblnCheckAll Then
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From 病案审查目录 a,(SELECT /*+ rule */ id,名称 FROM 病案审查分类 START WITH id=[1] CONNECT BY PRIOR ID = 上级ID)b " & vbCrLf & _
                    "Where a.分类id = b.ID and a.适用对象=[2]" & strWhere1
        Else
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From 病案审查目录 a,病案审查分类 b" & vbCrLf & _
                    "Where a.分类id = b.ID and a.分类id=[1] and a.适用对象=[2]" & strWhere1
        End If
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, strKey, mlngObject)
    Else
        gstrSQL = "" & _
                "Select " & con_vsfField & vbCrLf & _
                "From 病案审查目录 a,病案审查分类 b" & vbCrLf & _
                "Where a.分类id = b.ID and a.适用对象=[1] And" & vbCrLf & strWhere
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlngObject)
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
    lblInfo.Caption = "总共:" & mRsAuditItem.RecordCount & "条审查记录。"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataFill()
    Dim i           As Long
    
    On Error GoTo ErrH
    Set vsfAuditItem.DataSource = mRetMain
       
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
    lblInfo.Caption = "总共:" & mRetMain.RecordCount & "条审查记录。"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsfAuditItem.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
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
    Case 1
        cmdCancel.Move picPane(Index).Width - cmdCancel.Width - 60
        cmdOk.Move cmdCancel.Left - cmdOk.Width - 30
    End Select
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = 13 Then
       If txtSearch.Text = "" Then Exit Sub
        Call GetAuditItem(mlngObject, mstrSummitID, txtSearch.Text)
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub vsfAuditItem_DblClick()
    Call cmdOk_Click
End Sub

'==============================================================================
'=功能：行列变换时
'==============================================================================
Private Sub vsfAuditItem_RowColChange()
    Dim rsTemp          As ADODB.Recordset
    Dim varPos          As Variant
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
'        stbThis.Panels(2) = "当前显示有 0 个项目。"
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
        gstrSQL = conFieldFiles
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("文件ID")), AuditFileTran(zlCheck.Cmb_ID(.cboUsed), 0))
        Set .vsfFiles.DataSource = rsTemp
'        '读取 病历文件内容
'        If .txtFileID.Tag <> "" Then
'            gstrSQL = "select 名称 from 病历文件列表 where ID = [1] "
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, Val(.txtFileID.Tag))
'            If Not zlCheck.Connection_ChkRsState(rsTemp) Then
'                .txtFileID.Text = "" & rsTemp.Fields!名称
'            Else
'                .txtFileID.Text = ""
'            End If
'        Else
'            .txtFileID.Text = ""
'        End If
    End With
    mlng分类ID = vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("分类ID"))
    
'    stbThis.Panels(2) = "当前显示有 " & vsfAuditItem.Rows - 1 & " 个项目。"
    varPos = zlCheck.Connection_GetBookMark(mRsAuditItem, "ID=" & CStr("" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))))
    DoEvents
    If Not IsNull(varPos) Then
        If Val(varPos) > 0 Then mRsAuditItem.Bookmark = varPos
    End If
    
    Call TreeviewSelect(mlng分类ID, tvwAuditType)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Private Sub GetAuditItem(intObject As Long, strFileID As String, Optional shortName As String = "")
Dim rsData As ADODB.Recordset, strSubid As String, strReturn As String
Dim i As Long
On Error GoTo ErrH
    If IsNumeric(strFileID) Then
        '检测文件ID存在与否
        '电子病案记录 如果存在则直接取文件ID关联，否则直接按类型读取
        gstrSQL = "Select 文件ID From 电子病历记录 a Where a.ID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFileID)
        If zlCheck.Connection_ChkRsState(rsData) Then
            strFileID = 0
        Else
            strFileID = "" & rsData.Fields!文件ID
        End If
    Else
        If Not gobjEmr Is Nothing Then
            If InStr(strFileID, "|") > 0 Then
                strSubid = Split(strFileID, "|")(1)
                strFileID = Split(strFileID, "|")(0)
            End If
            gstrSQL = "Select RawtoHex(Antetype_Id) 文件ID From Bz_Doc_Tasks A Where Real_Doc_Id = Hextoraw(:docid)" & IIf(strSubid = "", "", " And Subdoc_Id =:subdocid")
            strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strFileID & "^" & DbType.T_String & "^docid" & IIf(strSubid = "", "", "|" & strSubid & "^" & DbType.T_String & "^subdocid"), rsData)
            If strReturn <> "" Then strFileID = 0
            If Not rsData Is Nothing Then
            If rsData.RecordCount > 0 Then
                strFileID = rsData!文件ID
            End If
            End If
        End If
    End If
    If strFileID = "0" Then
        gstrSQL = "Select /*+ rule */ '' as 图标,A.ID, A.分类ID,A.编码, A.名称, A.简码, A.分值,A.分制,b.名称 as 分类,decode(a.适用对象,1,'住院医嘱',2,'住院病历',3,'护理病历',4,'护理记录',5,'首页记录',6,'医嘱报告',7,'疾病证明',8,'知情文件','未定义') as 适用对象,a.说明,a.审查依据,适用对象 as 适用编码,文件ID,适用环节 From 病案审查目录  A ,病案审查分类 B,病案审查方案 C where  A.分类ID =B.id And B.方案ID =C.ID And C.启用时间 is Not Null And A.适用对象 = " & CStr(intObject)
    Else
        gstrSQL = "Select /*+ rule */ '' as 图标,A.ID, A.分类ID,A.编码, A.名称, A.简码, A.分值,A.分制,b.名称 as 分类,decode(a.适用对象,1,'住院医嘱',2,'住院病历',3,'护理病历',4,'护理记录',5,'首页记录',6,'医嘱报告',7,'疾病证明',8,'知情文件','未定义') as 适用对象,a.说明,a.审查依据,适用对象 as 适用编码,文件ID,适用环节 From 病案审查目录 A ,病案审查分类 B,病案审查方案 C  where A.分类ID =B.id And B.方案ID =C.ID And C.启用时间 is Not Null And A.适用对象 = " & CStr(intObject) & " And (A.文件ID is null or instr(','|| A.文件ID || ',' , ','|| '" & strFileID & "' || ',')>0 )"
    End If
    If shortName <> "" Then
        shortName = UCase(shortName)
        gstrSQL = gstrSQL & vbCrLf & "And (A.编码 like '%" & shortName & "%' or A.简码 like '%" & shortName & "%' or A.名称 like '%" & shortName & "%')"
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsData.RecordCount > 0 Then
        
        Set vsfAuditItem.DataSource = rsData
       
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
        lblInfo.Caption = "总共:" & rsData.RecordCount & "条审查记录。"
      
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
End Sub

Private Function TreeviewSelect(ByVal lng分类ID As Long, tvwMain As TreeView)
    '设置选择的项
    Dim i As Long
    Dim NodeMain As MSComctlLib.Node
    Dim nodeChild As MSComctlLib.Node
    Dim str分类ID  As String
    str分类ID = CStr(lng分类ID)
    
    On Error Resume Next
    If lng分类ID = 0 Then Exit Function
    If ObjPtr(tvwAuditType) > 0 Then
        Set NodeMain = tvwMain.Nodes.Item(1)
        If NodeMain.Key = "A" & str分类ID Then
            NodeMain.Selected = True
            Exit Function
        End If
        
        
        If NodeMain.Children > 0 Then
            Set nodeChild = NodeMain.Child
            For i = 1 To NodeMain.Children
                If nodeChild.Key = "A" & str分类ID Then
                    nodeChild.Selected = True
                    Exit For
                End If
                
                If Not nodeChild.Child Is Nothing Then
                   If SetSelectTvwChild("A" & str分类ID, nodeChild) Then
                        Exit For
                   End If
                End If
                Set nodeChild = nodeChild.Next
                
            Next
        End If
    End If
End Function

Public Function SetSelectTvwChild(ByVal strTvwMain As String, tvwNode As Node) As Boolean
    Dim nodeChild As MSComctlLib.Node
    Dim i As Long
    On Error Resume Next
    
    
    If tvwNode.Key = strTvwMain Then
        tvwNode.Selected = True
        SetSelectTvwChild = True
        Exit Function
    End If
    
    If tvwNode.Children > 0 Then
        Set nodeChild = tvwNode.Child
        For i = 1 To tvwNode.Children
    
            If nodeChild.Key = strTvwMain Then
                nodeChild.Selected = True
                SetSelectTvwChild = True
                Exit Function
            End If
            
        
            If Not nodeChild.Child Is Nothing Then
               If SetSelectTvwChild(strTvwMain, nodeChild) Then
                    SetSelectTvwChild = True
                    Exit Function
               End If
            End If
            Set nodeChild = nodeChild.Next
            
        Next
    End If
    
End Function
