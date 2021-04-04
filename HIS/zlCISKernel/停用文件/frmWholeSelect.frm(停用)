VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmWholeSelect 
   Caption         =   "收费成套项目选择"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10605
   Icon            =   "frmWholeSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10605
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10605
      TabIndex        =   9
      Top             =   6195
      Width           =   10605
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Enabled         =   0   'False
         Height          =   380
         Left            =   7275
         TabIndex        =   3
         ToolTipText     =   "快键:F2"
         Top             =   165
         Width           =   1250
      End
      Begin VB.CheckBox chkSub 
         Caption         =   "显示所有下级项目(&S)"
         Height          =   210
         Left            =   195
         TabIndex        =   10
         Top             =   270
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   380
         Left            =   8880
         TabIndex        =   4
         Top             =   180
         Width           =   1250
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   12000
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   0
         X2              =   12000
         Y1              =   90
         Y2              =   90
      End
   End
   Begin VB.PictureBox picTree 
      BorderStyle     =   0  'None
      Height          =   2640
      Left            =   255
      ScaleHeight     =   2640
      ScaleWidth      =   3015
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3150
      Width           =   3015
      Begin MSComctlLib.TreeView TvwWholeSet 
         Height          =   2490
         Left            =   30
         TabIndex        =   8
         ToolTipText     =   "快速定位快键:F4"
         Top             =   270
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   4392
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "img16"
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
   Begin VB.PictureBox picWholeSubItems 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   3720
      ScaleHeight     =   2655
      ScaleWidth      =   6450
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2745
      Width           =   6450
      Begin VB.CommandButton cmdCls 
         Caption         =   "全清"
         Height          =   330
         Left            =   900
         TabIndex        =   14
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdALL 
         Caption         =   "全选"
         Height          =   330
         Left            =   75
         TabIndex        =   13
         Top             =   30
         Width           =   795
      End
      Begin VB.CheckBox chk缺省显示 
         Caption         =   "缺省选择所有项目"
         Height          =   210
         Left            =   4950
         TabIndex        =   12
         Top             =   105
         Width           =   1845
      End
      Begin VSFlex8Ctl.VSFlexGrid vsWholeSet 
         Height          =   4680
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "快速定位:F6"
         Top             =   405
         Width           =   11355
         _cx             =   20029
         _cy             =   8255
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         Rows            =   4
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmWholeSelect.frx":6852
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
         Editable        =   2
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
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   1665
      Left            =   3690
      ScaleHeight     =   1665
      ScaleWidth      =   5370
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   870
      Width           =   5370
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   780
         TabIndex        =   11
         ToolTipText     =   "快速定位快键:F3"
         Top             =   30
         Width           =   2595
      End
      Begin MSComctlLib.ListView lvwWholeSetItem 
         Height          =   1455
         Left            =   15
         TabIndex        =   1
         Top             =   390
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lbl查找 
         AutoSize        =   -1  'True
         Caption         =   "查找(&F)"
         Height          =   180
         Left            =   135
         TabIndex        =   0
         Top             =   90
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1125
      Top             =   2370
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
            Picture         =   "frmWholeSelect.frx":6ACC
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":7066
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":7600
            Key             =   "成药"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":7B9A
            Key             =   "诊疗"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":8134
            Key             =   "草药"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":86CE
            Key             =   "方案"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   750
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":8C68
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":90C0
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":9514
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":A366
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":B1B8
            Key             =   "RootS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":B312
            Key             =   "Exp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":B46C
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":B8BE
            Key             =   "RootR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":BD10
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":C168
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":C5BC
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":CA10
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":CE64
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWholeSelect.frx":DCB6
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   420
      Top             =   1035
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmWholeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String, mintColumn1 As Integer
Private mblnOk As Boolean, mblnFirst As Boolean, mblnNotClick As Boolean
Private mblnItem As Boolean
Private mrsOutSel As ADODB.Recordset
Private Const mstrLvwWholeSet As String = "名称,1500,0,1;编码,800,0,2;拼音,1400,0,0;五笔,1400,0,0;使用范围,1000,0,0;所属分类,2400,0,0"
Private Enum mPanceIdx
    pan_Tree = 1
    pan_WholeSet = 2
    pan_WholeItems = 3
    pan_Cmd = 4
End Enum
Private mrs成套项目 As ADODB.Recordset
 
Private Sub cmdALL_Click()
    Dim i As Long
    With vsWholeSet
        .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = -1
        .Cell(flexcpForeColor, 1, 1, .Rows - 1, .Cols - 1) = vbBlue
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("收费项目"))) <> 0 Then
                .TextMatrix(i, .ColIndex("缺省数量")) = IIf(Val(.TextMatrix(i, .ColIndex("缺省数量"))) = 0, .Cell(flexcpData, i, .ColIndex("缺省数量")), .TextMatrix(i, .ColIndex("缺省数量")))
                .TextMatrix(i, .ColIndex("缺省付数")) = IIf(Val(.TextMatrix(i, .ColIndex("缺省付数"))) = 0, .Cell(flexcpData, i, .ColIndex("缺省付数")), .TextMatrix(i, .ColIndex("缺省付数")))
            End If
        Next
    End With
End Sub

Private Sub cmdCls_Click()
    Dim i As Long
    With vsWholeSet
        .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
        .Cell(flexcpForeColor, 1, 1, .Rows - 1, .Cols - 1) = .ForeColor
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("收费项目"))) <> 0 Then
                .TextMatrix(i, .ColIndex("缺省数量")) = ""
                .TextMatrix(i, .ColIndex("缺省付数")) = ""
            End If
        Next
    End With
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case pan_Cmd
        Item.Handle = picDown.hWnd
    Case pan_Tree
        Item.Handle = picTree.hWnd
    Case pan_WholeSet
        Item.Handle = picList.hWnd
    Case pan_WholeItems
        Item.Handle = picWholeSubItems.hWnd
    End Select
End Sub
Private Sub SetOkEnable()
    Dim blnEabled As Boolean
    blnEabled = Not lvwWholeSetItem.SelectedItem Is Nothing
    cmdOK.Enabled = blnEabled
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2010-09-02 15:21:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    Dim lngHeight As Long
    lngHeight = picDown.Height \ Screen.TwipsPerPixelY
    
    With dkpMan
        Set objPane = .CreatePane(pan_Cmd, 300, 100, DockLeftOf, Nothing)
        objPane.Title = "按钮": objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoHideable Or PaneNoFloatable
        objPane.Handle = picDown.hWnd
        objPane.MaxTrackSize.Height = lngHeight: objPane.MinTrackSize.Height = lngHeight:
        objPane.Tag = pan_Cmd
        
        Set objPane = .CreatePane(pan_Tree, 300, 100, DockTopOf, objPane)
        objPane.Title = "成套分类": objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoHideable Or PaneNoFloatable
        objPane.Handle = picTree.hWnd: objPane.Tag = pan_Tree
        
         Set objPane = .CreatePane(pan_WholeSet, 400, 400, DockRightOf, objPane)
        objPane.Title = "成套项目"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd: objPane.Tag = pan_WholeSet
        
         Set objPane = .CreatePane(pan_WholeItems, 400, 400, DockBottomOf, objPane)
        objPane.Title = "成套项目组成"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picWholeSubItems.hWnd: objPane.Tag = pan_WholeItems
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
     zlRestoreDockPanceToReg Me, dkpMan, "区域"
    dkpMan.RecalcLayout: DoEvents
End Function

Public Function ShowSelect(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
     ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示成套项目选择器(选择器入口)
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '出参:rsOutSel-成功时,返回选择的成套项目(有字段:细目ID,编码,名称,序号,从属父号,执行科室....)
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-02 11:52:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mblnOk = False: mblnFirst = True
    Me.Show 1, frmMain
    Set rsOutSel = mrsOutSel
    ShowSelect = mblnOk
End Function
Private Function InitSelFinelds() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化选择的字段
    '返回:初始成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-02 14:12:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mrsOutSel = New ADODB.Recordset
    mrsOutSel.Fields.Append "成套ID", adBigInt, , adFldIsNullable
    mrsOutSel.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    mrsOutSel.Fields.Append "序号", adBigInt, , adFldIsNullable
    mrsOutSel.Fields.Append "从属父号", adBigInt, , adFldIsNullable
    mrsOutSel.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsOutSel.Fields.Append "付数", adDouble, , adFldIsNullable
    mrsOutSel.Fields.Append "数量", adDouble, , adFldIsNullable
    mrsOutSel.Fields.Append "单价", adDouble, , adFldIsNullable
    mrsOutSel.CursorLocation = adUseClient
    mrsOutSel.LockType = adLockOptimistic
    mrsOutSel.CursorType = adOpenStatic
    mrsOutSel.Open
    InitSelFinelds = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub 调整从属父号()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整从属父号
    '编制:刘兴洪
    '日期:2011-01-03 13:17:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lng父号 As Long, lng成套ID As Long
    Dim lngRow As Long
    
    With vsWholeSet
        lng成套ID = 0: lngRow = 1
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("收费项目"))) <> 0 Then
            
                lng父号 = Val(.Cell(flexcpData, i, .ColIndex("序号")))
                lng成套ID = Val(.TextMatrix(i, .ColIndex("成套ID")))
                .Cell(flexcpData, i, .ColIndex("序号")) = lngRow
                .TextMatrix(i, .ColIndex("序号")) = lngRow
                If Val(.TextMatrix(i, .ColIndex("从属父号"))) = 0 Then
                    For j = i + 1 To .Rows - 1
                        If lng成套ID = Val(.TextMatrix(j, .ColIndex("成套ID"))) And lng父号 = Val(.TextMatrix(j, .ColIndex("从属父号"))) Then
                            .TextMatrix(j, .ColIndex("从属父号")) = lngRow
                        End If
                    Next
                End If
                lngRow = lngRow + 1
            Else
                .Cell(flexcpData, i, .ColIndex("序号")) = ""
                .TextMatrix(i, .ColIndex("序号")) = ""
            End If
        Next
    End With
End Sub
Private Function BulidingRecord() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:构建选择的记录集
    '返回:构建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-02 14:23:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    If InitSelFinelds = False Then Exit Function

    
    With vsWholeSet
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("收费项目"))) <> 0 And GetVsGridBoolColVal(vsWholeSet, i, .ColIndex("选择")) Then
                mrsOutSel.AddNew
                mrsOutSel!成套ID = Val(Mid(lvwWholeSetItem.SelectedItem.Key, 2))
                mrsOutSel!收费细目ID = Val(.Cell(flexcpData, i, .ColIndex("收费项目")))
                mrsOutSel!序号 = Val(.Cell(flexcpData, i, .ColIndex("序号")))
                mrsOutSel!从属父号 = Get从属父号(i, Val(.TextMatrix(i, .ColIndex("从属父号"))))
                mrsOutSel!付数 = Val(.TextMatrix(i, .ColIndex("缺省付数")))
                mrsOutSel!数量 = Val(.TextMatrix(i, .ColIndex("缺省数量")))
                mrsOutSel!单价 = Val(.TextMatrix(i, .ColIndex("缺省价格")))
                mrsOutSel!执行科室ID = Val(.Cell(flexcpData, i, .ColIndex("缺省执行科室")))
                mrsOutSel.Update
            End If
        Next
    End With
    If mrsOutSel.RecordCount = 0 Then
        MsgBox "未选择成套项目,请检查", vbOKOnly + vbInformation, gstrSysName
        vsWholeSet.SetFocus
        Exit Function
    End If
    BulidingRecord = True
End Function
Private Function Get从属父号(ByVal lngRow As Long, ByVal lng从属父号 As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取从属父号
    '返回:获取从属父号:如果主项未选择,则不从属任何项,返回0
    '编制:刘兴洪
    '日期:2011-01-02 16:02:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If lng从属父号 = 0 Then Exit Function
    With vsWholeSet
        For i = lngRow - 1 To 1 Step -1
            If Val(.Cell(flexcpData, i, .ColIndex("序号"))) = lng从属父号 Then
                If GetVsGridBoolColVal(vsWholeSet, i, .ColIndex("选择")) Then
                    Get从属父号 = lng从属父号
                Else
                    Get从属父号 = 0
                End If
                Exit Function
            End If
        Next
    End With
    Get从属父号 = 0
End Function


Public Sub FillWholeSetTree()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:填充成套分类数据
    '编制:刘兴洪
    '日期:2010-08-24 14:55:07
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, objNode As Node
    Dim strPreKey As String
    Err = 0: On Error GoTo Errhand:
    strSQL = "" & _
    "   Select id,上级ID,编码,名称 " & _
    "   From 成套项目分类  " & _
    "   Start with 上级id is null Connect by Prior   Id=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With TvwWholeSet
        If Not .SelectedItem Is Nothing Then strPreKey = .SelectedItem.Key
        .Nodes.Clear
       Set objNode = .Nodes.Add(, , "Root", "所有成套", "Close", "Expend")
       objNode.Expanded = True
       objNode.Sorted = True
       Do While Not rsTemp.EOF
            If IsNull(rsTemp!上级ID) Then
                Set objNode = .Nodes.Add("Root", tvwChild, "K" & Nvl(rsTemp!ID), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称), "Close", "Expend")
            Else
                Set objNode = .Nodes.Add("K" & rsTemp!上级ID, tvwChild, "K" & Nvl(rsTemp!ID), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称), "Close", "Expend")
            End If
            objNode.Sorted = True
            If objNode.Key = strPreKey Then
                objNode.EnsureVisible
                objNode.Selected = True
                objNode.Expanded = True
            End If
            objNode.Sorted = True
            rsTemp.MoveNext
       Loop
       TvwWholeSet.Tag = ""
       If .SelectedItem Is Nothing Then .Nodes("Root").Selected = True
       If Not .SelectedItem Is Nothing Then
            Call tvwWholeSet_NodeClick(.SelectedItem)
       End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function FillWholeItem(ByVal lng分类id As Long, Optional blnSearch As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:填充成套项目
    '入参:lng分类id-分类ID,0-所有分类
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-25 15:41:48
    '问题:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim mrs成套项目 As ADODB.Recordset, strSQL As String, strWhere As String
    Dim strPreKey As String, objListItem As ListItem, lngCol As Long, strInput As String
    On Error GoTo errHandle
    
    Screen.MousePointer = vbHourglass
    If Not lvwWholeSetItem.SelectedItem Is Nothing Then
        strPreKey = lvwWholeSetItem.SelectedItem.Key
    End If
    If blnSearch = False Or mrs成套项目 Is Nothing Then
            strWhere = " And ( A.人员ID=[2] "
            'If InStr(1, mstrPrivs, ";本科成套方案;") > 0 Then
                strWhere = strWhere & " OR Exists(Select 1 From 成套项目使用科室 A1 ,部门人员 B1 Where A1.成套ID=A.ID And A1.科室ID=B1.部门Id and B1.人员id=[2]) "
            'End If
            'If InStr(1, mstrPrivs, ";全院成套方案;") > 0 Then
                strWhere = strWhere & " OR nvl(A.范围,0)=0 "
            'End If
            strWhere = strWhere & ")"
            strSQL = "" & _
            "   Select  A.Id,A.分类ID,A.编码,A.名称,A.拼音,A.五笔,decode(nvl(范围,0),0,'全院',1,'指定科室',decode(A.人员id,Null,'指定操作员',B.姓名)) As 使用范围," & _
            "              C.名称 as 所属分类 " & _
            "   From 成套收费项目 A,人员表 B " & _
                    IIf(lng分类id = 0, ",成套项目分类 C", " ,(Select ID,上级ID,编码,名称 From  成套项目分类  Start With Id =[1] Connect By Prior Id=上级id ) C") & _
            "   Where a.人员id=b.Id(+) And A.分类id=C.ID " & strWhere
            Set mrs成套项目 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng分类id, UserInfo.ID)
    End If
     If Trim(txtFind.Text) <> "" Then
            strInput = Trim(txtFind.Text)
            If IsNumeric(strInput) Then
                    strInput = "'" & gstrLike & Trim(txtFind.Text) & "%'"
                    mrs成套项目.Filter = "编码 like " & UCase(strInput)
            ElseIf zlCommFun.IsNumOrChar(strInput) Then
                    'gbytCode As Byte '简码生成方式，0-拼音,1-五笔,2-两者
                    strInput = "'" & gstrLike & Trim(txtFind.Text) & "%'"
                    If gbytCode = 0 Then
                        mrs成套项目.Filter = "拼音 like " & UCase(strInput)
                    ElseIf gbytCode = 1 Then
                        mrs成套项目.Filter = "五笔 like " & UCase(strInput)
                    Else
                        mrs成套项目.Filter = "拼音 like " & UCase(strInput) & " Or 五笔 like " & UCase(strInput)
                    End If
            Else
                    strInput = "'" & gstrLike & Trim(txtFind.Text) & "%'"
                    mrs成套项目.Filter = "名称 like " & strInput
            End If
    Else
           mrs成套项目.Filter = 0
    End If
    If mrs成套项目.RecordCount <> 0 Then mrs成套项目.MoveFirst
    LockWindowUpdate lvwWholeSetItem.hWnd
    mblnNotClick = True
    With lvwWholeSetItem
        .ListItems.Clear
        .View = lvwReport
        Do While Not mrs成套项目.EOF
            '添加节点
            Set objListItem = .ListItems.Add(, "K" & mrs成套项目!ID, Nvl(mrs成套项目!编码) & "-" & Nvl(mrs成套项目!名称), "Item", "Item")
            objListItem.Tag = Nvl(mrs成套项目!分类ID)
            ' "名称,1500,0,1;编码,800,0,2;简码,1400,0,0;使用范围,400,0,0;所属分类,2400,0,0"
            '根据ListView的列名从数据库取数
            For lngCol = 2 To lvwWholeSetItem.ColumnHeaders.Count
                objListItem.SubItems(lngCol - 1) = Nvl(mrs成套项目.Fields(lvwWholeSetItem.ColumnHeaders(lngCol).Text))
            Next
            If mrs成套项目.AbsolutePosition = 1 Then '缺省为第一行选中
                objListItem.Selected = True
            End If
            If objListItem.Key = strPreKey Then
                objListItem.Selected = True
                objListItem.EnsureVisible
            End If
            mrs成套项目.MoveNext
        Loop
        lvwWholeSetItem.Checkboxes = True
    End With
    mblnNotClick = False
    If blnSearch = False Then lvwWholeSetItem.Tag = ""
    If Not lvwWholeSetItem.SelectedItem Is Nothing Then
        lvwWholeSetItem.Tag = ""
        Call lvwWholeSetItem_ItemClick(lvwWholeSetItem.SelectedItem)
    Else
        '清除成套项目的一些数据
        Call zlClearDownWholeSetItem
    End If
    LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    FillWholeItem = True
    Exit Function
errHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Screen.MousePointer = vbHourglass
        Resume
    End If
    mblnNotClick = False
    LockWindowUpdate 0
End Function
Public Function GetSelectWholeID() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择的成套数据
    '返回:成套ID;多个用逗号分离
    '编制:刘兴洪
    '日期:2011-01-03 10:49:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, i As Long
    Dim strTemp As String
    i = 1
    For Each objItem In lvwWholeSetItem.ListItems
        If objItem.Checked Or objItem.Selected Then
            strTemp = strTemp & "," & Mid(objItem.Key, 2)
            If Len(strTemp) > 1980 Then
               Exit For
            End If
        End If
    Next
    If strTemp <> "" Then
        GetSelectWholeID = Mid(strTemp, 2)
    ElseIf Not Me.lvwWholeSetItem.SelectedItem Is Nothing Then
        GetSelectWholeID = Mid(lvwWholeSetItem.SelectedItem.Key, 2)
   End If
End Function
Public Sub setLvwSelectColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置ListView选中的颜色
    '编制:刘兴洪
    '日期:2011-01-03 10:49:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, i As Long
    Err = 0: On Error GoTo Errhand:
    For Each objItem In lvwWholeSetItem.ListItems
        If objItem.Checked Then
            objItem.ForeColor = vbBlue
            objItem.Bold = True
        Else
            objItem.ForeColor = lvwWholeSetItem.ForeColor
            objItem.Bold = False
        End If
        For i = 0 To lvwWholeSetItem.ColumnHeaders.Count - 2
            objItem.ListSubItems(i + 1).ForeColor = objItem.ForeColor
            objItem.ListSubItems(i + 1).Bold = objItem.Bold
        Next
    Next
Errhand:
End Sub


Private Function FillWholeSetItemChildData(ByVal lng成套ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载成套项目子数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-25 17:42:49
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objListItem As ListItem, lng父号 As Long, j As Long, i As Long
    Dim rsOthers As ADODB.Recordset, lng主项ID As Long
    Dim str成套 As String, blnCheck As Boolean
    Dim lng成套IDA As Long, lngSplit As Long
    Dim lngTemp As Long
    Dim lngSelect成套ID As Long, m As Long
    On Error GoTo errHandle
    str成套 = GetSelectWholeID  '以后处理多选情况
    lngSelect成套ID = 0: blnCheck = False
    If Not Me.lvwWholeSetItem.SelectedItem Is Nothing Then
        lngSelect成套ID = Val(Mid(lvwWholeSetItem.SelectedItem.Key, 2))
        blnCheck = lvwWholeSetItem.SelectedItem.Checked
    End If
    
    
    gstrSQL = "" & _
       "   Select /*+ Rule*/  D.成套ID,A.主项id, A.从项id, A.固有从属, A.从项数次 " & _
       "   From 收费从属项目 A, 成套收费项目组合 D,Table(f_Num2List([1])) M" & _
       "   Where A.主项id = D.收费细目id And D.成套id =M.Column_value"
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str成套)
    
    strSQL = "" & _
    "   Select /*+ rule */  '' as 标志,A.成套ID,J.名称 as 成套项目名称,A.序号,B.类别, A.成套id, A.收费细目id, B.编码, B.名称, B.计算单位, B.规格,  " & _
    "           A.从属父号,nvl(A.付数,0) as 付数, A.数量, A.单价, A.执行科室id, " & _
    "          decode(C.编码,NULL,'',C.编码||'-') ||C.名称 As 执行科室 " & _
    "   From 成套收费项目组合 A,成套收费项目 J,Table(f_Num2List([1])) M, 收费项目目录 B, 部门表 C ,药品规格 D,诊疗项目目录 E" & _
    "   Where A.收费细目id = B.ID and A.成套ID=J.ID and A.收费细目ID=D.药品ID(+) and D.药名ID=E.id(+) And A.执行科室id = C.ID(+)  " & _
    "               And A.成套id =M.Column_value " & _
    "   Order By decode([2],A.成套ID,0,A.成套ID),A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str成套, lngSelect成套ID)
    
    With vsWholeSet
        .Redraw = flexRDNone
        .Clear 1
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = .ColIndex("标志"): .SubtotalPosition = flexSTAbove
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        lng父号 = 0: lng成套IDA = 0: lngSplit = 0
        .MergeCells = flexMergeRestrictRows
        
        Do While Not rsTemp.EOF
ReDO:
            If Val(Nvl(rsTemp!从属父号)) = 0 Then
                    lng父号 = Nvl(rsTemp!序号)
                    lng主项ID = Val(Nvl(rsTemp!收费细目ID))
            End If
            
            If lng成套IDA <> Val(Nvl(rsTemp!成套ID)) Then
                For m = 2 To .Cols - 1
                    .TextMatrix(i, m) = Nvl(rsTemp!成套项目名称)
                    .Cell(flexcpData, i, m) = ""
                    .MergeRow(i) = True
                    .Cell(flexcpBackColor, i, m, i, m) = &HFFC0C0
                    .Cell(flexcpFontBold, i, m) = True
                Next
                .IsSubtotal(i) = True
                .RowOutlineLevel(i) = 1
                lng成套IDA = Val(Nvl(rsTemp!成套ID))
                .TextMatrix(i, .ColIndex("成套ID")) = lng成套IDA
                If lngSelect成套ID = lng成套IDA Then
                     If "," & str成套 & "," <> "," & lngSelect成套ID & "," Then
                        .TextMatrix(i, .ColIndex("选择")) = IIf(blnCheck And chk缺省显示.Value = 1, -1, 0)
                     Else
                        .TextMatrix(i, .ColIndex("选择")) = IIf(chk缺省显示.Value = 1, -1, 0)
                     End If
                Else
                    .TextMatrix(i, .ColIndex("选择")) = IIf(chk缺省显示.Value = 1, -1, 0)
                End If
                If GetVsGridBoolColVal(vsWholeSet, i, .ColIndex("选择")) Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                Else
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
                End If
                i = i + 1: .Rows = .Rows + 1
                GoTo ReDO:
            End If
            .TextMatrix(i, .ColIndex("从属父号")) = Nvl(rsTemp!从属父号)
            If Val(.TextMatrix(i, .ColIndex("从属父号"))) > 0 Then
                rsOthers.Filter = "成套ID=" & Val(Nvl(rsTemp!成套ID)) & " and 主项ID=" & lng主项ID & " and 从项id=" & Val(Nvl(rsTemp!收费细目ID))
                If Not rsOthers.EOF Then
                    .TextMatrix(i, .ColIndex("从属数次")) = Val(Nvl(rsOthers!从项数次))
                    .Cell(flexcpData, i, .ColIndex("从属父号")) = Val(Nvl(rsOthers!固有从属))
                End If
            End If
            
            .TextMatrix(i, .ColIndex("类别")) = CStr(Nvl(rsTemp!类别))
            .TextMatrix(i, .ColIndex("选择")) = IIf(chk缺省显示.Value = 1, -1, 0)
            .TextMatrix(i, .ColIndex("序号")) = Nvl(rsTemp!序号)
            .Cell(flexcpData, i, .ColIndex("序号")) = Val(Nvl(rsTemp!序号))
            .TextMatrix(i, .ColIndex("收费项目")) = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
            .Cell(flexcpData, i, .ColIndex("收费项目")) = Val(Nvl(rsTemp!收费细目ID))
            .TextMatrix(i, .ColIndex("规格")) = Nvl(rsTemp!规格)
            .TextMatrix(i, .ColIndex("缺省付数")) = IIf(chk缺省显示.Value = 1, IIf(Val(Nvl(rsTemp!付数)) = 0, 1, Val(Nvl(rsTemp!付数))), "")
            .Cell(flexcpData, i, .ColIndex("缺省付数")) = IIf(Val(Nvl(rsTemp!付数)) = 0, 1, Val(Nvl(rsTemp!付数)))
            .TextMatrix(i, .ColIndex("缺省数量")) = IIf(chk缺省显示.Value = 1, FormatEx(Val(Nvl(rsTemp!数量)), 5, False), "")
            .Cell(flexcpData, i, .ColIndex("缺省数量")) = FormatEx(Val(Nvl(rsTemp!数量)), 5, False)
            .TextMatrix(i, .ColIndex("缺省价格")) = FormatEx(Val(Nvl(rsTemp!单价)), 8, False)
            .TextMatrix(i, .ColIndex("缺省执行科室")) = Nvl(rsTemp!执行科室)
            .Cell(flexcpData, i, .ColIndex("缺省执行科室")) = Val(Nvl(rsTemp!执行科室ID))
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTemp!计算单位)
            .TextMatrix(i, .ColIndex("成套ID")) = Nvl(rsTemp!成套ID)
            If Val(Nvl(rsTemp!从属父号)) = 0 Then
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 2
            ElseIf lng父号 = Val(.TextMatrix(i, .ColIndex("从属父号"))) Then
                    If i > 2 Then
                        If Val(.TextMatrix(i - 1, .ColIndex("从属父号"))) <> 0 Then
                            .IsSubtotal(i - 1) = False
                            .RowOutlineLevel(i - 1) = 2
                        End If
                    End If
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 3
            End If
            
            If lngSelect成套ID = Val(.TextMatrix(i, .ColIndex("成套ID"))) Then
               ' .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HFFC0C0  ' &H80000003
                If "," & str成套 & "," <> "," & lngSelect成套ID & "," Then
                    If blnCheck Then
                        .TextMatrix(i, .ColIndex("选择")) = IIf(chk缺省显示.Value = 1, -1, 0)
                    Else
                        .TextMatrix(i, .ColIndex("选择")) = 0
                        .TextMatrix(i, .ColIndex("缺省付数")) = ""
                        .TextMatrix(i, .ColIndex("缺省数量")) = ""
                    End If
                End If
            End If
            If GetVsGridBoolColVal(vsWholeSet, i, .ColIndex("选择")) Then
                '设置颜色
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        
        .Cell(flexcpBackColor, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = &HE7CFBA
        .Cell(flexcpBackColor, 1, .ColIndex("缺省付数"), .Rows - 1, .ColIndex("缺省付数")) = &HE7CFBA
        .Cell(flexcpBackColor, 1, .ColIndex("缺省数量"), .Rows - 1, .ColIndex("缺省数量")) = &HE7CFBA
        .Redraw = flexRDBuffered
        .ColWidth(.ColIndex("标志")) = 600
    End With
    '重新调整序号
    Call 调整从属父号

    FillWholeSetItemChildData = True
    Exit Function
errHandle:
    vsWholeSet.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
        vsWholeSet.Redraw = flexRDNone
    End If
End Function
Private Sub zlClearDownWholeSetItem()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除指定成套项目的组成和使用科室数据
    '编制:刘兴洪
    '日期:2010-08-25 16:35:03
    '问题:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsWholeSet
        .Rows = 2
        .Clear 1
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '选择
    If lvwWholeSetItem.SelectedItem Is Nothing Then Exit Sub
    If CheckIsValied = False Then Exit Sub
    
    If BulidingRecord = False Then Exit Sub
    If mrsOutSel.RecordCount = 0 Then
        MsgBox "未选择指定的成套项目或数据不正确,请检查", vbInformation + vbOKOnly, gstrSysName
        Set mrsOutSel = Nothing
        Exit Sub
    End If
    mblnOk = True
    Unload Me
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
         Bottom = cmdOK.Height / Screen.TwipsPerPixelY
End Sub

Private Sub dkpMan_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
   Bottom = cmdOK.Height / Screen.TwipsPerPixelY
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call zlClearDownWholeSetItem
    Call FillWholeSetTree
    Call SetOkEnable
    If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
    Case vbKeyF4
        TvwWholeSet.SetFocus
    Case vbKeyF6
          vsWholeSet.SetFocus
          vsWholeSet.Col = vsWholeSet.ColIndex("缺省数量")
    Case vbKeyF2
        Call cmdOK_Click
    Case Else
    End Select
End Sub

Private Sub Form_Load()
    lvwWholeSetItem.ListItems.Clear
    zlControl.LvwSelectColumns lvwWholeSetItem, mstrLvwWholeSet, True
    RestoreWinState Me, App.ProductName
    chk缺省显示.Value = IIf(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "上次缺省选择", "1") = "1", 1, 0)
    Call InitPanel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    SaveWinState Me, App.ProductName
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "上次缺省选择", chk缺省显示.Value
End Sub

Private Sub lvwWholeSetItem_GotFocus()
    Call SetOkEnable
End Sub

Private Sub lvwWholeSetItem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call FillWholeSetItemChildData(Val(Mid(Item.Key, 2)))
    Call setLvwSelectColor
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    Line1(0).x1 = 0: Line1(0).x2 = picDown.ScaleWidth
    Line1(1).x1 = Line1(0).x1: Line1(1).x2 = Line1(0).x2
    cmdCancel.Left = picDown.ScaleWidth - cmdCancel.Width - 50
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        lvwWholeSetItem.Left = .ScaleLeft
        lvwWholeSetItem.Width = .ScaleWidth
        lvwWholeSetItem.Height = .ScaleHeight - lvwWholeSetItem.Top
    End With
End Sub

Private Sub picTree_Resize()
    Err = 0: On Error Resume Next
    With picTree
        TvwWholeSet.Left = .ScaleLeft
        TvwWholeSet.Width = .ScaleWidth
        TvwWholeSet.Top = .ScaleTop
        TvwWholeSet.Height = .ScaleHeight
    End With
End Sub

Private Sub picWholeSubItems_Resize()
        Err = 0: On Error Resume Next
        With picWholeSubItems
             chk缺省显示.Left = .ScaleWidth - chk缺省显示.Width - 50
            vsWholeSet.Left = .ScaleLeft
            vsWholeSet.Width = .ScaleWidth
            vsWholeSet.Height = .ScaleHeight - vsWholeSet.Top
        End With
End Sub

Private Sub tvwWholeSet_NodeClick(ByVal Node As MSComctlLib.Node)
        '加载成套项目数据
        If TvwWholeSet.Tag <> Node.Key Then
            TvwWholeSet.Tag = Node.Key
            Call FillWholeItem(Val(Mid(Node.Key, 2)))
        End If
        Call SetOkEnable
End Sub

Private Sub txtFind_Change()
    If TvwWholeSet.SelectedItem Is Nothing Then
        FillWholeItem 0, True
    Else
        FillWholeItem Val(Mid(TvwWholeSet.SelectedItem.Key, 2)), True
    End If
    DoEvents
    If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsWholeSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, blnCheck As Boolean
    With vsWholeSet
        Select Case Col
        Case .ColIndex("缺省数量")
            .TextMatrix(Row, Col) = FormatEx(Val(Val(.TextMatrix(Row, Col))), 5, False)
            If Val(.TextMatrix(Row, Col)) <> 0 And GetVsGridBoolColVal(vsWholeSet, Row, .ColIndex("选择")) = False Then
                 .TextMatrix(Row, .ColIndex("选择")) = -1
            ElseIf Val(.TextMatrix(Row, Col)) = 0 Then
                 .TextMatrix(Row, .ColIndex("选择")) = 0
                 .TextMatrix(Row, .ColIndex("缺省付数")) = ""
            End If
            If GetVsGridBoolColVal(vsWholeSet, Row, .ColIndex("选择")) Then
                   .TextMatrix(Row, .ColIndex("缺省付数")) = IIf(Val(.TextMatrix(Row, .ColIndex("缺省付数"))) = 0, .Cell(flexcpData, Row, .ColIndex("缺省付数")), .TextMatrix(Row, .ColIndex("缺省付数")))
            Else
                 .TextMatrix(Row, .ColIndex("缺省付数")) = ""
            End If
        Case .ColIndex("缺省付数")
            .TextMatrix(Row, Col) = FormatEx(Val(Val(.TextMatrix(Row, Col))), 0, False)
            If Val(.TextMatrix(Row, Col)) <> 0 And GetVsGridBoolColVal(vsWholeSet, Row, .ColIndex("选择")) = False Then
                 .TextMatrix(Row, .ColIndex("选择")) = -1
                 .TextMatrix(Row, .ColIndex("缺省数量")) = IIf(Val(.TextMatrix(Row, .ColIndex("缺省数量"))) = 0, .Cell(flexcpData, Row, .ColIndex("缺省数量")), .TextMatrix(Row, .ColIndex("缺省数量")))
            ElseIf Val(.TextMatrix(Row, Col)) = 0 Then
                 .TextMatrix(Row, .ColIndex("选择")) = 0
                 .TextMatrix(Row, .ColIndex("缺省数量")) = ""
            End If
        Case .ColIndex("选择")
            blnCheck = GetVsGridBoolColVal(vsWholeSet, Row, Col)
            If Val(.Cell(flexcpData, Row, .ColIndex("收费项目"))) <> 0 Then
                If blnCheck Then
                    If Val(.TextMatrix(Row, .ColIndex("缺省数量"))) = 0 Then
                        .TextMatrix(Row, .ColIndex("缺省数量")) = FormatEx(Val(.Cell(flexcpData, Row, .ColIndex("缺省数量"))), 5, False)
                        .TextMatrix(Row, .ColIndex("缺省付数")) = .Cell(flexcpData, Row, .ColIndex("缺省付数"))
                    End If
                Else
                    .TextMatrix(Row, .ColIndex("缺省数量")) = ""
                    .TextMatrix(Row, .ColIndex("缺省付数")) = ""
                End If
            End If
            If Val(.Cell(flexcpData, Row, .ColIndex("收费项目"))) = 0 Then
                '选中单据,此张全选,或全清
                For i = Row + 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("成套ID"))) = Val(.TextMatrix(Row, .ColIndex("成套ID"))) And Val(.Cell(flexcpData, i, .ColIndex("收费项目"))) <> 0 Then
                         If blnCheck Then
                            If Val(.TextMatrix(i, .ColIndex("缺省数量"))) = 0 Then
                                .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("缺省数量"))), 5, False)
                            End If
                            If Val(.TextMatrix(i, .ColIndex("缺省付数"))) = 0 Then
                                .TextMatrix(i, .ColIndex("缺省付数")) = .Cell(flexcpData, i, .ColIndex("缺省付数"))
                            End If
                            
                         Else
                            .TextMatrix(i, .ColIndex("缺省数量")) = ""
                            .TextMatrix(i, .ColIndex("缺省付数")) = ""
                         End If
                          .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
                    Else
                        Exit For
                    End If
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = IIf(blnCheck, vbBlue, .ForeColor)
                Next
            End If
            
        Case Else
        End Select
        Call Set从属项目(Row)
        Call ReCale从属项目(Row, Val(.TextMatrix(Row, .ColIndex("缺省数量"))))
        
        If GetVsGridBoolColVal(vsWholeSet, Row, .ColIndex("选择")) Then
            .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlue
        Else
            .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = .ForeColor
        End If
    End With
End Sub
Private Sub Set从属项目(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置从属项目
    '参数:lngRow-设置指定的行
    '编制:刘兴洪
    '日期:2011-01-02 16:08:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsWholeSet
        If Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) = 0 Then
            For i = lngRow + 1 To .Rows - 1
                 If Val(.Cell(flexcpData, i, .ColIndex("收费项目"))) = 0 Then Exit Sub
                 If Val(.Cell(flexcpData, lngRow, .ColIndex("序号"))) = Val(.TextMatrix(i, .ColIndex("从属父号"))) Then
                        .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(lngRow, .ColIndex("选择"))
                      If GetVsGridBoolColVal(vsWholeSet, i, .ColIndex("选择")) Then
                            If Val(.TextMatrix(i, .ColIndex("缺省数量"))) = 0 Then
                                .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("缺省数量"))), 5, False)
                            End If
                            If Val(.TextMatrix(i, .ColIndex("缺省付数"))) = 0 Then
                                .TextMatrix(i, .ColIndex("缺省付数")) = .Cell(flexcpData, i, .ColIndex("缺省付数"))
                            End If
                            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                      Else
                            .TextMatrix(i, .ColIndex("缺省数量")) = ""
                            .TextMatrix(i, .ColIndex("缺省付数")) = ""
                            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
                      End If
                 Else
                    Exit For
                 End If
            Next
        End If
    End With
End Sub
Private Sub vsWholeSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long, int固定从属 As Integer
    
    With vsWholeSet
    
        Select Case Col
        Case .ColIndex("选择")
        Case .ColIndex("缺省数量")
            If Val(.Cell(flexcpData, Row, .ColIndex("收费项目"))) = 0 Then
                Cancel = True: Exit Sub
            End If
            If Val(.TextMatrix(Row, .ColIndex("从属父号"))) <> 0 Then
                int固定从属 = Val(.Cell(flexcpData, i, .ColIndex("从属父号")))
                If int固定从属 = 1 Or 2 Then  '固定的从属和按比例从属
                    Cancel = True
                End If
                '非固有从属,不允许改数量
            End If
        Case .ColIndex("缺省付数")
            If Val(.Cell(flexcpData, Row, .ColIndex("收费项目"))) = 0 Then
                Cancel = True: Exit Sub
            End If
            If Not .TextMatrix(Row, .ColIndex("类别")) = "7" Then
                Cancel = True
            End If
            If Val(.TextMatrix(Row, .ColIndex("从属父号"))) <> 0 Then
                int固定从属 = Val(.Cell(flexcpData, i, .ColIndex("从属父号")))
                If int固定从属 = 1 Or 2 Then  '固定的从属和按比例从属
                    Cancel = True
                End If
                '非固有从属,不允许改数量
            End If
        Case Else
             Cancel = True
        End Select
    End With
End Sub

Private Sub vsWholeSet_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsWholeSet
        Select Case Col
        Case .ColIndex("标志")
             Cancel = True
        Case Else
        End Select
    End With
End Sub


Private Sub lvwWholeSetItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo errHandle
    If mintColumn1 = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwWholeSetItem.SortOrder = IIf(lvwWholeSetItem.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn1 = ColumnHeader.Index - 1
        lvwWholeSetItem.SortKey = mintColumn1
        lvwWholeSetItem.SortOrder = lvwAscending
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwWholeSetItem_DblClick()
    If Not mblnItem Then Exit Sub
    If Me.lvwWholeSetItem.SelectedItem Is Nothing Then Exit Sub
    Call cmdOK_Click
End Sub
Private Sub lvwWholeSetItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '刘兴洪:27327
    '为成套项目维护时,需要单独处理
    mblnItem = True
    If lvwWholeSetItem.Tag <> Item.Key Then
        Call FillWholeSetItemChildData(Val(Mid(Item.Key, 2)))
    End If
    lvwWholeSetItem.Tag = Item.Key
    SetOkEnable
End Sub

Private Sub lvwWholeSetItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lvwWholeSetItem.SelectedItem Is Nothing Then Exit Sub
        Call cmdOK_Click
    End If
End Sub
Private Sub lvwWholeSetItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub
Private Sub vsWholeSet_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsWholeSet
        If .Col >= .ColIndex("缺省数量") And .Row = .Rows - 1 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Select Case .Col
        Case .ColIndex("缺省数量")
            If .Row < .Rows - 1 Then
                .Col = .Col: .Row = .Row + 1
            End If
        Case .ColIndex("缺省付数")
            If .Row < .Rows - 1 Then
                .Col = .Col: .Row = .Row + 1
            End If
        Case .ColIndex("选择")
            If .Row < .Rows - 1 Then
                .Col = .Col: .Row = .Row + 1
            End If
        End Select
    End With
End Sub

Private Sub vsWholeSet_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsWholeSet
        Select Case Col
        Case .ColIndex("缺省数量"), .ColIndex("缺省付数")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsWholeSet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsWholeSet_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsWholeSet
        Select Case .Col
            Case .ColIndex("缺省数量")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case .ColIndex("缺省付数")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    KeyAscii = 0
                End If
            Case Else

        End Select
    End With
End Sub
Private Sub vsWholeSet_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    '数据验证
    With vsWholeSet
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("缺省数量")
                If zlDblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                strKey = Format(Val(strKey), "0.00000")
                .EditText = strKey
                .TextMatrix(Row, .Col) = strKey
            Case .ColIndex("缺省付数")
                If zlDblIsValid(strKey, 4, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                strKey = Format(Val(strKey), "0.00000")
                .EditText = strKey
                .TextMatrix(Row, .Col) = strKey
            End Select
    End With
End Sub
Private Sub ReCale从属项目(ByVal lngRow As Long, ByVal dblNum As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算从属项目数量
    '入参:dblNum-数量
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-31 11:30:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int固定从属 As Integer, i As Long, dblTemp As Double
    
    If dblNum = 0 Then Exit Sub
    With vsWholeSet
        If Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) <> 0 Then Exit Sub
        For i = lngRow + 1 To .Rows - 1
             If Val(.TextMatrix(i, .ColIndex("从属父号"))) = Val(.Cell(flexcpData, lngRow, .ColIndex("序号"))) Then
                int固定从属 = Val(.Cell(flexcpData, i, .ColIndex("从属父号")))
                If int固定从属 = 0 Then '非固有从属
                    'dblTemp = IIf(dblNum < 0, -1, 1) * Val(.TextMatrix(i, .ColIndex("从属数次")))
                    ' .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(dblTemp, 5)
                ElseIf int固定从属 = 1 Then '固定的从属
                    dblTemp = IIf(dblNum < 0, -1, 1) * IIf(Val(.TextMatrix(i, .ColIndex("从属数次"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("从属数次"))))
                    .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(dblTemp, 5)
                ElseIf int固定从属 = 2 Then '按比例从属
                    dblTemp = dblNum * Val(.TextMatrix(i, .ColIndex("从属数次")))
                    .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(dblTemp, 5)
                End If
             End If
        Next
    End With
End Sub
Public Function CheckIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查录入的单价是否有效
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-03 11:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lng收费细目ID As Long
    Dim bln中药 As Boolean
    
    With vsWholeSet
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsWholeSet, i, .ColIndex("选择")) Then
                 lng收费细目ID = Val(.Cell(flexcpData, i, .ColIndex("收费项目")))
                 bln中药 = .TextMatrix(i, .ColIndex("类别")) = "7"
                 If lng收费细目ID <> 0 Then
                    For j = i + 1 To .Rows - 1
                        If Val(.Cell(flexcpData, j, .ColIndex("收费项目"))) <> 0 Then
                            If GetVsGridBoolColVal(vsWholeSet, j, .ColIndex("选择")) Then
                                If lng收费细目ID = Val(.Cell(flexcpData, j, .ColIndex("收费项目"))) Then
                                    MsgBox "收费项目『" & .TextMatrix(j, .ColIndex("收费项目")) & " 』在第" & .TextMatrix(j, .ColIndex("序号")) & "行中已经存在,请检查!"
                                    .Row = j: .SetFocus
                                    Exit Function
                                End If
                                If (bln中药 And .TextMatrix(j, .ColIndex("类别")) <> "7") Then
                                    MsgBox "收费项目『" & .TextMatrix(j, .ColIndex("收费项目")) & " 』在第" & .TextMatrix(j, .ColIndex("序号")) & "行中包含了非中草药项目,请检查!"
                                    .Row = j: .SetFocus
                                    Exit Function
                                End If
                                
                                If (bln中药 = False And .TextMatrix(j, .ColIndex("类别")) = "7") Then
                                    MsgBox "收费项目『" & .TextMatrix(j, .ColIndex("收费项目")) & " 』在第" & .TextMatrix(j, .ColIndex("序号")) & "行中包含了中草药项目,请检查!"
                                    .Row = j: .SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next
    End With
    CheckIsValied = True
End Function
