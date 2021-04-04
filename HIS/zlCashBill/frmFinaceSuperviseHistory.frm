VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFinaceSuperviseHistory 
   BorderStyle     =   0  'None
   Caption         =   "历史收款信息"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRollingCurtain 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   5715
      ScaleHeight     =   3090
      ScaleWidth      =   3825
      TabIndex        =   13
      Top             =   4575
      Width           =   3825
      Begin VB.PictureBox picImgPlanRC 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   17
         Top             =   45
         Width           =   210
         Begin VB.Image imgColPlanRC 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinaceSuperviseHistory.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   1800
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinaceSuperviseHistory.frx":054E
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
         OutlineBar      =   1
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2355
         TabIndex        =   15
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   225
      ScaleHeight     =   2520
      ScaleWidth      =   3435
      TabIndex        =   11
      Top             =   4125
      Width           =   3435
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   -15
         TabIndex        =   12
         Top             =   -15
         Width           =   2865
         _Version        =   589884
         _ExtentX        =   5054
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picCollect 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   120
      ScaleHeight     =   2490
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   510
      Width           =   10170
      Begin VB.PictureBox picImgPlanCollect 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   16
         Top             =   1035
         Width           =   210
         Begin VB.Image imgColPlanCollect 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinaceSuperviseHistory.frx":061B
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   15
            Width           =   195
         End
      End
      Begin VB.TextBox txtNO 
         Height          =   345
         Left            =   1020
         TabIndex        =   10
         Top             =   540
         Width           =   3360
      End
      Begin VB.ComboBox cboDate 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "重新过滤数据(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7995
         TabIndex        =   1
         Top             =   95
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCollect 
         Height          =   1800
         Left            =   0
         TabIndex        =   2
         Top             =   900
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinaceSuperviseHistory.frx":0B69
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
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   2325
         TabIndex        =   3
         Top             =   135
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   42532867
         CurrentDate     =   41520
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   4785
         TabIndex        =   4
         Top             =   135
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   42532867
         CurrentDate     =   41520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   5625
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         Caption         =   "收款单号"
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   607
         Width           =   840
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4485
         TabIndex        =   7
         Top             =   187
         Width           =   225
      End
      Begin VB.Label lblHistoryDate 
         AutoSize        =   -1  'True
         Caption         =   "过滤时间"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   187
         Width           =   840
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2355
         TabIndex        =   5
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFinaceSuperviseHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private Enum mPaneIndex
    EM_PN_CollectList = 270101  '收款列表
    EM_PN_DetailList = 270102   '收款及票据汇总
End Enum
Private mblnNotBrush As Boolean '不刷新数据
Private mobjChargeBill As clsChargeBill
Private mlngCollectID As Long '收款收款ID
Private mstrCollectNO As String   '收款单据号
Private mblnDel As Boolean   '是否已作废
Private mint记录性质 As Integer
Private Enum mPgIndex
    EM_PG_收款及票据 = 250101
    EM_PG_轧帐信息 = 250102
End Enum
Public Sub zlInitVar(ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '编制:刘兴洪
    '日期:2013-09-09 14:41:46
    '说明:加载窗体后,立即调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
End Sub
Private Sub InitPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2013-09-22 17:07:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(EM_PG_收款及票据, "收款及票据信息", mobjChargeBill.GetChargeAndBillTotalForm.hWnd, 0)
    objItem.Tag = EM_PG_收款及票据
    Set objItem = tbPage.InsertItem(EM_PG_轧帐信息, "轧帐信息", picRollingCurtain.hWnd, 0)
    objItem.Tag = EM_PG_轧帐信息
    With tbPage
        Set tbPage.PaintManager.Font = Me.Font
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化区域
    '编制:刘兴洪
    '日期:2013-09-22 17:48:09
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngHeight  As Long
    With dkpMan
        lngHeight = 2490 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(mPaneIndex.EM_PN_CollectList, 400, lngHeight, DockRightOf, Nothing)
        objPane.Title = "收款信息"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picCollect.hWnd
        objPane.MinTrackSize.Height = lngHeight * 0.9
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_DetailList, 400, 400, DockBottomOf, objPane)
        objPane.Title = "": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件
    '编制:刘兴洪
    '日期:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsCollect
        .Clear
        .Rows = 2: .Cols = 15: i = 1
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "记录性质": i = i + 1
        .TextMatrix(0, i) = "类别": i = i + 1
        .TextMatrix(0, i) = "收款单号": i = i + 1
        .TextMatrix(0, i) = "收费员": i = i + 1
        .TextMatrix(0, i) = "收款部门": i = i + 1
        .TextMatrix(0, i) = "收款说明": i = i + 1
        .TextMatrix(0, i) = "冲预交款": i = i + 1
        .TextMatrix(0, i) = "借入合计": i = i + 1
        .TextMatrix(0, i) = "借出合计": i = i + 1
        .TextMatrix(0, i) = "财务收款人": i = i + 1
        .TextMatrix(0, i) = "财务收款时间": i = i + 1
        .TextMatrix(0, i) = "作废人": i = i + 1
        .TextMatrix(0, i) = "作废时间": i = i + 1
        .ColData(0) = "-1|1"
       For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "收款部门" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "记录性质" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "收款单号" Or .ColKey(i) = "作废人" Or .ColKey(i) = "作废时间" Then .ColData(i) = "1|0"
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "收款单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsCollect, Me.Name, "收款信息列表", False
    End With
    With vsRollingCurtain
        .Clear
        .Rows = 2: .Cols = 19: i = 1
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "轧帐单号": i = i + 1
        .TextMatrix(0, i) = "轧帐类别": i = i + 1
        .TextMatrix(0, i) = "开始时间": i = i + 1
        .TextMatrix(0, i) = "终止时间": i = i + 1
        .TextMatrix(0, i) = "轧帐人": i = i + 1
        .TextMatrix(0, i) = "轧帐时间": i = i + 1
        .TextMatrix(0, i) = "收款部门": i = i + 1
        .TextMatrix(0, i) = "轧帐说明": i = i + 1
        .TextMatrix(0, i) = "冲预交款": i = i + 1
        .TextMatrix(0, i) = "借入合计": i = i + 1
        .TextMatrix(0, i) = "借出合计": i = i + 1
        .TextMatrix(0, i) = "小组收款人": i = i + 1
        .TextMatrix(0, i) = "小组收款时间": i = i + 1
        .TextMatrix(0, i) = "财务收款人": i = i + 1
        .TextMatrix(0, i) = "财务收款时间": i = i + 1
        .TextMatrix(0, i) = "作废人": i = i + 1
        .TextMatrix(0, i) = "作废时间": i = i + 1
        .ColData(0) = "-1|1"
       For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "收款部门" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "作废人" Or .ColKey(i) = "作废时间" Or _
               .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Then .ColData(i) = "1|0"
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False
    End With
End Sub
Private Function LoadHistoryData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载历史收款数据
    '返回:数据加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 17:08:50
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, i As Long, blnDel As Boolean
    Dim dtStartDate As Date, dtEndDate As Date
    On Error GoTo errHandle
    Call GetDateRange(dtStartDate, dtEndDate)
    
    If dtpEndDate - dtStartDate > 180 Then
        '以收费员作为统计条件
        Call MsgBox("你所设置的时间范围超过了半年,不允许大范围的查询", vbInformation + vbOKOnly, gstrSysName)
    End If
    If txtNO.Text <> "" Then
        strWhere = " And  A.NO = [3] "
    Else
        strWhere = " And  A.登记时间 Between [1] And [2] "
    End If
    
    strSQL = "" & _
    "   Select /*+ rule */a.Id,A.记录性质,decode(A.记录性质,4,'轧帐收款',5,'手工收款','轧帐归零') as 类别,a.No As 收款单号, a.收款员 As 收费员,  " & _
    "         b.名称 As 收款部门, a.摘要 As 收款说明, " & _
    "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
    "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
    "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计, " & _
    "         a.登记人 as 财务收款人,To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 财务收款时间,  " & _
    "         a.作废人, To_Char(a.作废时间, 'yyyy-mm-dd hh24:mi:ss') As 作废时间 " & _
    "  From 人员收缴记录 A, 部门表 B " & _
    "  Where a.收款部门id = b.Id(+) And a.记录性质  in ( 4,5,6) " & strWhere & _
    "  Order by 登记时间 desc,收款单号 desc,小组收款时间 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStartDate, dtEndDate, UCase(txtNO.Text))
    With vsCollect
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .FixedRows = 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("记录性质")) = NVL(rsTemp!记录性质)
            .TextMatrix(.Rows - 1, .ColIndex("类别")) = NVL(rsTemp!类别)
            .TextMatrix(.Rows - 1, .ColIndex("收款单号")) = NVL(rsTemp!收款单号)
            .TextMatrix(.Rows - 1, .ColIndex("收费员")) = NVL(rsTemp!收费员)
            '.TextMatrix(.Rows - 1, .ColIndex("收款部门")) = Nvl(rsTemp!收款部门)
            .TextMatrix(.Rows - 1, .ColIndex("收款说明")) = NVL(rsTemp!收款说明)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = NVL(rsTemp!冲预交款)
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = NVL(rsTemp!借入合计)
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = NVL(rsTemp!借出合计)
            .TextMatrix(.Rows - 1, .ColIndex("财务收款人")) = NVL(rsTemp!财务收款人)
            .TextMatrix(.Rows - 1, .ColIndex("财务收款时间")) = NVL(rsTemp!财务收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("作废人")) = NVL(rsTemp!作废人)
            .TextMatrix(.Rows - 1, .ColIndex("作废时间")) = NVL(rsTemp!作废时间)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then .Rows = .Rows - 1
'        If rsTemp.RecordCount <> 0 Then
'            Set .DataSource = rsTemp
'        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "收款部门" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "记录性质" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "收款单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        For i = 1 To .Rows - 1
            blnDel = Trim(.TextMatrix(i, .ColIndex("作废时间"))) <> ""
            If blnDel Then
                '作废记录，用红色字体
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
        Next
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsCollect, Me.Name, "收款信息列表", False
        If .Enabled And .Visible Then .SetFocus
    End With
    mblnNotBrush = False
    '加载明细数据
    Call LoadDetail
    LoadHistoryData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化窗体数据
    '编制:刘兴洪
    '日期:2013-09-11 17:46:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngCollectID = 0
    mstrCollectNO = ""
    mblnDel = False: mint记录性质 = 4
    Call InitPage
    Call InitGrid '初始化网格
    With cboDate
        .Clear
        .AddItem "今日"
        .ItemData(.NewIndex) = 1: .ListIndex = 0
        .AddItem "昨日"
        .ItemData(.NewIndex) = 2
        .AddItem "本周"
        .ItemData(.NewIndex) = 3
        .AddItem "上周"
        .ItemData(.NewIndex) = 4
        .AddItem "本月"
        .ItemData(.NewIndex) = 5
        .AddItem "指定时间"
        .ItemData(.NewIndex) = 9
    End With
    dtpEndDate.Value = zlDatabase.Currentdate
    dtpEndDate.MaxDate = Format(dtpEndDate.Value, "yyyy-mm-dd 23:59:59")
    dtpStartDate.Value = DateAdd("m", -1, dtpEndDate.Value)
    dtpStartDate.MaxDate = dtpEndDate.MaxDate
    Call SetCtrlVisible
End Sub
Private Sub SetCtrlVisible()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Visible属性
    '编制:刘兴洪
    '日期:2013-09-11 18:21:29
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date, dtStartDate As Date
    dtpStartDate.Visible = False: dtpEndDate.Visible = False
    lblEndDate.Visible = False
    Select Case cboDate.ItemData(cboDate.ListIndex)
    Case 1 '今日
        dtDate = zlDatabase.Currentdate
        lblRange.Caption = Format(dtDate, "yyyy-mm-dd 00:00:00") & "～" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 2 '昨日
        dtDate = DateAdd("d", -1, zlDatabase.Currentdate)
        lblRange.Caption = Format(dtDate, "yyyy-mm-dd 00:00:00") & "～" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 3 '本周
        dtDate = zlDatabase.Currentdate
        dtStartDate = DateAdd("d", -1 * (Weekday(dtDate) - 2), dtDate)
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "～" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 4 '上周
        dtDate = zlDatabase.Currentdate
        dtDate = DateAdd("d", -1 * (Weekday(dtDate) - 2), dtDate)
        dtStartDate = DateAdd("d", -7, dtDate)
        dtDate = DateAdd("d", 6, dtStartDate)
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "～" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 5 '本月
        dtDate = zlDatabase.Currentdate
        dtStartDate = CDate(Format(dtDate, "yyyy") & "-" & Month(dtDate) & "-01")
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "～" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 9 '指定日期
        lblRange.Visible = False
        dtpStartDate.Visible = True: dtpEndDate.Visible = True
        lblEndDate.Visible = True
        If dtpStartDate.Enabled And dtpStartDate.Visible Then dtpStartDate.SetFocus
    End Select
End Sub
Private Function GetDateRange(ByRef dtStartDate As Date, ByRef dtEndDate As Date) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取时间范围
    '入参:dtStartDate-开始时间
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 18:45:57
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant
    On Error GoTo errHandle
    Select Case cboDate.ItemData(cboDate.ListIndex)
    Case 9 '指定日期
        dtStartDate = dtpStartDate.Value
        dtEndDate = dtpEndDate.Value
    Case Else '1, 2, 3, 4, 5 '今日 '昨日'本周 '上周'本月
        varData = Split(lblRange.Caption, "～")
        dtStartDate = CDate(varData(0))
        dtEndDate = CDate(varData(1))
    End Select
    GetDateRange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cboDate_Click()
    Call SetCtrlVisible
End Sub

Private Sub cmdRefresh_Click()
    Call LoadHistoryData
End Sub

Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub

Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Load()
    mblnDel = False
    Set mobjChargeBill = New clsChargeBill
    Call InitPanel
    Call InitFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "收款信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
    Set mobjChargeBill = Nothing
End Sub
 
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
        tbPage.Width = .ScaleWidth
    End With
End Sub

Private Sub picRollingCurtain_Resize()
    Err = 0: On Error Resume Next
    With picRollingCurtain
        vsRollingCurtain.Left = .ScaleLeft
        vsRollingCurtain.Top = .ScaleTop
        vsRollingCurtain.Height = .ScaleHeight
        vsRollingCurtain.Width = .ScaleWidth
    End With
End Sub
Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
    zlCommFun.OpenIme False
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Or Trim(txtNO.Text) = "" Then Exit Sub
    txtNO.Text = GetFullNO(Trim(txtNO.Text), 141)
    Call LoadHistoryData
End Sub

Private Sub vsCollect_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
End Sub

Private Sub vsRollingCurtain_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsRollingCurtain)
    vsRollingCurtain.Tag = "0"
End Sub

Private Sub vsRollingCurtain_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsRollingCurtain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsRollingCurtain, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsRollingCurtain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub picCollect_Resize()
    Err = 0: On Error Resume Next
    Line1.X2 = picCollect.Width
    With picCollect
        vsCollect.Left = .ScaleLeft
        vsCollect.Top = txtNO.Top + txtNO.Height + 100
        vsCollect.Height = .ScaleHeight - vsCollect.Top - 50
        vsCollect.Width = .ScaleWidth
    End With
End Sub
Private Sub vsCollect_GotFocus()
    Call zl_VsGridGotFocus(vsCollect)
    vsCollect.Tag = "1"
End Sub
Private Sub vsCollect_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsCollect)
    vsCollect.Tag = "0"
End Sub
Private Sub vsCollect_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "收款信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsCollect_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsCollect, OldRow, NewRow, OldCol, NewCol)
    'If OldRow = NewRow Or mblnNotBrush Then Exit Sub
    Call LoadDetail
End Sub
Private Sub vsCollect_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "收款信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Function LoadDetail() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载明细数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 11:17:09
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date, strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long
    Dim lng收款ID As Long, strNO As String, blnDel As Boolean
    With vsCollect
        If .Row >= 1 Then
            lng收款ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strNO = Trim(.TextMatrix(.Row, .ColIndex("收款单号")))
            blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
             mint记录性质 = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
        End If
    End With
    mlngCollectID = lng收款ID
    mstrCollectNO = strNO
    mblnDel = blnDel
    
    On Error GoTo errHandle
    dtStartDate = CDate("1991-01-01"): dtEndDate = dtStartDate
    
    If Val(tbPage.Selected.Tag) = mPgIndex.EM_PG_收款及票据 Then
        '加载票据汇总
        If lng收款ID = 0 Then
            mobjChargeBill.ClearChargeAndBillTotalForm
        Else
            If mobjChargeBill.LoadChargeAndBillTotalData(Me, mlngModule, mstrPrivs, 4, lng收款ID, dtStartDate, dtEndDate, True, blnDel) = False Then Exit Function
        End If
        LoadDetail = True
        'Exit Function
    End If
    '加载轧帐信息列表
    strSQL = "" & _
    "   Select /*+ rule */a.Id,a.No As 轧帐单号,Substr(Decode(是否挂号,1,',挂号','') || Decode(是否就诊卡,1,',就诊卡','') || Decode(是否消费卡,1,',消费卡','') || Decode(是否收费,1,',收费','') || Decode(是否结帐,1,',结帐','') || Decode(预交类别,1,',预交',2,',门诊预交',3,',住院预交',''),2) As 轧帐类别," & _
    "         a.开始时间, a.终止时间, a.登记人 As 轧帐人, a.登记时间 As 轧帐时间,  " & _
    "         b.名称 As 收款部门, a.摘要 As 轧帐说明, " & _
    "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
    "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
    "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计, " & _
    "         a.小组收款人, To_Char(a.小组收款时间, 'yyyy-mm-dd hh24:mi:ss') As 小组收款时间,  " & _
    "         a.财务收款人,To_Char(a.财务收款时间, 'yyyy-mm-dd hh24:mi:ss') As 财务收款时间,  " & _
    "         a.作废人, To_Char(a.作废时间, 'yyyy-mm-dd hh24:mi:ss') As 作废时间 " & _
    "  From 人员收缴记录 A, 部门表 B " & _
    "  Where a.收款部门id = b.Id(+) And a.记录性质  in (3,1) and A.财务收款ID= [1]" & _
    "  Order by 小组轧账ID Desc,登记时间 desc,轧帐单号 desc,小组收款时间 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng收款ID)
    
    With vsRollingCurtain
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .OutlineBar = flexOutlineBarComplete
        .FixedRows = 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = NVL(rsTemp!轧帐单号)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐类别")) = NVL(rsTemp!轧帐类别)
            .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = NVL(rsTemp!开始时间)
            .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = NVL(rsTemp!终止时间)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐人")) = NVL(rsTemp!轧帐人)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = NVL(rsTemp!轧帐时间)
            '.TextMatrix(.Rows - 1, .ColIndex("收款部门")) = Nvl(rsTemp!收款部门)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐说明")) = NVL(rsTemp!轧帐说明)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = NVL(rsTemp!冲预交款)
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = NVL(rsTemp!借入合计)
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = NVL(rsTemp!借出合计)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款人")) = NVL(rsTemp!小组收款人)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款时间")) = NVL(rsTemp!小组收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("财务收款人")) = NVL(rsTemp!财务收款人)
            .TextMatrix(.Rows - 1, .ColIndex("财务收款时间")) = NVL(rsTemp!财务收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("作废人")) = NVL(rsTemp!作废人)
            .TextMatrix(.Rows - 1, .ColIndex("作废时间")) = NVL(rsTemp!作废时间)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then .Rows = .Rows - 1
'        If rsTemp.RecordCount <> 0 Then
'            Set .DataSource = rsTemp
'        End If
        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "收款部门" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "轧帐单号" Then
                .OutlineCol = i
            End If
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        For i = 1 To .Rows - 1
            blnDel = Trim(.TextMatrix(i, .ColIndex("作废时间"))) <> ""
            If blnDel Then
                '作废记录，用红色字体
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            If Val(.TextMatrix(i, .ColIndex("ID"))) <> 0 Then
                If .TextMatrix(i, .ColIndex("小组收款时间")) = "" Then
                    .IsSubtotal(i) = True
                Else
                    .RowOutlineLevel(i) = 1
                End If
            End If
        Next
        .Outline (0)
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False
        If .Enabled And .Visible Then .SetFocus
    End With
    
    LoadDetail = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Property Get GetChargeCollectID() As Long
    GetChargeCollectID = mlngCollectID
End Property

Public Property Get GetChargeCollectNO() As String
    GetChargeCollectNO = mstrCollectNO
End Property
 
Private Function CheckCancelValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查作废收款数据的合法性
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 15:44:57
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng收款ID As Long, strNO As String, blnDel As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, strPerson As String
    Dim int记录性质 As Integer
    On Error GoTo errHandle
    With vsCollect
        lng收款ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("收款单号")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
        strPerson = Trim(.TextMatrix(.Row, .ColIndex("收费员")))
        int记录性质 = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
    End With
    
    If blnDel Then
        MsgBox "收款单号为:" & strNO & "的收款记录已经被作废，不允许再作废!", vbInformation + vbOKOnly, gstrSysName
        If vsCollect.Enabled And vsCollect.Visible Then vsCollect.SetFocus
        Exit Function
    End If
    
    If int记录性质 = 4 Then
        '检查是否最后一次收款
        strSQL = "" & _
        "   Select 1 From 人员收缴记录  " & _
        "   Where 登记时间>(Select Max(登记时间) From 人员收缴记录 where ID=[1] ) " & _
        "               And ID+0<>[1] AND Rownum <2 And 记录性质=4 And 收款员=[2] and 作废时间 is null "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng收款ID, strPerson)
        If rsTemp.EOF = False Then
            MsgBox "注意:" & vbCrLf & _
                            "       收款单号为:" & strNO & "的收款记录不是你最后一次" & _
                           "的收款记录,为了保证收款数据正确，你必须从最后一次收款记录开始作废!", vbInformation + vbOKOnly, gstrSysName
            If vsCollect.Enabled And vsCollect.Visible Then vsCollect.SetFocus
            Exit Function
        End If
    End If
    CheckCancelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function CancelData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废当前收款数据
    '返回:作废成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 15:44:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng收款ID As Long, strNO As String, blnDel As Boolean
    Dim strDate As String, strSQL As String
    
    On Error GoTo errHandle
    With vsCollect
        If .Row < 1 Then Exit Function
        If .ColIndex("收款单号") < 0 _
            Or .ColIndex("ID") < 0 _
            Or .ColIndex("作废时间") < 0 Then Exit Function
        lng收款ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("收款单号")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
        mint记录性质 = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
        If strNO = "" Then Exit Function
    End With
    If CheckCancelValied = False Then Exit Function
    
    If MsgBox("你是否真的要将收款单号为:" & strNO & "进行作废吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    'Zl_财务收款记录_Cancel
    strSQL = "Zl_财务收款记录_Cancel("
    '  Id_In       In 人员收缴记录.Id%Type,
    strSQL = strSQL & "" & lng收款ID & ","
    ' 记录性质_In In 人员收缴记录.记录性质%Type,
    strSQL = strSQL & "" & mint记录性质 & ","
    '  作废人_In   In 人员收缴记录.作废人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  作废时间_In In 人员收缴记录.作废时间%Type
    strSQL = strSQL & "to_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    With vsCollect
        .TextMatrix(.Row, .ColIndex("作废人")) = UserInfo.姓名
        .TextMatrix(.Row, .ColIndex("作废时间")) = strDate
        .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
         mblnDel = True
    End With
    CancelData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String, blnPrintRollingCurtain As Boolean
    Dim rsTemp As ADODB.Recordset, lng收款ID As Long, strNO As String, blnDel As Boolean
    Err = 0: On Error GoTo ErrHand:
    blnPrintRollingCurtain = False
    If Val(vsCollect.Tag) = 0 Then
        '打印收款及票据汇总
        With vsCollect
            If .Row < 1 Then Exit Sub
            If .ColIndex("收款单号") < 0 _
                  Or .ColIndex("ID") < 0 _
                  Or .ColIndex("作废时间") < 0 Then Exit Sub
              lng收款ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
              strNO = Val(.TextMatrix(.Row, .ColIndex("收款单号")))
              blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
              If lng收款ID = 0 Then Exit Sub
        End With
        If Val(tbPage.Selected.Tag) = mPgIndex.EM_PG_收款及票据 Then
            Call mobjChargeBill.zlPrint(bytMode): Exit Sub
        End If
        '打印轧帐信息
        blnPrintRollingCurtain = True
    End If
    
    '输出收款信息
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr单位名称 & "财务收款清册"
    Set objRow = New zlTabAppRow
    If blnPrintRollingCurtain Then
        objRow.Add "收款单号：" & strNO
    Else
        If lblRange.Visible Then
            objRow.Add "时间范围：" & lblRange.Caption
        Else
            objRow.Add "时间范围：" & Format(dtpStartDate, "yyyy-mm-dd HH:MM:SS") & "至" & Format(dtpEndDate, "yyyy-mm-dd HH:MM:SS")
        End If
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = IIf(blnPrintRollingCurtain, vsRollingCurtain, vsCollect)
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打收款收据
    '编制:刘兴洪
    '日期:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, int记录性质 As Integer
    If Not (zlStr.IsHavePrivs(mstrPrivs, "收款收据打印") And zlStr.IsHavePrivs(mstrPrivs, "重打收款收据")) Then Exit Sub
    With vsCollect
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("收款单号")))
        int记录性质 = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
    End With
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500", Me, "NO=" & strNO, "记录性质=" & int记录性质, 2)
End Sub

Public Sub zlRefresh()
    '重新进行数据刷新
    Call cmdRefresh_Click
End Sub

Public Sub ShowChargeList(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细收款数据
    '编制:刘兴洪
    '日期:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng收款ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    
    With vsCollect
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("收款单号")))
        lng收款ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
    End With
    
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
   Call frmNew.ShowMe(frmMain, mlngModule, mstrPrivs, 4, CStr(lng收款ID), dtStartDate, dtEndDate, blnDel)
   If Not frmNew Is Nothing Then Unload frmNew
   Set frmNew = Nothing
End Sub

Public Property Get IsAllowViewChargeList() As Boolean
    '是否允许查看明细
    Dim int记录性质 As Integer, lngID As Long
    With vsCollect
        If .Row < 1 Then Exit Property
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Property
        int记录性质 = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
        IsAllowViewChargeList = int记录性质 = 4
    End With
End Property

Public Property Get IsAllowCollectCancel() As Boolean
    '是否允许收款作废
    Dim int记录性质 As Integer, lngID As Long
    With vsCollect
        If .Row < 1 Then Exit Property
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Property
        int记录性质 = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
        IsAllowCollectCancel = int记录性质 <> 6 And mblnDel = False
    End With
End Property

Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用自定义报表
    '入参:lngSys-系统号
    '        strRptCode-报表编号
    '编制:刘兴洪
    '日期:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDeptID As Long
    Dim lng收款ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    With vsCollect
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("收款单号")))
        lng收款ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
    End With
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "收款单号=" & strNO, _
        "收款ID=" & lng收款ID, _
        "作废标志=" & IIf(blnDel, 1, 0))
End Sub

Private Sub imgColPlanCollect_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanCollect.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + imgColPlanCollect.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCollect, lngLeft, lngTop, imgColPlanCollect.Height)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "收款信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlanCollect_Click()
    Call imgColPlanCollect_Click
End Sub

Private Sub imgColPlanRC_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanRC.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + imgColPlanRC.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlanRC.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlanRC_Click()
    Call imgColPlanRC_Click
End Sub




