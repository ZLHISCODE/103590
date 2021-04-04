VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeRollingCurtainHistory 
   BorderStyle     =   0  'None
   Caption         =   "收费员历史轧帐信息"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
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
   ScaleHeight     =   7365
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRollingCurtain 
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   285
      ScaleHeight     =   1980
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   1695
      Width           =   10170
      Begin VB.ComboBox cboDate 
         Height          =   330
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   127
         Width           =   1260
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "重新过滤数据(&R)"
         Height          =   350
         Left            =   7290
         TabIndex        =   5
         Top             =   105
         Width           =   1800
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   1800
         Left            =   165
         TabIndex        =   1
         Top             =   555
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeRollingCurtainHistory.frx":0000
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
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   9
            Top             =   60
            Width           =   210
            Begin VB.Image imgColPlan 
               Height          =   195
               Left            =   0
               Picture         =   "frmChargeRollingCurtainHistory.frx":00B5
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   2325
         TabIndex        =   2
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
         Format          =   116654083
         CurrentDate     =   41520
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   4785
         TabIndex        =   3
         Top             =   135
         Width           =   2430
         _ExtentX        =   4286
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
         Format          =   116654083
         CurrentDate     =   41520
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
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblHistoryDate 
         AutoSize        =   -1  'True
         Caption         =   "过滤时间"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   180
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
         TabIndex        =   4
         Top             =   195
         Width           =   225
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
Attribute VB_Name = "frmChargeRollingCurtainHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private Enum mPaneIndex
    EM_PN_Filter = 270101   '过滤条件
    EM_PN_RollingList = 270101  '轧帐列表
    EM_PN_ChargeBillTotal = 270102   '收款及票据汇总
End Enum
Private mblnNotBrush As Boolean '不刷新数据
Private mobjChargeBill As clsChargeBill
Private mlngRollingCurtainID As Long '轧帐ID
Private mstrRollingCurtainNO As String   '轧帐单据号
Private mblnDel As Boolean   '是否已轧帐作废
Private mfrmMain As Object

Public Sub zlInitVar(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '编制:刘兴洪
    '日期:2013-09-09 14:41:46
    '说明:加载窗体后,立即调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mlngModule = lngModule: mstrPrivs = strPrivs
End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngHeight As Long
    
    With dkpMan
        'Set .ImageList = zlCommFun.GetPubIcons
        lngHeight = 3980 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(mPaneIndex.EM_PN_RollingList, 400, lngHeight, DockRightOf, Nothing)
        objPane.Title = "轧帐信息"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picRollingCurtain.hWnd
        objPane.MinTrackSize.Height = lngHeight * 0.6
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_Filter, 400, 400, DockBottomOf, objPane)
        objPane.Title = "收款及票据汇总": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = mobjChargeBill.GetChargeAndBillTotalForm.hWnd
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
        .ColKey(0) = "过滤"
        .ColData(0) = "-1|1"
       For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Or _
               .ColKey(i) = "作废人" Or .ColKey(i) = "作废时间" Then
                .ColData(i) = "1|0"
            End If
            If .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            If .ColKey(i) = "收款部门" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Or .ColKey(i) = "轧帐类别" Then
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
    '功能:加载历史轧帐数据
    '返回:数据加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 17:08:50
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, i As Long, blnDel As Boolean
    Dim dtStartDate As Date, dtEndDate As Date, strTemp As String
    Dim strValue As String
    On Error GoTo errHandle
    Call GetDateRange(dtStartDate, dtEndDate)
    If dtpEndDate - dtStartDate > 90 Then
        '以收费员作为统计条件
        strWhere = " And a.收款员 = [1] And a.登记时间+0 Between [2] And [3] "
    Else
        '以时间作为统计条件
        strWhere = " And a.收款员||'' = [1] And a.登记时间 Between [2] And [3] "
    End If
    
    strSQL = "" & _
    "   Select /*+ rule */a.Id,a.No As 轧帐单号, a.开始时间, a.终止时间, a.登记人 As 轧帐人, a.登记时间 As 轧帐时间,  " & _
    "         b.名称 As 收款部门, a.摘要 As 轧帐说明, " & _
    "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
    "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
    "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计, " & _
    "         a.小组收款人, To_Char(a.小组收款时间, 'yyyy-mm-dd hh24:mi:ss') As 小组收款时间,  " & _
    "         a.财务收款人,To_Char(a.财务收款时间, 'yyyy-mm-dd hh24:mi:ss') As 财务收款时间,  " & _
    "         a.作废人, To_Char(a.作废时间, 'yyyy-mm-dd hh24:mi:ss') As 作废时间, " & _
    "         a.是否挂号,a.是否就诊卡,a.是否消费卡,a.是否收费,a.预交类别 As 是否预交,a.是否结帐 " & _
    "  From 人员收缴记录 A, 部门表 B " & _
    "  Where a.收款部门id = b.Id(+) And a.记录性质 = 1 " & strWhere & _
    "  Order by 登记时间 desc,轧帐单号 desc,小组收款时间 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, dtStartDate, dtEndDate)
    With vsRollingCurtain
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .FixedRows = 1
        Do While Not rsTemp.EOF
            strTemp = ""
            strValue = ""
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = Nvl(rsTemp!轧帐单号)
            '0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡
            If Val(Nvl(rsTemp!是否挂号)) = 1 Then
                strTemp = ",挂号"
                strValue = ",4"
            End If
            If Val(Nvl(rsTemp!是否就诊卡)) = 1 Then
                strTemp = strTemp & ",就诊卡"
                strValue = strValue & ",5"
            End If
            If Val(Nvl(rsTemp!是否消费卡)) = 1 Then
                strTemp = strTemp & ",消费卡"
                strValue = strValue & ",6"
            End If
            If Val(Nvl(rsTemp!是否收费)) = 1 Then
                strTemp = strTemp & ",收费"
                strValue = strValue & ",1"
            End If
            If Val(Nvl(rsTemp!是否预交)) = 1 Then
                strTemp = strTemp & ",预交"
                strValue = strValue & ",2"
            ElseIf Val(Nvl(rsTemp!是否预交)) = 2 Then
                strTemp = strTemp & ",门诊预交"
                strValue = strValue & ",21"
            ElseIf Val(Nvl(rsTemp!是否预交)) = 3 Then
                strTemp = strTemp & ",住院预交"
                strValue = strValue & ",22"
            End If
            If Val(Nvl(rsTemp!是否结帐)) = 1 Then
                strTemp = strTemp & ",结帐"
                strValue = strValue & ",3"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("轧帐类别")) = Mid(strTemp, 2)
            .Cell(flexcpData, .Rows - 1, .ColIndex("轧帐类别")) = Mid(strValue, 2)
            .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = Format(Nvl(rsTemp!开始时间), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = Format(Nvl(rsTemp!终止时间), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("轧帐人")) = Nvl(rsTemp!轧帐人)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = Format(Nvl(rsTemp!轧帐时间), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("轧帐说明")) = Nvl(rsTemp!轧帐说明)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Nvl(rsTemp!冲预交款)
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Nvl(rsTemp!借入合计)
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Nvl(rsTemp!借出合计)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款人")) = Nvl(rsTemp!小组收款人)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款时间")) = Nvl(rsTemp!小组收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("财务收款人")) = Nvl(rsTemp!财务收款人)
            .TextMatrix(.Rows - 1, .ColIndex("财务收款时间")) = Nvl(rsTemp!财务收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("作废人")) = Nvl(rsTemp!作废人)
            .TextMatrix(.Rows - 1, .ColIndex("作废时间")) = Nvl(rsTemp!作废时间)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then .Rows = .Rows - 1
'        If rsTemp.RecordCount <> 0 Then
'            Set .DataSource = rsTemp
'        End If
        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
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
               
                .RowData(i) = 0
            ElseIf Trim(.TextMatrix(i, .ColIndex("小组收款时间"))) = "" _
                    And Trim(.TextMatrix(i, .ColIndex("财务收款时间"))) = "" Then
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H80000018
                    .RowData(i) = 1 '表示未收款, &H80000018
            Else
                .RowData(i) = 0
            End If
        Next
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False
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
    mlngRollingCurtainID = 0
    mstrRollingCurtainNO = ""
    mblnDel = False
    
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
    '入参:
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
    Call mfrmMain.RefreshBasic
    Call LoadHistoryData
End Sub

Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub

Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Load()
    Set mobjChargeBill = New clsChargeBill
    Call InitPanel
    Call InitFace
End Sub
Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    Set mobjChargeBill = Nothing
End Sub
Private Sub picRollingCurtain_Resize()
        Err = 0: On Error Resume Next
        With picRollingCurtain
            vsRollingCurtain.Left = .ScaleLeft
            vsRollingCurtain.Top = cboDate.Top + cboDate.Height + 100
            vsRollingCurtain.Height = .ScaleHeight - vsRollingCurtain.Top - 50
            vsRollingCurtain.Width = .ScaleWidth
        End With
End Sub

Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
    vsRollingCurtain.Tag = "1"
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
    If OldRow >= 1 And OldRow <= vsRollingCurtain.Rows - 1 Then
        With vsRollingCurtain
            If Trim(.TextMatrix(OldRow, .ColIndex("小组收款时间"))) = "" And _
                Trim(.TextMatrix(OldRow, .ColIndex("财务收款时间"))) = "" And Trim(.TextMatrix(OldRow, .ColIndex("作废时间"))) = "" Then
              .Cell(flexcpBackColor, OldRow, 1, OldRow, .Cols - 1) = &H80000018
        End If
        End With
    End If
    If OldRow = NewRow Then Exit Sub
    '不刷新数据，退出
    If mblnNotBrush = True Then Exit Sub
    Call LoadDetail
End Sub
Private Sub vsRollingCurtain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Function LoadDetail() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载明细数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 11:17:09
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date
    Dim lng轧帐ID As Long, strNO As String, blnDel As Boolean
  
    With vsRollingCurtain
        If .Row >= 1 Then
            lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strNO = Val(.TextMatrix(.Row, .ColIndex("轧帐单号")))
            blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
        End If
    End With
    mlngRollingCurtainID = lng轧帐ID
    mstrRollingCurtainNO = strNO
    mblnDel = blnDel
    
    On Error GoTo errHandle
    dtStartDate = CDate("1991-01-01"): dtEndDate = dtStartDate
    If lng轧帐ID = 0 Then
        mobjChargeBill.ClearChargeAndBillTotalForm
    Else
        If mobjChargeBill.LoadChargeAndBillTotalData(Me, mlngModule, mstrPrivs, 1, lng轧帐ID, dtStartDate, dtEndDate, True, blnDel) = False Then Exit Function
    End If
    LoadDetail = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get GetChargeRollingCurtainID() As Long
    GetChargeRollingCurtainID = mlngRollingCurtainID
End Property
Public Property Get GetChargeRollingCurtainNO() As String
    GetChargeRollingCurtainNO = mstrRollingCurtainNO
End Property
Public Property Get GetChargeRollingCurtainDel() As Boolean
    GetChargeRollingCurtainDel = mblnDel
End Property
Private Function CheckCancelValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查作废轧帐数据的合法性
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 15:44:57
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng轧帐ID As Long, strNO As String, blnDel As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str轧帐类别 As String, strWhere As String
    Dim strRollingType As String
    On Error GoTo errHandle
    With vsRollingCurtain
        lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Val(.TextMatrix(.Row, .ColIndex("轧帐单号")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
        str轧帐类别 = "," & .Cell(flexcpData, .Row, .ColIndex("轧帐类别")) & ","
    End With
    
    If blnDel Then
        MsgBox "轧帐单号为:" & strNO & "的轧帐记录已经被作废，不允许再作废!", vbInformation + vbOKOnly, gstrSysName
        If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
        Exit Function
    End If
    
    '检查是否已经被收款
    With vsRollingCurtain
        If .TextMatrix(.Row, .ColIndex("小组收款时间")) <> "" Then
            MsgBox "轧帐单号为:" & strNO & "的轧帐记录已经小组收款，不允许作废!", vbInformation + vbOKOnly, gstrSysName
            If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
            Exit Function
        End If
        If .TextMatrix(.Row, .ColIndex("财务收款时间")) <> "" Then
            MsgBox "轧帐单号为:" & strNO & "的轧帐记录已经财务收款，不允许作废!", vbInformation + vbOKOnly, gstrSysName
            If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
            Exit Function
        End If
    End With
    '检查是否最后一次轧帐
    If InStr(str轧帐类别, ",1,") > 0 Then
        strWhere = "是否收费 = 1"
    End If
    If InStr(str轧帐类别, ",2,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "预交类别 = 1"
    End If
    If InStr(str轧帐类别, ",21,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "预交类别 = 2"
    End If
    If InStr(str轧帐类别, ",22,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "预交类别 = 3"
    End If
    If InStr(str轧帐类别, ",3,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "是否结帐 = 1"
    End If
    If InStr(str轧帐类别, ",4,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "是否挂号 = 1"
    End If
    If InStr(str轧帐类别, ",5,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "是否就诊卡 = 1"
    End If
    If InStr(str轧帐类别, ",6,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "是否消费卡 = 1"
    End If
    If strWhere <> "" Then strWhere = " And (" & strWhere & ")"
    strSQL = "" & _
    "   Select Max(NO) as NO From 人员收缴记录  " & _
    "   Where 登记时间>(Select Max(登记时间) From 人员收缴记录 where ID=[1] and 记录性质=1 ) " & _
    "               And ID+0 <> [1] and 记录性质=1  and 收款员=[2]  And 作废时间 Is Null " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng轧帐ID, UserInfo.姓名)
    
    If rsTemp.EOF = False Then
        If Nvl(rsTemp!NO) <> "" Then
        MsgBox "注意: " & vbCrLf & _
                        "     轧帐单号为:" & strNO & "的轧帐记录不是你最后一次的轧帐记录," & vbCrLf & _
                       "为了保证轧帐数据正确，你必须从最后一次轧帐记录[" & Nvl(rsTemp!NO) & "]开始作废!", vbInformation + vbOKOnly, gstrSysName
        If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
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
    '功能:作废当前轧帐数据
    '返回:作废成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 15:44:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng轧帐ID As Long, strNO As String, blnDel As Boolean
    Dim strDate As String, strSQL As String
    
    On Error GoTo errHandle
    With vsRollingCurtain
        If .Row < 1 Then Exit Function
        If .ColIndex("小组收款时间") < 0 _
            Or .ColIndex("轧帐单号") < 0 _
            Or .ColIndex("ID") < 0 _
            Or .ColIndex("作废时间") < 0 _
            Or .ColIndex("财务收款时间") < 0 Then Exit Function
        lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Val(.TextMatrix(.Row, .ColIndex("轧帐单号")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
    End With
    If MsgBox("你是否真的要将轧帐单号为:" & strNO & "进行作废吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If CheckCancelValied = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    'Zl_收费员轧帐记录_Cancel
    strSQL = "Zl_收费员轧帐记录_Cancel("
    '  Id_In       In 人员收缴记录.Id%Type,
    strSQL = strSQL & "" & lng轧帐ID & ","
    '  作废人_In   In 人员收缴记录.作废人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  作废时间_In In 人员收缴记录.作废时间%Type
    strSQL = strSQL & "to_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    With vsRollingCurtain
        .TextMatrix(.Row, .ColIndex("作废人")) = UserInfo.姓名
        .TextMatrix(.Row, .ColIndex("作废时间")) = strDate
        .Cell(flexcpForeColor, .Row, 1, .Row, .Cols - 1) = vbRed
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
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    If Val(vsRollingCurtain.Tag) = 0 Then
        '打印收款及票据汇总
        With vsRollingCurtain
            If .Row < 1 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("轧帐单号")) = "" Then Exit Sub
        End With
        Call mobjChargeBill.zlPrint(bytMode): Exit Sub
    End If
    '输出轧帐信息
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr单位名称 & "收费员轧帐清册"
    Set objRow = New zlTabAppRow
    If lblRange.Visible Then
        objRow.Add "时间范围：" & lblRange.Caption
    Else
        objRow.Add "时间范围：" & Format(dtpStartDate, "yyyy-mm-dd HH:MM:SS") & "至" & Format(dtpEndDate, "yyyy-mm-dd HH:MM:SS")
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsRollingCurtain
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
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Public Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打缴款书
    '编制:刘兴洪
    '日期:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String
    If Not (zlStr.IsHavePrivs(mstrPrivs, "缴款书打印") And zlStr.IsHavePrivs(mstrPrivs, "重打缴款书")) Then Exit Sub
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("轧帐单号")))
    End With
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1506", Me, "NO=" & strNO, 2)
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
    Dim lng轧帐ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("轧帐单号")))
        lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
    End With
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
   Call frmNew.ShowMe(frmMain, mlngModule, mstrPrivs, 1, CStr(lng轧帐ID), dtStartDate, dtEndDate, blnDel)
   If Not frmNew Is Nothing Then Unload frmNew
   Set frmNew = Nothing
    
End Sub
Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用自定义报表
    '入参:lngSys-系统号
    '        strRptCode-报表编号
    '编制:刘兴洪
    '日期:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDeptID As Long
    Dim lng轧帐ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("轧帐单号")))
        lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("作废时间"))) <> ""
    End With
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "NO=" & strNO, _
        "轧帐ID=" & lng轧帐ID, _
        "作废标志=" & IIf(blnDel, 1, 0))
End Sub
Public Sub zlDefaultSetFocus()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省输入
    '编制:刘兴洪
    '日期:2013-10-16 14:23:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If cboDate.Enabled And cboDate.Visible Then
        cboDate.SetFocus
    ElseIf dtpStartDate.Enabled And dtpStartDate.Visible Then
        dtpStartDate.SetFocus
    ElseIf vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then
        vsRollingCurtain.SetFocus
    End If
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
