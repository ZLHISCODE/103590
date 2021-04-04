VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatiCureCardConsumeMgr 
   BorderStyle     =   0  'None
   Caption         =   "预存管理"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   9225
      TabIndex        =   6
      Top             =   150
      Width           =   1100
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   210
      ScaleHeight     =   465
      ScaleWidth      =   9360
      TabIndex        =   9
      Top             =   180
      Width           =   9360
      Begin VB.Frame fraType 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5475
         TabIndex        =   11
         Top             =   90
         Width           =   3600
         Begin VB.OptionButton optType 
            Caption         =   "门诊预存"
            Height          =   180
            Index           =   0
            Left            =   -15
            TabIndex        =   3
            Top             =   0
            Width           =   1035
         End
         Begin VB.OptionButton optType 
            Caption         =   "住院预存"
            Height          =   180
            Index           =   1
            Left            =   1155
            TabIndex        =   4
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton optType 
            Caption         =   "门诊和住院"
            Height          =   180
            Index           =   2
            Left            =   2340
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   30
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103874563
         CurrentDate     =   40722
      End
      Begin MSComCtl2.DTPicker dtp结束日期 
         Height          =   315
         Left            =   3315
         TabIndex        =   2
         Top             =   30
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103874563
         CurrentDate     =   40722
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   0
         Left            =   3195
         TabIndex        =   10
         Top             =   75
         Width           =   180
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "时间范围(&E)"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   90
         Width           =   990
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2250
      Left            =   195
      TabIndex        =   7
      Top             =   1320
      Width           =   5745
      _cx             =   10134
      _cy             =   3969
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiCureCardConsumeMgr.frx":0000
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
      ExplorerBar     =   7
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   8
         Top             =   45
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmPatiCureCardConsumeMgr.frx":0107
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmPatiCureCardConsumeMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mstrCardNo As String, mlngCardTypeID As Long, mlng病人ID As Long
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '弹出菜单操作
Public Event AfterRowChange(ByVal vsGrid As VSFlexGrid) '弹出菜单操作

Public Function zlReLoadData(ByVal lng病人ID As Long, ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载数据
    '返回:加载成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-06-28 15:30:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrCardNo = strCardNo: mlngCardTypeID = lngCardTypeID: mlng病人ID = lng病人ID
    Err = 0: On Error GoTo ErrHand:
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2011-06-28 15:31:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("卡号")) = "1|1"
        .ColAlignment(.ColIndex("本期消费")) = flexAlignRightCenter
    End With
End Sub

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long, lngRow As Long
    Dim str类别 As String, dbl余额 As Double, strSQL As String
    Dim strTable As String, strHTable As String, strIF As String, blnDataMove As String
    
    str类别 = ""
    If optType(2).value Then
        str类别 = str类别 & "1,2"
        strIF = "  And Nvl(a.预交类别, 2) in (1,2)"
    ElseIf optType(1).value Then
        str类别 = str类别 & "2"
        strIF = "  And Nvl(a.预交类别, 2)=2"
    Else
        str类别 = str类别 & "1"
        strIF = "  And Nvl(a.预交类别, 2)=1"
    End If
    
    mblnHaveData = False
    Err = 0: On Error GoTo ErrHand:
    '98616:李南春,2016/7/26,统计预交使用情况，要排除冲预交为0 的记录。不直接用预交记录统计是因为历史数据没有升级
    '84389:李南春,2015/5/6,多单据一次结算后，不再以no区分费用冲预交金额，改用结账id
    '问题:50472
    strTable = "" & _
    " Select 病人id, 收款时间, 0 As 类型,结帐id, Nvl(金额, 0) As 金额, 0 As 冲预交 " & _
    "  From 病人预交记录 A " & _
    "  Where 收款时间 >= [2] And 记录性质 =1  And 病人id =[1] " & strIF & _
    "  Union All " & _
    "  Select a.病人id, b.收费时间 As 收款时间,2 As 类型, b.id as 结帐id, 0 As 金额, Nvl(冲预交, 0) As 冲预交 " & _
    "  From 病人预交记录 A, 病人结帐记录 B " & _
    "  Where b.收费时间 >= [2] And Mod(a.记录性质, 10) = 1  And a.结帐id = b.Id And a.病人id =[1] " & strIF & _
    "  Union All " & _
    "  Select 病人id, 收费时间, 1 As  类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交 " & _
    "  From (Select a.病人id, b.登记时间 As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额, Max(Nvl(a.冲预交, 0)) As 冲预交 " & _
    "         From 病人预交记录 A, 门诊费用记录 B " & _
    "         Where b.登记时间 >= [2] And Mod(a.记录性质, 10) =1 And Nvl(b.记帐费用, 0) = 0 And a.结帐id = b.结帐id And b.病人id =[1] And " & _
    "           b.记录性质 In (1, 4) And Nvl(a.冲预交, 0) <> 0 " & strIF & _
    "         Group By a.病人id, b.登记时间, a.No, b.结帐id) " & _
    "  Group By 病人id, 收费时间, 结帐id " & _
    "  Union All " & _
    "  Select 病人id, 收费时间, 1 as 类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交 " & _
    "  From (Select a.病人id, b.登记时间 As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额, Max(Nvl(a.冲预交, 0)) As 冲预交 " & _
    "         From 病人预交记录 A, 住院费用记录 B " & _
    "         Where b.登记时间 >= [2] And Mod(a.记录性质, 10) =1 And a.结帐id = b.结帐id And b.病人id =[1] And b.记录性质 = 5 And " & _
    "           Nvl(b.记帐费用, 0) = 0 And Nvl(a.冲预交, 0) <> 0 " & _
    "         Group By a.病人id, b.登记时间, a.No, b.结帐id) " & _
    "  Group By 病人id, 收费时间, 结帐id"
    blnDataMove = zlDatabase.DateMoved(Format(dtp开始日期.value, "yyyy-mm-dd"), , , Me.Caption)
    
    If blnDataMove Then
        strHTable = Replace(strTable, "病人预交记录", "H病人预交记录")
        strHTable = Replace(strHTable, "住院费用记录", "H住院费用记录")
        strHTable = Replace(strHTable, "门诊费用记录", "H门诊费用记录")
        strTable = strTable & " UNION ALL " & strHTable
    End If
    strSQL = " " & _
        "   Select /*+ RULE */ 类别,收款时间, 业务类型, Sum(期初余额) As 期初余额, Sum(本期充值) As 本期充值, Sum(本期消费) As 本期消费 " & _
        "   From (With 预交 As ( " & strTable & ")" & _
        "          Select  0 as  类别,'' As 收款时间, '期初' As 业务类型, Sum(Nvl(预交余额, 0)) As 期初余额, 0 As 本期充值, 0 As 本期消费 " & _
        "          From 病人余额  A" & _
        "          Where 病人id = [1] And 性质 = 1 " & Replace(strIF, "预交类别", "类型") & _
        "          Union All " & _
        "          Select 0 as 类别,'' As 收款时间, '期初' As 业务类型, -1 * Sum(Nvl(金额, 0)) + Sum(Nvl(冲预交, 0)) As 期初余额, 0,  0 As 本期消费 " & _
        "          From 预交 " & _
        "          Where  收款时间 >= [2] " & _
        "          Group By To_Char(收款时间, 'yyyy-mm-dd') " & _
        "          Union All " & _
        "          Select 1 as 类别,To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, '充值' As 业务类型, 0 As 期初余额, Sum(Nvl(金额, 0)) As 充值, 0 As 本期消费 " & _
        "          From 预交 " & _
        "          Where  收款时间 Between [2] And [3] " & _
        "          Having Sum(Nvl(金额, 0))<>0 " & _
        "          Group By To_Char(收款时间, 'yyyy-mm-dd')" & _
        "         Union All " & _
        "          Select 1 as 类别, To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, decode(类型,1,'收费',2,'结帐','消费') As 业务类型, 0 As 期初余额, 0 As 充值, " & _
        "                 Sum(Nvl(冲预交, 0)) As 消费 " & _
        "          From 预交 " & _
        "          Where  收款时间 Between [2] And [3] " & _
        "           Having Sum(Nvl(冲预交, 0))<>0  Group By To_Char(收款时间, 'yyyy-mm-dd') ,Decode(类型, 1, '收费', 2, '结帐', '消费')) " & _
        "          Group By  类别,收款时间, 业务类型" & _
        "          Order By 类别,收款时间 "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, dtp开始日期.value, dtp结束日期.value)
    With Me.vsGrid
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Clear 1
        .Row = 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("收款时间")) = Nvl(rsTemp!收款时间)
            .TextMatrix(lngRow, .ColIndex("业务类型")) = Nvl(rsTemp!业务类型)
            .TextMatrix(lngRow, .ColIndex("期初余额")) = Format(Val(Nvl(rsTemp!期初余额)), "####0.00;-###0.00; ;")
            .TextMatrix(lngRow, .ColIndex("本期充值")) = Format(Val(Nvl(rsTemp!本期充值)), "####0.00;-###0.00; ;")
            .TextMatrix(lngRow, .ColIndex("本期消费")) = Format(Val(Nvl(rsTemp!本期消费)), "####0.00;-###0.00; ;")
            dbl余额 = dbl余额 + Val(Nvl(rsTemp!期初余额)) + Val(Nvl(rsTemp!本期充值)) - Val(Nvl(rsTemp!本期消费))
            .TextMatrix(lngRow, .ColIndex("期未余额")) = Format(dbl余额, "####0.00;-###0.00;;")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        Call InitVsGrid
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "帐户入出列表", True
        .ColWidth(.ColIndex("标志")) = 285
        .ColAlignment(.ColIndex("标志")) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
    End With
    mblnHaveData = rsTemp.RecordCount > 0
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsGrid.Redraw = flexRDBuffered
End Sub

Private Sub cmdRefresh_Click()
    Call LoadDataToRpt
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    dtp结束日期.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtp结束日期.value = Format(dtp结束日期.MaxDate, "yyyy-mm-dd 23:59:59")
    dtp开始日期.MaxDate = dtp结束日期.MaxDate
    dtp开始日期.value = Format(DateAdd("m", -1, dtp开始日期.MaxDate), "yyyy-mm-dd 00:00:00")
    Call InitVsGrid
    Call vsGrid_GotFocus
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single
    Err = 0: On Error Resume Next
    If Me.ScaleWidth < 10455 Then
        fraType.Top = dtp结束日期.Top + dtp结束日期.Height + 120
        picFilter.Height = 445 + dtp结束日期.Height
        fraType.Left = dtp开始日期.Left
        cmdRefresh.Top = picFilter.Top + picFilter.Height - cmdRefresh.Height - 50
    Else
        fraType.Top = dtp结束日期.Top + (dtp结束日期.Height - fraType.Height) \ 2
        fraType.Left = dtp结束日期.Left + dtp结束日期.Width + 100
        picFilter.Height = 465
        cmdRefresh.Top = picFilter.Top + dtp结束日期.Top
    End If
    cmdRefresh.Left = Me.ScaleWidth - cmdRefresh.Width - 100
    With vsGrid
        .Left = ScaleLeft: .Top = picFilter.Top + picFilter.Height
        .Width = ScaleWidth: .Height = ScaleHeight - .Top
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "帐户入出列表", True
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "帐户入出列表", True
End Sub

 
Private Sub optType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub picImg_Click()
    Call imgCol_Click
End Sub
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2011-06-28 15:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset, vsGrid As VSFlexGrid
    Err = 0: On Error GoTo errH:
    gstrSQL = "Select   A.姓名,A.性别,A.年龄 From 病人信息 A where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID)
    If rsTemp.EOF = True Then Exit Sub '无卡信息，退出
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
        
    objPrint.Title.Text = gstrUnitName & "帐户入出情况"
    
    objRow.Add "姓名：" & Nvl(rsTemp!姓名)
    objRow.Add "年龄：" & Nvl(rsTemp!年龄)
    objRow.Add "性别：" & Nvl(rsTemp!性别)
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("标志") Then .ColWidth(intCol) = 0
        Next
    End With
    Set objPrint.Body = vsGrid
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "帐户入出列表", True
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow <> NewRow Then
        RaiseEvent AfterRowChange(vsGrid)
    End If
End Sub
Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "帐户入出列表", True
End Sub

 Private Sub vsGrid_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub
Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub
Private Sub vsGrid_LostFocus()
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(vsGrid)
End Sub
Private Sub dtp结束日期_Change()
     If dtp结束日期.value > dtp开始日期.MaxDate Then dtp结束日期.value = dtp开始日期.MaxDate
    If dtp结束日期.value < dtp开始日期.value Then
        dtp开始日期.value = dtp结束日期.value
    End If
End Sub
Private Sub dtp结束日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtp开始日期_Change()
    If dtp开始日期.value > dtp结束日期.MaxDate Then dtp开始日期.value = dtp结束日期.MaxDate
    If dtp结束日期.value < dtp开始日期.value Then
        dtp结束日期.value = dtp开始日期.value
    End If
End Sub
Private Sub dtp开始日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Public Function zlShowReport(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:弹出医疗收支清册报表
    '入参:lng病人ID 病人ID号
    '编制:王吉
    '日期:2012-06-12 15:59:59
    '问题50122
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str类别 As String
    Dim strDate As String
    
    strDate = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("收款时间"))
    
    str类别 = ""
    If optType(2).value Then
        str类别 = "3"
    ElseIf optType(1).value Then
        str类别 = "2"
    Else
        str类别 = "1"
    End If
     If vsGrid.Row >= 2 Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1107_2", Me, "病人ID=" & lng病人ID, "日期=" & CDate(strDate), "预交类别=" & str类别)
End Function
