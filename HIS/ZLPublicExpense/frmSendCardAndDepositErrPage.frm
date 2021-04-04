VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSendCardAndDepositErrPage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picEndDate 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   7635
      ScaleHeight     =   330
      ScaleWidth      =   1320
      TabIndex        =   4
      Top             =   3675
      Width           =   1320
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   148832259
         CurrentDate     =   40777
      End
   End
   Begin VB.PictureBox picStartDate 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5925
      ScaleHeight     =   300
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   3645
      Width           =   1425
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   148832259
         CurrentDate     =   40777
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsErrList 
      Height          =   3615
      Left            =   765
      TabIndex        =   0
      Top             =   1515
      Width           =   4545
      _cx             =   8017
      _cy             =   6376
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSendCardAndDepositErrPage.frx":0000
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
      Begin VB.PictureBox picErrImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   45
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   60
         Width           =   210
         Begin VB.Image imgErrImg 
            Height          =   195
            Left            =   0
            Picture         =   "frmSendCardAndDepositErrPage.frx":0059
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsthis 
      Left            =   570
      Top             =   690
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSendCardAndDepositErrPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------
'入口变量
Private mlngModule As Long
Private mblnNotRefresh As Boolean
Private mfrmMain As Object
Private mint应用场景 As Integer
'----------------------------------------------------------------------
'2.菜单相关变量
Private mblnNotChange As Boolean
Private mlngPreID As Long   '上次选择的异常ID
Private mobjCombox As CommandBarComboBox  '下拉列表
Private mobjDateLable As CommandBarControl  '日期控件
Private Const conMenu_Combox = 3820   '下拉框
Private Const conMenu_StartDate = 3824    '开始日期
Private Const conMenu_EndDate = 3825    '终止日期
Private Const conMenu_LableRange = 3827    '至
Private Const conMenu_LableDate = 3826    '终止日期
Private mintDateType As Integer
Private mlng异常ID As Long
Private mbln读其他操作员 As Boolean

'----------------------------------------------------------
'接口:
'  1.zlRefreshData-重新刷新数据
'  2.zlInit-初始化接口

'----------------------------------------------------------
Public Sub zlRefreshData()
    Call LoadErrDataToGrid
    
End Sub

Public Function zlInit(ByVal frmMain As Object, ByVal int应用场景 As Integer, ByVal lngModule As Long, Optional bln读其他操作员 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化功能
    '入参:int应用场景-1-医疗卡发卡;2-病人信息登记;3-病人入院 登记;4-预约挂号接收
    '    bln读其他操作员-允许读取其他操作员的异常单据
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 17:16:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mint应用场景 = int应用场景: Set mfrmMain = frmMain: mlngModule = lngModule
    zlInit = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Private Sub Form_Load()
    Call zlDefCommandBars
End Sub
 Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    vsErrList.Top = Top
    vsErrList.Left = Left: vsErrList.Width = Right - Left
    vsErrList.Height = Bottom - Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
End Sub

Private Sub vsErrList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModule, vsErrList, Me.Name, "异常列表", False
End Sub
 Private Sub vsErrList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Then Exit Sub
    zl_VsGridRowChange vsErrList, OldRow, NewRow, OldCol, NewCol, GRD_GOTFOCUS_COLORSEL
    If OldRow = NewRow Then Exit Sub
    With vsErrList
        mlng异常ID = 0
        If .Row < 0 Or .ColIndex("ID") < 0 Then Exit Sub
        mlng异常ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Sub
   
Private Sub vsErrList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsErrList
        Select Case Col
        Case .ColIndex("标志")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsErrList_GotFocus()
    zl_VsGridGotFocus vsErrList, GRD_GOTFOCUS_COLORSEL
End Sub

Private Sub vsErrList_LostFocus()
    zl_VsGridLostFocus vsErrList, GRD_LOSTFOCUS_COLORSEL
End Sub

Private Sub vsErrList_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModule, vsErrList, Me.Name, "异常列表", False
End Sub
Private Function ExcuteErrOper(Optional ByVal bln作废 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行异常操作
    '入参:bln作废-是否作废异常
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 16:37:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, lng病人ID As Long, lng异常ID As Long, bln作废异常 As Boolean
    Dim bln允许作废 As Boolean
    On Error GoTo errHandle
    
    With vsErrList
        lng异常ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        bln作废异常 = Trim(.TextMatrix(.Row, .ColIndex("异常方式"))) = "作废异常"
        bln允许作废 = Val(.TextMatrix(.Row, .ColIndex("同步状态"))) < 2
        If lng异常ID = 0 Then Exit Function
    End With
    If bln作废 And Not bln允许作废 Then
        MsgBox "当前异常记录已调用接口或已产生费用，不能作废，请点【异常重收】操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If bln作废异常 Then
       If Not bln作废 Then
            MsgBox "当前异常记录为作废异常记录，不能重收，请点【异常作废】操作!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
       End If
       If MsgBox("你是否真的要作废当前的异常数据么?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
       If Excute_ReCancel(lng异常ID) = False Then Exit Function
       MsgBox "作废成功", vbInformation + vbOKOnly, gstrSysName
       ExcuteErrOper = True
       Call LoadErrDataToGrid
       Exit Function
    End If
    
    'int操作类型-0-增加;1-异常重收;2-异常作废
    If frmSendCardAndDepositErrEdit.zlShowWindow(Me, IIf(bln作废, 2, 1), lng异常ID, mlngModule) = False Then Exit Function
    ExcuteErrOper = True
    Call LoadErrDataToGrid
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Excute_ReCancel(ByVal lng异常ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行异常重退操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-12-02 15:16:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllErrData As Collection, strSql As String, rsTemp As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim objService As clsService
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    strSql = "" & _
    "Select ID, 操作场景,nvl( 是否作废,0) as 是否作废, 业务id, 是否病历费, 病人id, 主页id, 预交单号, 医疗卡单号, 卡类别id, 发卡卡号, 同步状态, 交易信息, 登记时间, 操作员姓名 " & _
    "     From 病人结算异常记录 " & _
    "     Where ID =[1] "
    
    'int场景:1-医疗卡发卡;2-病人信息登记;3-病人入院 登记;4-预约挂号接收
    Set rsTemp = zlDatabase.OpenSQLRecordLob(strSql, Me.Caption, lng异常ID)
    If rsTemp.EOF Then
        MsgBox "读取异常数据失败，可能因并发原因被他人重收或作废，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Val(Nvl(rsTemp!是否作废)) <> 1 Then
        MsgBox "当前异常记录不是异常的作废记录，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllErrData = New Collection
    cllErrData.Add Array("异常ID", Val(Nvl(rsTemp!ID)))
    cllErrData.Add Array("操作场景", Val(Nvl(rsTemp!操作场景)))
    cllErrData.Add Array("业务ID", Val(Nvl(rsTemp!业务ID)))
     
    '删除医疗卡变动记录
    gcnOracle.BeginTrans: blnTrans = True
    If Zl_病人结算异常记录_Modify(2, cllErrData) = False Then
         gcnOracle.RollbackTrans: blnTrans = False: Exit Function
    End If
    
    If objService.zl_PatiSvr_DelCardChangeInfo(Val(Nvl(rsTemp!病人ID)), Val(Nvl(rsTemp!业务ID)), Val(Nvl(rsTemp!卡类别ID)), Trim(Nvl(rsTemp!发卡卡号))) = False Then
       gcnOracle.RollbackTrans: blnTrans = False: Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    Call LoadErrDataToGrid
    Excute_ReCancel = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsErrList_DblClick()
    Call ExcuteErrOper
End Sub
Private Sub imgErrImg_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    vRect = zlControl.GetControlRect(picErrImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picErrImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsErrList, lngLeft, lngTop, imgErrImg.Height)
    zl_vsGrid_Para_Save mlngModule, vsErrList, Me.Name, "异常列表", False
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objControl As CommandBarControl, cbrToolBar As CommandBar
    Dim objCustomControl As CommandBarControlCustom
    
    Err = 0: On Error GoTo errHand:
    
    Set cbsthis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsthis.VisualTheme = xtpThemeOffice2003
    With cbsthis.Options
        
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 16, 16
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
        
    End With
    cbsthis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsthis.DeleteAll
    Set cbrToolBar = cbsthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    
    cbrToolBar.EnableDocking xtpFlagStretched
    
    
    With cbrToolBar.Controls
        Set mobjCombox = .Add(xtpControlComboBox, conMenu_Combox, "缺省显示")
        mobjCombox.Width = 2600 / Screen.TwipsPerPixelX
        mobjCombox.Style = xtpComboLabel
        
        Set mobjDateLable = .Add(xtpControlLabel, conMenu_LableDate, "2017-01-01 23:59:59-2017-02-02 23:59:59")
        mobjDateLable.Visible = False
 
         Set objCustomControl = .Add(xtpControlCustom, conMenu_StartDate, "")
        objCustomControl.Handle = picStartDate.hWnd
        
        Set objControl = .Add(xtpControlLabel, conMenu_LableRange, " ～ ")
        Set objCustomControl = .Add(xtpControlCustom, conMenu_EndDate, "")
        objCustomControl.Handle = picEndDate.hWnd
        

        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): objControl.BeginGroup = True
        objControl.Flags = xtpFlagRightAlign
     
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "异常重收(&R)")
        objControl.Flags = xtpFlagRightAlign: objControl.IconId = 231
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "异常作废(&C)")
        objControl.Flags = xtpFlagRightAlign
    End With
    
    For Each objControl In cbrToolBar.Controls
        If objControl.Type <> xtpControlLabel And objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlComboBox Then
          objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    Call InitErrDate
    Call LoadErrDataToGrid
    zlDefCommandBars = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnOk As Boolean, objCombox As CommandBarComboBox
    On Error GoTo errHandle
    Select Case Control.ID
    Case conMenu_View_Refresh   '刷新
        Call LoadErrDataToGrid
    Case conMenu_Edit_ErrReBalance  '异常重收
        ExcuteErrOper False
    Case conMenu_Edit_ErrCancelBalance  '异常作废
        ExcuteErrOper True
    Case conMenu_Combox
        Set objCombox = Control
        mintDateType = objCombox.ListIndex
        Call ChargeComboxDate(mintDateType)
        Call LoadErrDataToGrid
        
    Case Else
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function ChargeComboxDate(ByVal intListIndex As Integer)
    Dim dtStartDate As Date, dtEndDate As Date
    
    intListIndex = intListIndex - 1
    Select Case intListIndex
        Case 0 '所有异常
        Case 1 '今日
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
        Case 2 '最近2天
            dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 3 '最近3天
            dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 4  '最近一周
            dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 5  '本月
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm") & "-01 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case Else
            dtStartDate = CDate(Format(dtpStartDate.value, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpEndDate.value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    If Not mobjDateLable Is Nothing Then
        mobjDateLable.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS") & "~" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
    End If
End Function

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Refresh   '刷新
    Case conMenu_Edit_ErrReBalance  '异常重收
        Control.Enabled = True ' mlng异常ID <> 0
    Case conMenu_Edit_ErrCancelBalance  '异常作废
        Control.Enabled = mlng异常ID <> 0
    Case conMenu_Combox
    Case conMenu_StartDate
       Control.Visible = mintDateType = 7
    Case conMenu_EndDate
       Control.Visible = mintDateType = 7
    Case conMenu_LableRange
       Control.Visible = mintDateType = 7
    Case conMenu_LableDate
       Control.Visible = mintDateType <> 7 And mintDateType <> 1
    Case Else
    End Select
    Exit Sub
End Sub
Private Function LoadErrDataToGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载异常数据给网格
    '入参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-05 22:21:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date, lngID As Long, i As Long
    Dim strSql As String, strPreCardNO  As String
    Dim rsTemp As ADODB.Recordset
     
    On Error GoTo errHandle
    
     Call zlCommFun.ShowFlash("正在加载病人异常结算数据,请稍等...", Me)
    
    Select Case mintDateType - 1
    Case 0 '所有异常
        dtStartDate = CDate(Format("1900-01-01", "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format("3000-01-01", "yyyy-mm-dd") & " 23:59:59")
    Case 1 '今日
        dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
    Case 2 '前一天至今日
        dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case 3 '前二天至今日
        dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case 4  '前一周至今日
        dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case 5  '本月
        dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm") & "-01 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case Else
        dtStartDate = CDate(Format(dtpStartDate.value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEndDate.value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    
    
    strSql = " " & _
    "   Select '' as 标志,A.ID,decode(a.操作场景,1,'医疗卡发卡',2,'病人信息登记',3,'病人入院登记',4,'预约挂号接收','其他')as  操作场景, " & _
    "       decode(nvl(a.是否作废,0),1,'作废异常','收费异常') as 异常方式,a.业务ID,decode(nvl(a.是否病历费,0),0,'','√') as 病历费,a.病人id,a.主页id, " & _
    "       a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.预交单号,a.预交金额, " & _
    "       a.医疗卡单号,a.卡费金额,a.卡类别ID,a.卡类别名称,a.发卡卡号,a.操作员姓名,a.登记时间,a.同步状态,a.交易信息 " & _
    "   From 病人结算异常记录 A " & _
    "   Where  A.操作场景=[1] And A.登记时间 between [2] and [3] " & IIf(mbln读其他操作员, "", "And 操作员姓名=[4]") & _
    "   Order by Decode(操作员姓名,[4],1,0)"
    
    Set rsTemp = zlDatabase.OpenSQLRecordLob(strSql, Me.Caption, mint应用场景, dtStartDate, dtEndDate, UserInfo.姓名)
       
       
    With vsErrList
        If .Row > 0 And .ColIndex("ID") >= 0 Then
            mlngPreID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            If mlngPreID <> 0 And .ColIndex("ID") >= 0 Then
                lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            End If
        End If
        
        .Redraw = flexRDNone
        .Clear: .Rows = 2: .Cols = 1
        .Cell(flexcpForeColor, 1, .FixedCols - 1, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpText, 0, 0, .Rows - 1, .Cols - 1) = ""
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        .Row = 1

        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            ''ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            .ColData(i) = "0||0"  '不能选择
            If .ColKey(i) Like "*ID" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True: .ColWidth(i) = True
                .ColData(i) = "-1||1"    '不能选择
            ElseIf .ColKey(i) Like "*时间" Or .ColKey(i) Like "*日期" Or .ColKey(i) = "状态" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*金额" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .ColData(.ColIndex("发卡卡号")) = "1||0": .ColData(.ColIndex("标志")) = "-1||1"
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
        zl_vsGrid_Para_Restore mlngModule, vsErrList, Me.Name, "异常作废", False, True
        .ColWidth(.ColIndex("标志")) = 285
        .ColAlignment(.ColIndex("标志")) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
    End With
    Call vsErrList_AfterRowColChange(-1, 0, vsErrList.Row, 0)
    Call zlCommFun.StopFlash
    LoadErrDataToGrid = True = True
    Exit Function
errHandle:
    Call zlCommFun.StopFlash
    vsErrList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitErrDate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化异常查找日期
    '编制:刘兴洪
    '日期:2019-11-05 22:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strValue As String
    Call GetRegInFor(g私有模块, Me.Name, "异常单据查询", strValue)
    i = Val(strValue)
    With mobjCombox
        .Clear
        .AddItem "所有异常情况"
        .ListIndex = 1
        If i = 0 Then .ListIndex = 1
        .AddItem "今日"
        If i = 1 Then .ListIndex = 2
        .AddItem "最近两天"
        If i = 2 Then .ListIndex = 3
        .AddItem "最近三天"
        If i = 3 Then .ListIndex = 4
        .AddItem "最近一周"
        If i = 4 Then .ListIndex = 5
        .AddItem "本月"
        If i = 5 Then .ListIndex = 6
        .AddItem "自定义时间范围"
        
        dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpEndDate.MaxDate = dtpStartDate.MaxDate
        dtpEndDate.value = dtpEndDate.MaxDate
        dtpStartDate.value = DateAdd("d", -7, dtpEndDate.MaxDate)
        mintDateType = mobjCombox.ListIndex
    End With
End Sub
