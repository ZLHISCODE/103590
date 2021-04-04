VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFinanceSupervisePersonOthers 
   BorderStyle     =   0  'None
   Caption         =   "其他人员收款操作"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   285
      ScaleHeight     =   810
      ScaleWidth      =   11940
      TabIndex        =   10
      Top             =   645
      Width           =   11940
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "重新提取收款数据(&G)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7440
         TabIndex        =   4
         Top             =   0
         Width           =   2190
      End
      Begin VB.ComboBox cboDept 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   450
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.TextBox txtMemo 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   8
         Top             =   420
         Width           =   4590
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "收款(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8530
         TabIndex        =   9
         Top             =   405
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   0
         Width           =   2355
         _ExtentX        =   4154
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
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   168755203
         CurrentDate     =   41520
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Top             =   30
         Width           =   2355
         _ExtentX        =   4154
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
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   168755203
         CurrentDate     =   41520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   3125
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "收款部门"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   5
         Top             =   495
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "截止收款时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3750
         TabIndex        =   2
         Top             =   75
         Width           =   1260
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "上次收款时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   0
         Top             =   75
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收款说明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   7
         Top             =   525
         Width           =   840
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
Attribute VB_Name = "frmFinanceSupervisePersonOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPaneIndex
    EM_PN_Filter = 260101   '过滤条件
    EM_PN_ChargeBillTotal = 260102  '收款及票据汇总
End Enum
Private mobjChargeBill As clsChargeBill
Private mlngModule As Long, mstrPrivs As String
Private mstrPreDate As String, mstrPersonName As String
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mobjChargeBill.zlPrint(bytMode, "", txtMemo.Text)
End Sub
Public Sub zlInitVar(ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '出参:
    '编制:刘兴洪
    '日期:2013-09-09 14:41:46
    '说明:加载窗体后,立即调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    Call SetControlEanbled
    Call InitFace: Call SetPopedom
    Call mobjChargeBill.ClearChargeAndBillTotalForm
End Sub

Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面数据
    '编制:刘兴洪
    '日期:2013-09-11 14:05:08
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String, rsTemp As ADODB.Recordset
'    strSQL = "" & _
'    "   Select Distinct a.Id, a.编码, a.名称,b.缺省" & vbNewLine & _
'    "   From 部门表 a, 部门人员 b" & vbNewLine & _
'    "   Where a.Id = b.部门id And b.人员ID=[1] " & vbNewLine & _
'     "              And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
'    "               And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
'    "   Order By a.编码"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
'
'    With cboDept
'        .Clear
'        Do While Not rsTemp.EOF
'            .AddItem Nvl(rsTemp!编码) & "-" & rsTemp!名称
'            .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
'            If Val(Nvl(rsTemp!缺省)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
'            rsTemp.MoveNext
'        Loop
'        If .ListIndex < 0 And .ListCount <> 0 Then .ListIndex = 0
'    End With
    dtpEndDate.MaxDate = DateAdd("s", 1, zlDatabase.Currentdate)
    dtpEndDate.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
End Sub

Private Sub SetPopedom()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置权限控制
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 12:02:06
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdOK.Visible = zlStr.IsHavePrivs(mstrPrivs, "其他人员收款") And zlStr.IsHavePrivs(mstrPrivs, "轧帐收款")
End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngFilterHeight As Long, lngBillHeight As Long
    lngFilterHeight = 810 / Screen.TwipsPerPixelX
    lngBillHeight = 1000 / Screen.TwipsPerPixelX
    With dkpMan
        'Set .ImageList = zlCommFun.GetPubIcons
        Set objPane = .CreatePane(EM_PN_Filter, 100, lngFilterHeight, DockLeftOf, Nothing)
        objPane.Title = "过滤条件": objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngFilterHeight: objPane.MaxTrackSize.Height = lngFilterHeight
        objPane.Handle = picFilter.hWnd
        Set objPane = .CreatePane(EM_PN_ChargeBillTotal, 400, 400, DockBottomOf, objPane)
        objPane.Title = "收款及票据汇总"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
        objPane.Handle = mobjChargeBill.GetChargeAndBillTotalForm.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function
Private Function CheckValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据输入的合法性
    '返回:数据输入合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 11:45:04
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
'    If cboDept.ListIndex < 0 Then
'        MsgBox "注意:" & vbCrLf & "   未选择收款部门,请选择收款部门!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
'        If cboDept.Enabled And cboDept.Visible Then cboDept.SetFocus
'        Exit Function
'    End If
    If InStr(Trim(txtMemo.Text), "'") > 0 Then
        MsgBox "注意:" & vbCrLf & "   收款说明不允许有单引号!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If

    If zlCommFun.ActualLen(txtMemo.Text) > 50 Then
        MsgBox "注意:" & vbCrLf & "   收款说明最多只能输入50个字符或25个汉字,请重新输入!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    
    CheckValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub BillPrint(ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:轧帐票据打印
    '编制:刘兴洪
    '日期:2013-09-11 11:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    blnPrint = False
    If Not zlStr.IsHavePrivs(mstrPrivs, "收款收据打印") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("收款收据打印方式", glngSys, mlngModule))     '使用医生站的相关参数
    Case 0    '不打印
        Exit Sub
    Case 1    '自助动打印
        blnPrint = True
    Case 2    '选择打印
        If MsgBox("你是否要打印缴款书？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            blnPrint = True
        End If
    End Select
    If blnPrint = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500", Me, "NO=" & strNO, "记录性质=4", 2)
End Sub

Public Function SaveData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存轧帐数据
    '返回:收款数据保存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 11:39:42
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String
    Dim strStartDate As String, strEndDate As String, lngDeptID As Long
    Dim strMemo As String, blnOK As Boolean
    On Error GoTo errHandle
    If CheckValied = False Then Exit Function
    strStartDate = Format(dtpStartDate, "yyyy-mm-dd HH:MM:SS")
    strEndDate = Format(dtpEndDate, "yyyy-mm-dd HH:MM:SS")
    lngDeptID = 0
'   lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    strMemo = Trim(txtMemo.Text)
    blnOK = mobjChargeBill.GetChargeAndBillTotalForm.SaveData(strStartDate, strEndDate, strMemo, lngDeptID, strNO, lngID)
    'SaveData(ByVal strStartDate As String, ByVal strEndDate As String, _
    ByVal strMemo As String, ByVal lngDeptID As Long, _
    ByRef strNo As String, ByRef lngID As Long)
    If blnOK Then
        '票据打印
        dtpStartDate.Value = dtpEndDate.Value: dtpStartDate.Enabled = False
        dtpEndDate.MaxDate = DateAdd("s", 1, zlDatabase.Currentdate)
        dtpEndDate.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        Call BillPrint(strNO)
        '重新加载数据
        cmdRefresh_Click
    End If
    SaveData = blnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    If SaveData() = False Then Exit Sub
End Sub
Private Sub cmdRefresh_Click()
    If dtpEndDate - dtpStartDate > 14 Then
        If MsgBox("当前选择的操作员所设置的轧帐时间过长,提取轧帐数据可能需要一定的时间，你是否还要继续轧帐收款?", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    End If
    Call mobjChargeBill.LoadChargeAndBillTotalData(Me, mlngModule, mstrPrivs, 5, 0, dtpStartDate, dtpEndDate, _
        False, , IIf(mstrPersonName = "", "-", mstrPersonName), "0")
    If txtMemo.Enabled And txtMemo.Visible Then
        txtMemo.SetFocus: Exit Sub
    End If
    If cmdOK.Enabled And cmdOK.Visible Then
        cmdOK.SetFocus: Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Set mobjChargeBill = New clsChargeBill
    Call InitPanel
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mobjChargeBill = Nothing
    mstrPersonName = ""
End Sub
Private Sub picFilter_Resize()
    Err = 0: On Error Resume Next
    Line1.X2 = picFilter.Width
    With picFilter
        'cmdRefresh.Left = cmdOK.Left - cmdRefresh.Width - 50
        cmdOK.Left = .ScaleWidth - cmdOK.Width - 100
        If cmdOK.Left - txtMemo.Left - 50 < 1000 Then
            txtMemo.Width = 1000
        Else
            txtMemo.Width = cmdOK.Left - txtMemo.Left - 50
        End If
    End With
End Sub
Private Sub txtMemo_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtMemo
End Sub
Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub
Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
Private Sub txtMemo_LostFocus()
    zlCommFun.OpenIme False
End Sub
Public Property Get GetCashMoney() As Double
    '获取现金金额
    GetCashMoney = mobjChargeBill.GetChargeAndBillTotalForm.GetCashMoney
End Property
Public Sub zlRefresh()
    '重新进行数据刷新
    Call cmdRefresh_Click
End Sub

Public Function zlLoadPersonData(ByVal strPersonName As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载非收费员的数据
    '入参:strPersonName-指定的收费员
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-29 17:13:24
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrPersonName = strPersonName
    Call SetControlEanbled
    On Error GoTo errHandle
    Call SetPreRollingCurtainDate
    Call cmdRefresh_Click
    zlLoadPersonData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetPreRollingCurtainDate()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置上次轧帐时间(或收款时间)
    '编制:刘兴洪
    '日期:2013-09-29 17:18:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    mstrPreDate = ""
    dtpStartDate.MinDate = 0
    '取上次收款时间,先从余额表中取数
    strSQL = "Select Max(上次轧帐时间) as 上次轧帐时间 From 人员缴款余额 Where 收款员=[1] and 性质=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrPersonName)
    mstrPreDate = Format(rsTemp!上次轧帐时间, "yyyy-mm-dd HH:MM:SS")
    If mstrPreDate = "" Then
        '2.如果人员缴款余额.上次轧帐时间 为NULL,则按最后一次轧账的截止时间为该时间；
        strSQL = "" & _
        "   Select to_Char(Max(终止时间),'yyyy-mm-dd hh24:mi:ss') as 终止时间 " & _
        "   From 人员收缴记录  " & _
        "   Where 记录性质=1 and 作废时间 is null And 收款员=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrPersonName)
        If Not rsTemp.EOF Then
            mstrPreDate = Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS")
            If mstrPreDate <> "" Then
                dtpStartDate.Value = CDate(mstrPreDate)
                dtpStartDate.MinDate = dtpStartDate.Value
                dtpEndDate.MinDate = dtpStartDate.Value
                dtpStartDate.Enabled = False: Exit Sub
            End If
        End If
    End If
    '3.如果人员缴款余额.上次轧帐时间 为NULL,则按最后一次轧账的截止时间为该时间；
    If mstrPreDate = "" Then
        '3.如果未轧过的,缺省为领取备用金时间(初次上岗领用备用金的时间);
        strSQL = "" & _
        "   Select to_Char(min(登记时间),'yyyy-mm-dd hh24:mi:ss') as 登记时间 " & _
        "   From 人员暂存记录  " & _
        "   Where 记录性质=1 and 收回时间 is null And 收款员=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
        If Not rsTemp.EOF Then mstrPreDate = Nvl(rsTemp!登记时间)
    End If
    '获取上次轧帐时间
    dtpEndDate.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    dtpStartDate.Enabled = dtpStartDate.Enabled And mstrPreDate = ""  '上次轧帐时间为空时,需要手工选择确定时间
    If mstrPreDate <> "" Then
        dtpStartDate.Value = CDate(mstrPreDate)
        dtpStartDate.MinDate = dtpStartDate.Value
        dtpEndDate.MinDate = dtpStartDate.Value
    Else
        dtpStartDate.MinDate = CDate("1901-01-01")
        dtpEndDate.MinDate = CDate("1901-01-01")
    End If
    
    If mstrPreDate = "" Then mstrPreDate = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-mm-dd HH:MM:SS")
    dtpStartDate.Value = CDate(mstrPreDate)
    '可能存在12点轧帐,时间很快就过秒,所以要+1
    dtpEndDate.MaxDate = Format(DateAdd("s", 1, dtpEndDate.Value), "yyyy-mm-dd HH:MM:SS")
    dtpStartDate.MaxDate = dtpEndDate.MaxDate
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetControlEanbled()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Eanbled属性
    '编制:刘兴洪
    '日期:2013-09-29 17:50:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtpEndDate.Enabled = mstrPersonName <> ""
    dtpStartDate.Enabled = mstrPersonName <> ""
    cmdOK.Enabled = mstrPersonName <> ""
'    cboDept.Enabled = mstrPersonName <> ""
    txtMemo.Enabled = mstrPersonName <> ""
    cmdRefresh.Enabled = mstrPersonName <> ""
End Sub

Public Sub ShowChargeList(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细收款数据
    '入参:frmMain-调用的主窗体
    '编制:刘兴洪
    '日期:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
    Call frmNew.ShowMe(frmMain, mlngModule, mstrPrivs, 5, "", dtpStartDate.Value, dtpEndDate.Value, False, _
        mstrPersonName, "0")
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
'    Dim lngDeptID As Long
'    If cboDept.ListIndex >= 0 Then lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    Call ReportOpen(gcnOracle, strRptCode, frmMain, _
        "开始收款日期=" & Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM:SS"), _
        "终止收款日期=" & Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM:SS"))
End Sub

