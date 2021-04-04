VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeRollingCurtain 
   BorderStyle     =   0  'None
   Caption         =   "收费员轧帐"
   ClientHeight    =   9525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   285
      ScaleHeight     =   1215
      ScaleWidth      =   11940
      TabIndex        =   11
      Top             =   645
      Width           =   11940
      Begin VB.TextBox txtRemain 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   450
         Width           =   2355
      End
      Begin VB.TextBox txtHandIn 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   5
         Top             =   450
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
         TabIndex        =   7
         Top             =   825
         Width           =   4590
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "轧帐(&Z)"
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
         Left            =   9630
         TabIndex        =   10
         Top             =   815
         Width           =   1100
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
         TabIndex        =   9
         Top             =   450
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "重新提取数据(&R)"
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
         Left            =   7485
         TabIndex        =   4
         Top             =   15
         Width           =   1750
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   15
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
         Format          =   197263363
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
         Format          =   197263363
         CurrentDate     =   41520
      End
      Begin VB.Label lblRemain 
         AutoSize        =   -1  'True
         Caption         =   "暂存金"
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
         Left            =   4380
         TabIndex        =   14
         Top             =   510
         Width           =   630
      End
      Begin VB.Label lblHandIn 
         AutoSize        =   -1  'True
         Caption         =   "上缴金额"
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
         TabIndex        =   13
         Top             =   510
         Width           =   840
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "轧帐说明(&M)"
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
         Left            =   150
         TabIndex        =   12
         Top             =   885
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   1500
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "上次轧帐时间"
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
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "截止轧帐时间"
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
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "收款部门(&D)"
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
         Left            =   150
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1155
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
Attribute VB_Name = "frmChargeRollingCurtain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPaneIndex
    EM_PN_Filter = 260101   '过滤条件
    EM_PN_ChargeBillTotal = 260102  '收款及票据汇总
End Enum
Private mobjChargeBill As clsChargeBill, mfrmMain As Object
Private mlngModule As Long, mstrPrivs As String, mdatEnd As Date, mdatBegin As Date
Private mstrPreDate As String, mstrOperatorName As String, mblnNotClick As Boolean
Private mdblDefaultHandIn As Double
Private mstrRollingType As String    '类别
Private mblnChangeEndDate As Boolean
Private mblnNotChange As Boolean, mstrDefaultTime As String

Public Sub RefreshPage()
    Call cmdRefresh_Click
End Sub

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

Public Sub zlInitVar(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal strPreDate As String, ByVal strOperatorName As String, ByVal strRollingType As String, Optional ByVal strDefaultTime As String = "")
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:lngModule-模块号
    '     strPrivs-权限串
    '     strPreDate-上次轧帐时间
    '     strOperatorName-操作员姓名
    '     strRollingType-轧帐类别
    '出参:
    '编制:刘兴洪
    '日期:2013-09-09 14:41:46
    '说明:加载窗体后,立即调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mlngModule = lngModule: mstrPrivs = strPrivs: mstrOperatorName = strOperatorName
    mstrPreDate = strPreDate: mstrRollingType = strRollingType
    mstrDefaultTime = strDefaultTime
    Call InitFace: Call SetPopedom
End Sub
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面数据
    '编制:刘兴洪
    '日期:2013-09-11 14:05:08
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date
    
    '获取上次轧帐时间
    dtDate = zlDatabase.Currentdate
    
    dtpStartDate.Enabled = mstrPreDate = "" '上次轧帐时间为空时,需要手工选择确定时间
    If mstrPreDate = "" Then mstrPreDate = Format(DateAdd("d", -7, dtDate), "yyyy-mm-dd 00:00:00")
    dtpStartDate.MaxDate = dtDate
    dtpEndDate.MaxDate = dtDate
    dtpStartDate.Value = CDate(mstrPreDate)
    If mstrDefaultTime = "" Then
        dtpEndDate.Value = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    Else
        dtpEndDate.Value = Format(mstrDefaultTime, "yyyy-MM-dd hh:mm:ss")
        Call mobjChargeBill.LoadChargeAndBillTotalData(Me, mlngModule, mstrPrivs, 1, 0, dtpStartDate, dtpEndDate, False, , , mstrRollingType)
        If mobjChargeBill.ChargeBillHaveData = False Then
            dtpEndDate.Value = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
        End If
        Call mobjChargeBill.ClearChargeAndBillTotalForm
    End If
End Sub

Private Sub SetPopedom()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置权限控制
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 12:02:06
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdOK.Visible = zlStr.IsHavePrivs(mstrPrivs, "轧帐")
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngFilterHeight As Long, lngBillHeight As Long
    lngFilterHeight = 1215 / Screen.TwipsPerPixelX
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
        MsgBox "注意:" & vbCrLf & "   轧帐说明不允许有单引号!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    
    '问题号:110281,焦博,2017/08/11,把轧账说明的上限从50个字符调整为500个字符
    If zlCommFun.ActualLen(txtMemo.Text) > 500 Then
        MsgBox "注意:" & vbCrLf & "   轧帐说明最多只能输入500个字符或250个汉字,请重新输入!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
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
    If Not zlStr.IsHavePrivs(mstrPrivs, "缴款书打印") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("缴款书打印方式", glngSys, mlngModule))     '使用医生站的相关参数
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
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1506", Me, "NO=" & strNO, 2)
End Sub

Public Function SaveData(ByRef lngID As Long, ByRef strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存轧帐数据
    '出参:lngId-轧帐ID
    '       strNo-轧帐单号
    '返回:轧帐数据保存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 11:39:42
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strStartDate As String, strEndDate As String, lngDeptID As Long
    Dim strMemo As String, blnOK As Boolean
    Dim objChargeBillTotal As frmChargeBillTotal
    On Error GoTo errHandle
    If CheckValied = False Then Exit Function
    
    strStartDate = Format(dtpStartDate, "yyyy-mm-dd HH:MM:SS")
    strEndDate = Format(dtpEndDate, "yyyy-mm-dd HH:MM:SS")
    lngDeptID = 0
    strMemo = Trim(txtMemo.Text)
    Set objChargeBillTotal = mobjChargeBill.GetChargeAndBillTotalForm
    blnOK = objChargeBillTotal.SaveData(strStartDate, strEndDate, strMemo, lngDeptID, strNO, lngID, Val(txtRemain.Text), mstrRollingType)
    
    If blnOK Then
        '票据打印
        dtpStartDate.Value = dtpEndDate.Value: dtpStartDate.Enabled = False
        dtpEndDate.MaxDate = DateAdd("d", 1, zlDatabase.Currentdate)
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
    Dim lngID As Long, strNO As String, datTemp As Date
    datTemp = zlDatabase.Currentdate
    If datTemp < dtpEndDate Then
        MsgBox "轧帐结束时间超过了当前的系统时间(" & Format(datTemp, "yyyy-mm-dd hh:mm:ss") & "),不允许轧帐！", vbCritical, gstrSysName
        Exit Sub
    End If
    If mdatEnd <> dtpEndDate Then
        If MsgBox("当前的轧帐时间与提取数据的轧帐时间不一致，是否按照新的轧帐时间重新刷新数据？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            mblnNotClick = True
            Call cmdRefresh_Click
            mblnNotClick = False
            Exit Sub
        Else
            If mdatEnd > CDate("2000-01-01") Then
                dtpEndDate = mdatEnd
            End If
            Exit Sub
        End If
    End If
    
    If mblnChangeEndDate Then
        If MsgBox("截止时间发生了变化,请重新提取需要轧帐的数据,是否重提轧帐数据?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
           If cmdRefresh.Enabled And cmdRefresh.Visible Then Call cmdRefresh_Click
        Else
            If cmdRefresh.Enabled And cmdRefresh.Visible Then cmdRefresh.SetFocus
        End If
        Exit Sub
    End If
    If SaveData(lngID, strNO) = False Then Exit Sub
End Sub

Public Sub SaveDataWithCheck()
    Call cmdOK_Click
End Sub

Private Sub cmdRefresh_Click()
    Call mfrmMain.RefreshBasic
    Call mobjChargeBill.LoadChargeAndBillTotalData(Me, mlngModule, mstrPrivs, 1, 0, dtpStartDate, dtpEndDate, False, , , mstrRollingType)
    txtHandIn.Text = Format(mobjChargeBill.GetHandIn, "0.00")
    mdblDefaultHandIn = Val(txtHandIn.Text)
    txtRemain.Text = "0.00"
    mdatEnd = dtpEndDate: mdatBegin = dtpStartDate
    If mblnNotClick = False Then zlCommFun.PressKey vbKeyTab
    mblnChangeEndDate = False
    
End Sub

Private Sub dtpEndDate_Change()
    If mblnNotChange Then Exit Sub
    mblnChangeEndDate = True
    
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Set mobjChargeBill = New clsChargeBill
    Call InitPanel
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mobjChargeBill = Nothing
    mstrOperatorName = ""
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
            txtMemo.Width = cmdOK.Left - txtMemo.Left - 200
        End If
    End With
End Sub

Private Sub txtHandIn_Change()
    txtRemain.Text = Format(mdblDefaultHandIn - Val(txtHandIn.Text), "0.00")
End Sub

Private Sub txtHandIn_GotFocus()
    zlControl.TxtSelAll txtHandIn
End Sub

Private Sub txtHandIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtHandIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("-") Then
        If InStr(1, txtHandIn.Text, "-") > 0 Then
            KeyAscii = 0
            Exit Sub
        Else
            Exit Sub
        End If
    Else
        '限定输入数字
        If (KeyAscii < Asc(".") Or KeyAscii = Asc("/") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
            KeyAscii = 0
            Exit Sub
        End If
        '小数点的判断
        If KeyAscii = Asc(".") And InStr(1, txtHandIn.Text, ".") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
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
Public Sub ShowChargeList(ByVal frmMain As Object, Optional strRollingType As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细收款数据
    '入参:frmMain-调用的主窗体
    '    strRollingType-轧帐类别,bytType=1时有效分别为:
    '               0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡
    '编制:刘兴洪
    '日期:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
    Call frmNew.ShowMe(frmMain, mlngModule, mstrPrivs, 1, "", dtpStartDate.Value, dtpEndDate.Value, False, , strRollingType)
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
    'Dim lngDeptID As Long
    'If cboDept.ListIndex >= 0 Then lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "开始轧帐日期=" & Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM:SS"), _
        "终止轧帐日期=" & Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM:SS"))
End Sub

Public Sub zlDefaultSetFocus()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省输入
    '编制:刘兴洪
    '日期:2013-10-16 14:23:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If dtpEndDate.Enabled And dtpEndDate.Visible Then
        dtpEndDate.SetFocus
'    ElseIf cboDept.Enabled And cboDept.Visible Then
'        cboDept.SetFocus
    ElseIf txtMemo.Enabled And txtMemo.Visible Then
        txtMemo.SetFocus
    End If
End Sub
Public Sub MainKeyDown(KeyCode As Integer, Shift As Integer)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:主窗体转入的功能键(即处理快键)
    '编制:刘兴洪
    '日期:2013-10-16 15:14:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Shift <> 4 Then Exit Sub
    If KeyCode = vbKeyZ Then
        If cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus
            Call cmdOK_Click
        End If
        Exit Sub
    End If
    If KeyCode = vbKeyO Then
        If cmdRefresh.Enabled And cmdRefresh.Visible Then
            cmdRefresh.SetFocus
            Call cmdRefresh_Click
        End If
        Exit Sub
    End If
    If KeyCode = vbKeyM Then
        If txtMemo.Enabled And txtMemo.Visible Then
            txtMemo.SetFocus
            zlControl.TxtSelAll txtMemo
        End If
        Exit Sub
    End If
End Sub

Private Sub txtRemain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
