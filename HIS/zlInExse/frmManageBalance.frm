VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBalance 
   Caption         =   "病人结帐管理"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmManageBalance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCons 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4140
      ScaleHeight     =   300
      ScaleWidth      =   7260
      TabIndex        =   4
      Top             =   1365
      Visible         =   0   'False
      Width           =   7260
      Begin VB.ComboBox cboDate 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   0
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   15
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   147390467
         CurrentDate     =   40777
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   285
         Left            =   4125
         TabIndex        =   7
         Top             =   15
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   147390467
         CurrentDate     =   40777
      End
      Begin VB.Label lbl缺省 
         AutoSize        =   -1  'True
         Caption         =   "缺省显示"
         Height          =   180
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblSplit 
         Caption         =   "～"
         Height          =   210
         Left            =   3870
         TabIndex        =   9
         Top             =   45
         Width           =   330
      End
      Begin VB.Label lblDateShow 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   2535
         TabIndex        =   8
         Top             =   45
         Width           =   90
      End
   End
   Begin VB.TextBox txtIdentify 
      Height          =   320
      Left            =   8295
      TabIndex        =   3
      Top             =   867
      Width           =   2160
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   315
      Left            =   7725
      TabIndex        =   10
      Top             =   870
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      IDKindStr       =   "姓|姓名或就诊卡|0|0|0|0|0|0;门|门诊号|0|0|0|0|0|0;住|住院号|0|0|0|0|0|0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "宋体"
      IDKind          =   -1
      DefaultCardType =   "0"
      BackColor       =   -2147483637
   End
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   2175
      _Version        =   589884
      _ExtentX        =   3836
      _ExtentY        =   2355
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8025
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   635
      SimpleText      =   $"frmManageBalance.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBalance.frx":05D1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13150
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   4110
      Top             =   2235
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmManageBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mcbrPopupMain As CommandBar, mcbrMenuView As CommandBarPopup, mcbrRefresh As CommandBarControl
Private mcbrCmb As CommandBarComboBox, mstrPrivs As String, mlngModule As Long
Private mfrmNormal As New frmBalanceTabNormal
Private mfrmErr As New frmBalanceTabErr
Private mfrmRefund As New frmBalanceTabRefund
Private mblnCancel As Boolean   '外部卸载窗体标识
Private mstrTitle As String '用于窗体个性化保存的窗体名
Private mrsInfo As ADODB.Recordset, mstrPrivsRollingCurtain As String
Private mblnFirst As Boolean, mstrWriteCardTypeIDs As String
Private mobjInPati As Object, mbln立即销帐 As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mbln传统模式 As Boolean

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnOk As Boolean
    On Error GoTo errHandle
    Select Case Control.ID
        Case conMenu_File_FeeCollect
            If zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "轧帐") = False Then Exit Sub
            Call zlExecuteChargeRollingCurtain(Me)
        Case conMenu_File_SetInsure
            gclsInsure.InsureSupport
        Case conMenu_File_CashCount
            Call frmMoneyEnum.ShowMe(Me)
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            Control.Checked = stbThis.Visible
            Form_Resize
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            Control.Checked = Not Control.Checked
            cbsThis(2).Visible = Not cbsThis(2).Visible
            Form_Resize
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not Control.Checked
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_Filter
            Call mfrmNormal.MakeFilter(Me, mlngModule, mstrPrivs)
            txtIdentify.Text = ""
        Case conMenu_File_PrintSet
            zlPrintSet
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case conMenu_Edit_RefundDeposit
            Call RefundDeposit
        Case conMenu_View_Refresh
            Select Case tabMain.Selected.Index
                Case 0
                    Call mfrmNormal.ReadData(0, mstrPrivs)
                Case 1
                    Call mfrmErr.ReadData
                Case 2
                    Call mfrmRefund.ReadData
            End Select
        Case conMenu_File_Parameter
            If zlStr.IsHavePrivs(mstrPrivs, "参数设置") = False Then Exit Sub
            frmSetExpence.mlngModul = mlngModule
            frmSetExpence.mstrPrivs = mstrPrivs
            frmSetExpence.mbytInFun = 1
            frmSetExpence.Show 1, Me
            Call InitLocPar(mlngModule)
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Preview
            Call zlRptPrint(2)
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_Help_Help
            ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
        Case conMenu_Help_Web_Home
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_Web_Forum
            Call zlWebForum(Me.hWnd)
        Case conMenu_Help_About
            ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
        Case conMenu_View_Location
            frmBalanceGo.Show 1, Me
            If gblnOK Then Call SeekBill(frmBalanceGo.optHead)
        Case conMenu_View_RefreshType_No
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked = False
            zlDatabase.SetPara "刷新方式", "0", glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        Case conMenu_View_RefreshType_Ask
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked = False
            zlDatabase.SetPara "刷新方式", "1", glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        Case conMenu_View_RefreshType_Auto
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            zlDatabase.SetPara "刷新方式", "2", glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        Case conMenu_Edit_ClinicBalance
            If zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") = False Then Exit Sub
            '门诊费用结帐
            If mbln传统模式 Then
                blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_门诊结帐, mstrPrivs)
            Else
                blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_门诊结帐, mstrPrivs)
            End If
            If blnOk Then Call RefreshData
        Case conMenu_Edit_InHosBalance
            If zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") = False Then Exit Sub
            '住院费用结帐
            If mbln传统模式 Then
                blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_住院结帐, mstrPrivs)
            Else
                blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_住院结帐, mstrPrivs)
            End If
            If blnOk Then Call RefreshData
        Case conMenu_Edit_ErrReBalance
            '异常重结
            With mfrmErr.vsfMain
                If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
                If mbln传统模式 Then
                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_重新结帐, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")))
                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_重新结帐, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")))
                End If
                If blnOk Then Call RefreshData
            End With
        Case conMenu_Edit_ErrCancelBalance
            '异常作废
            With mfrmErr.vsfMain
                If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
'                If mbln传统模式 Then
'                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_取消结帐, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")))
'                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_取消结帐, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")))
'                End If
                If blnOk Then Call RefreshData
            End With
        Case conMenu_Edit_ErrDelBalance
            '异常重退
            With mfrmRefund.vsfMain
                If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
'                If mbln传统模式 Then
'                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_重新作废, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")), True)
'                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_重新作废, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")), True)
'                End If
                If blnOk Then Call RefreshData
                
            End With
        Case conMenu_Edit_CancelBalance
            '单据作废
            With mfrmNormal.vsfMain
                If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
'                If mbln传统模式 Then
'                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_结帐作废, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")))
'                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_结帐作废, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("单据号")))
'                End If
                If blnOk Then Call RefreshData
                
            End With
        Case conMenu_Edit_BatchBalance
            If frmBalanceBat.ShowMe(Me, mstrPrivs) = True Then
                Call RefreshData
            End If
            
        Case conMenu_Edit_UnitBalance
            gblnOK = False
            frmBalanceUnit.ShowMe Me, 0, mlngModule, mstrPrivs
                
            If gblnOK Then
                Call RefreshData
            End If
        Case conMenu_Edit_FeeManage
            frmManageDue.mstrPrivs = mstrPrivs
            frmManageDue.mlngModul = mlngModule
            frmManageDue.Show 0, Me
        Case conMenu_Edit_ClinicToHos
            If InStr(1, mstrPrivs, ";门诊费用转住院;") = 0 Then Exit Sub
            If mobjInPati Is Nothing Then
                Err = 0: On Error Resume Next
                Set mobjInPati = CreateObject("zl9InPatient.clsInPatient")
                
                If Err <> 0 Then
                    MsgBox "注意:" & vbCrLf & "   住院病人部件(zl9InPatient)创建失败,请与系统管理员联系!"
                    Exit Sub
                End If
            End If
            Call mobjInPati.zlOutFeeToInFee(Me, gcnOracle, glngSys, mlngModule, mstrPrivs, gstrDBUser, 0, 0)
        Case conMenu_Edit_ToHosCancel
            If InStr(mstrPrivs, ";转住院费用销帐;") = 0 Or mbln立即销帐 Then Exit Sub
            If frmFeeRefundment.zlShowEdit(Me, 2, mlngModule, mstrPrivs) = False Then
                Exit Sub
            End If
            Call RefreshData
        Case conMenu_Edit_View
            Call ViewBalance(tabMain.Selected.Index)
        Case conMenu_Edit_PrintAmend
            Call PrintBill(1)
        Case conMenu_Edit_ReprintReceipt
            Call PrintBill(0)
        Case conMenu_Edit_PrintDetail
            Call PrintDetail
        Case conMenu_Edit_PrintAmendByPati
            '按病人补打票据
            If frmMakeupPrintBill.zlRePrintBill(Me, mlngModule, mstrPrivs) = True Then
                Call RefreshData
            End If
            
        Case conMenu_Edit_WriteCard
            Call WriteCard
        Case Else
            If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefundDeposit()
'---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:余额退款
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun          As Object
    
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Sub
    
    If objFun.RefundDeposit(glngSys, gcnOracle, Me, gstrDBUser) = False Then
        Set objFun = Nothing
        Exit Sub
    End If
    Set objFun = Nothing
End Sub

Private Sub WriteCard()
    Dim lngCardTypeID As Long, strExpend As String, lng病人ID As Long
    Dim lng结帐ID As Long, strNO As String, lng记录状态 As Long
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim bytFunc As Byte
    
    With mfrmNormal.vsfMain
        strNO = .TextMatrix(.Row, .ColIndex("单据号"))
        lng结帐ID = Val(.TextMatrix(.Row, .ColIndex("结帐ID")))
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        bytFunc = IIf(Val(.TextMatrix(.Row, .ColIndex("标志"))) = 1, 0, 1)
    End With
    '功能:将住院信息写入卡中
    '问题:56615
    If mstrWriteCardTypeIDs = "" Then Exit Sub
    If bytFunc = 0 Then '门诊记帐费用
        If Not zlStr.IsHavePrivs(mstrPrivs, "门诊信息写卡") Then Exit Sub
    Else
        If Not zlStr.IsHavePrivs(mstrPrivs, "住院信息写卡") Then Exit Sub
    End If
    
     If strNO = "" Then
        MsgBox "当前没有单据可以重新写卡！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If lng病人ID = 0 Or lng结帐ID = 0 Then Exit Sub
    If InStr(1, mstrWriteCardTypeIDs, ",") = 0 Then lngCardTypeID = Val(mstrWriteCardTypeIDs)
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, gobjSquare.objSquareCard, lngCardTypeID, bytFunc, lng结帐ID, lng病人ID)
End Sub

Private Sub CheckErrBill()
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim dtStartDate As Date, dtEndDate As Date

    dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
    dtEndDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")

    strSql = " Select Count(1) As 记录数" & vbNewLine & _
             " From 病人结帐记录" & vbNewLine & _
             " Where 收费时间 Between [1] And [2] And Nvl(结算状态, 0) = 1 And 记录状态 = 2 And 操作员姓名 = [3]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
    
    If rsTmp.RecordCount <> 0 Then
        If Val(NVL(rsTmp!记录数)) <> 0 Then
            tabMain.Item(2).Caption = "异常退费记录(" & Val(NVL(rsTmp!记录数)) & ")"
            If MsgBox("存在结帐异常退费记录,是否处理异常记录?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tabMain.Item(2).Selected = True
                Call mfrmRefund.ReadData
                Exit Sub
            End If
        End If
    Else
        tabMain.Item(2).Caption = "异常退费记录"
    End If
    
    strSql = " Select Count(1) As 记录数" & vbNewLine & _
             " From 病人结帐记录 A" & vbNewLine & _
             " Where a.收费时间 Between [1] And [2] And Nvl(a.结算状态, 0) = 1 And a.记录状态 In (1, 3) And a.操作员姓名 = [3] And" & vbNewLine & _
             "       Not Exists (Select 1 From 病人结帐记录 Where NO = a.No And 记录状态 = 2)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
    
    If rsTmp.RecordCount <> 0 Then
        If Val(NVL(rsTmp!记录数)) <> 0 Then
            tabMain.Item(1).Caption = "异常结算记录(" & Val(NVL(rsTmp!记录数)) & ")"
            If MsgBox("存在结帐异常结算记录,是否处理异常记录?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tabMain.Item(1).Selected = True
                Call mfrmErr.ReadData
                Exit Sub
            End If
        End If
    Else
        tabMain.Item(1).Caption = "异常结算记录"
    End If

End Sub

Public Sub ViewBalance(intTYPE As Integer)
    '查阅单据
    Select Case intTYPE
        Case 0
            If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结帐ID")) = "" Then Exit Sub
            frmPatiBalanceSplit.ShowMe Me, g_Ed_单据查看, mstrPrivs, , , _
                                       mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("单据号")), _
                                       Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("记录状态"))) = 2, _
                                       zlDatabase.NOMoved("病人结帐记录", mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("单据号")))
        Case 1
            If mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("结帐ID")) = "" Then Exit Sub
            frmPatiBalanceSplit.ShowMe Me, g_Ed_单据查看, mstrPrivs, , , mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("单据号")), False
        Case 2
            If mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("结帐ID")) = "" Then Exit Sub
            frmPatiBalanceSplit.ShowMe Me, g_Ed_单据查看, mstrPrivs, , , mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("单据号")), True
    End Select
End Sub

Private Sub RefreshData()
    If mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked Then
        If MsgBox("当前已发生操作,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Select Case tabMain.Selected.Index
                Case 0
                    Call mfrmNormal.ReadData(0, mstrPrivs)
                Case 1
                    Call mfrmErr.ReadData
                Case 2
                    Call mfrmRefund.ReadData
            End Select
        End If
    ElseIf mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked Then
        Select Case tabMain.Selected.Index
            Case 0
                Call mfrmNormal.ReadData(0, mstrPrivs)
            Case 1
                Call mfrmErr.ReadData
            Case 2
                Call mfrmRefund.ReadData
        End Select
    End If
End Sub

Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String, Optional ByVal intTYPE As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode-报表编号
    '     intType-报表操作类型:0-默认,1-直接预览,2-直接打印,3-输出到EXCEL
    '编制:刘尔旋
    '日期:2013-09-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String
    Select Case tabMain.Selected.Index
    Case 0
        With mfrmNormal.vsfMain
            strNO = .TextMatrix(.Row, .ColIndex("单据号"))
            If strNO = "" Then
                Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, intTYPE)
            Else
                Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, _
                "病人ID=" & .TextMatrix(.Row, .ColIndex("病人ID")), _
                "住院号=" & .TextMatrix(.Row, .ColIndex("住院号")), _
                "结帐ID=" & .TextMatrix(.Row, .ColIndex("结帐ID")), _
                "NO=" & strNO, _
                "记录状态=" & .TextMatrix(.Row, .ColIndex("记录状态")), intTYPE)
            End If
        End With
    Case Else
        Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, intTYPE)
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With tabMain
        .Left = Left
        .Top = Top
        .Width = Right - Left
        .Height = Bottom - Top
    End With
    picCons.Top = tabMain.Top + 15
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, blnCollect As Boolean, bytFunc As Byte
    
    If tabMain.Selected.Index = 0 Then
        '普通结帐记录按钮控制
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_Edit_View
                Control.Enabled = mfrmNormal.vsfMain.TextMatrix(1, mfrmNormal.vsfMain.ColIndex("结帐ID")) <> ""
            Case conMenu_Edit_ErrReBalance, conMenu_Edit_ErrCancelBalance, conMenu_Edit_ErrDelBalance
                Control.Enabled = False
                Control.Visible = False
            Case conMenu_Edit_CancelBalance
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
                If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结帐ID")) <> "" Then
                    Control.Enabled = Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("记录状态"))) = 1
                    If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("医保")) <> "" Then
                        If zlStr.IsHavePrivs(mstrPrivs, "保险结算") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "普通病人结算") = False Then Control.Enabled = False
                    End If
                Else
                    Control.Enabled = False
                End If
            Case conMenu_Edit_ReprintReceipt
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "重打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
                If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结帐ID")) <> "" Then
                    Control.Enabled = Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("记录状态"))) = 1 _
                                    And mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("票据号")) <> ""
                    If InStr(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("  结帐类型")), "门诊") > 0 Then
                        If zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") = False Then Control.Enabled = False
                    End If
                    If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("医保")) <> "" Then
                        If zlStr.IsHavePrivs(mstrPrivs, "保险结算") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "普通病人结算") = False Then Control.Enabled = False
                    End If
                Else
                    Control.Enabled = False
                End If
            Case conMenu_Edit_PrintAmend
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "补打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
                If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结帐ID")) <> "" Then
                    Control.Enabled = Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("记录状态"))) = 1 _
                                    And mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("票据号")) = ""
                    If InStr(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("  结帐类型")), "门诊") > 0 Then
                        If zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") = False Then Control.Enabled = False
                    End If
                    If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("医保")) <> "" Then
                        If zlStr.IsHavePrivs(mstrPrivs, "保险结算") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "普通病人结算") = False Then Control.Enabled = False
                    End If
                Else
                    Control.Enabled = False
                End If
            Case conMenu_Edit_WriteCard
                bytFunc = IIf(Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("标志"))) = 1, 0, 1)
                Control.Visible = (zlStr.IsHavePrivs(mstrPrivs, "住院信息写卡") Or zlStr.IsHavePrivs(mstrPrivs, "门诊信息写卡")) _
                                And mstrWriteCardTypeIDs <> ""
                Control.Enabled = (bytFunc = 0 And zlStr.IsHavePrivs(mstrPrivs, "门诊信息写卡")) _
                        Or (bytFunc = 1 And zlStr.IsHavePrivs(mstrPrivs, "住院信息写卡")) _
                        And mfrmNormal.vsfMain.TextMatrix(1, mfrmNormal.vsfMain.ColIndex("结帐ID")) <> ""
            Case conMenu_View_Filter
                Control.Visible = True
                Control.Enabled = True
        End Select
        IDKind.Visible = True
        txtIdentify.Visible = True
    ElseIf tabMain.Selected.Index = 1 Then
        '异常结算记录按钮控制
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_Edit_View
                Control.Enabled = mfrmErr.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("结帐ID")) <> ""
            Case conMenu_Edit_ErrReBalance, conMenu_Edit_ErrCancelBalance
                Control.Visible = True
                Control.Enabled = mfrmErr.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("结帐ID")) <> ""
            Case conMenu_Edit_ErrDelBalance, conMenu_Edit_CancelBalance
                Control.Visible = False
                Control.Enabled = False
            Case conMenu_Edit_ReprintReceipt, conMenu_Edit_PrintAmend
                Control.Visible = False
                Control.Enabled = False
            Case conMenu_Edit_WriteCard, conMenu_View_Filter
                Control.Visible = False
                Control.Enabled = False
        End Select
        IDKind.Visible = False
        txtIdentify.Visible = False
    ElseIf tabMain.Selected.Index = 2 Then
        '异常作废记录按钮控制
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_Edit_View
                Control.Enabled = mfrmErr.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("结帐ID")) <> ""
            Case conMenu_Edit_ErrReBalance, conMenu_Edit_ErrCancelBalance, conMenu_Edit_CancelBalance
                Control.Visible = False
                Control.Enabled = False
            Case conMenu_Edit_ErrDelBalance
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
                Control.Enabled = mfrmRefund.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("结帐ID")) <> ""
            Case conMenu_Edit_ReprintReceipt, conMenu_Edit_PrintAmend
                Control.Visible = False
                Control.Enabled = False
            Case conMenu_Edit_WriteCard, conMenu_View_Filter
                Control.Visible = False
                Control.Enabled = False
        End Select
        IDKind.Visible = False
        txtIdentify.Visible = False
    End If
End Sub

Public Sub FailInit()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:外部窗体调用卸载窗体,赋值变量从FORMLOAD中退出
    '编制:刘尔旋
    '日期:2013-10-11
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    mblnCancel = True
End Sub

Private Sub SetTabControl()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建TAB控件
    '编制:刘尔旋
    '日期:2013-09-04
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With tabMain
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        .InsertItem 1, "正常结算记录", mfrmNormal.hWnd, 0
        .InsertItem 2, "异常结算记录", mfrmErr.hWnd, 0
        .InsertItem 3, "异常退费记录", mfrmRefund.hWnd, 0
        .Item(0).Selected = True
        .PaintManager.BoldSelected = True
        .PaintManager.ClientFrame = xtpTabFrameNone
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModule)
    If mblnFirst Then
        mblnFirst = False
        tabMain.Item(0).Selected = True
        Call CheckErrBill
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mlngModule = glngModul
    mstrPrivs = gstrPrivs
    mstrPrivsRollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mblnCancel = False
    mstrTitle = "病人结帐管理"
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
    mbln传统模式 = Val(zlDatabase.GetPara("结帐界面风格", glngSys, mlngModule, "1")) = 0
    mstrWriteCardTypeIDs = ""
    If Not gobjSquare Is Nothing Then
        If Not gobjSquare.objSquareCard Is Nothing Then
            mstrWriteCardTypeIDs = gobjSquare.objSquareCard.zlGetAvailabilityWriteCardType
        End If
    End If
    Call zlDefCommandBars
    '创建TAB信息
    Call SetTabControl
    Call InitIDKind
    Call SetCboDate
    stbThis.Panels(3).Text = UserInfo.姓名
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
    '创建并检测税控打印对象
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(2)
        End If
        On Error GoTo 0
    End If
    '创建第三方票据打印部件
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.编号, UserInfo.姓名)
    End If
    On Error GoTo 0
End Sub

Public Sub ShowPopup()
    mcbrPopupMain.ShowPopup
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘尔旋
    '日期:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    If mfrmNormal.ActiveControl Is mfrmNormal.vsfMain And tabMain.Selected.Index = 0 Then
        With mfrmNormal.vsfMain
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("结帐ID"))) = 0 Then Exit Sub
        End With
        Call mfrmNormal.zlRptPrint(bytFunc)
    End If
    
    If mfrmErr.ActiveControl Is mfrmErr.vsfMain And tabMain.Selected.Index = 1 Then
        With mfrmErr.vsfMain
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("结帐ID"))) = 0 Then Exit Sub
        End With
        Call mfrmErr.zlRptPrint(bytFunc)
    End If
    
    If mfrmRefund.ActiveControl Is mfrmRefund.vsfMain And tabMain.Selected.Index = 2 Then
        With mfrmRefund.vsfMain
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("结帐ID"))) = 0 Then Exit Sub
        End With
        Call mfrmRefund.zlRptPrint(bytFunc)
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim intPara As Integer
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    '初始化设置
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&U)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_CashCount, "现金点钞(&D)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "收费扎帐(&M)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "轧帐")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_SetInsure, "保险类别(&I)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicBalance, "门诊结帐(&M)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InHosBalance, "住院结帐(&A)")
        mcbrControl.IconId = 3590
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BatchBalance, "批量中途结帐(&T)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "批量中途结帐")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_UnitBalance, "合约单位结帐(&U)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_FeeManage, "应收款管理(&Y)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "应收款管理")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RefundDeposit, "余额退款(&R)")
        mcbrControl.IconId = 3017
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicToHos, "门诊费用转住院(&Z)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "门诊费用转住院")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ToHosCancel, "转住院费用销帐(&X)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "转住院费用销帐") And Not mbln立即销帐
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "异常重结(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "异常作废(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrDelBalance, "异常重退(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelBalance, "单据作废(&D)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 4114
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_View, "查阅单据(&V)")
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "重打结帐票据(&R)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "重打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "补打结帐票据(&B)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "补打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintDetail, "打印结帐明细(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmendByPati, "按病人补打结帐票据(&P)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "补打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_WriteCard, "结帐信息写卡(&W)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = (zlStr.IsHavePrivs(mstrPrivs, "住院信息写卡") Or zlStr.IsHavePrivs(mstrPrivs, "门诊信息写卡")) _
                                And mstrWriteCardTypeIDs <> ""
    End With
    
    Set mcbrMenuView = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuView.ID = conMenu_ViewPopup
    With mcbrMenuView.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        cbrControl.Checked = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤(&F)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Location, "定位(&G)")
        intPara = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModule, "0"))
        Set mcbrRefresh = .Add(xtpControlPopup, conMenu_View_RefreshType, "刷新方式(&O)"): mcbrControl.BeginGroup = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_No, "操作后不刷新数据(&1)", -1, False)
        If intPara = 0 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Ask, "操作后提示刷新数据(&2)", -1, False)
        If intPara = 1 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Auto, "操作后自动刷新数据(&3)", -1, False)
        If intPara = 2 Then cbrControl.Checked = True
        mcbrRefresh.Visible = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的中联")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&K)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '创建弹出菜单
    Set mcbrPopupMain = cbsThis.Add("弹出菜单1", xtpBarPopup)
    With mcbrPopupMain.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到Excel")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "收费轧帐"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "轧帐")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicBalance, "门诊结帐(&M)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InHosBalance, "住院结帐(&A)")
        mcbrControl.IconId = 3590
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BatchBalance, "批量中途结帐(&T)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "批量中途结帐")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_UnitBalance, "合约单位结帐(&U)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_FeeManage, "应收款管理(&Y)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "应收款管理")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicToHos, "门诊费用转住院(&Z)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "门诊费用转住院")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ToHosCancel, "转住院费用销帐(&X)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "转住院费用销帐") And Not mbln立即销帐
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "异常重结(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "异常作废(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrDelBalance, "异常重退(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelBalance, "单据作废(&D)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 4114
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_View, "查阅单据(&V)")
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "重打结帐票据(&R)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "重打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "补打结帐票据(&B)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "补打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintDetail, "打印结帐明细(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmendByPati, "按病人补打结帐票据(&P)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "补打票据") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))) Or (zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_WriteCard, "结帐信息写卡(&W)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = (zlStr.IsHavePrivs(mstrPrivs, "住院信息写卡") Or zlStr.IsHavePrivs(mstrPrivs, "门诊信息写卡")) _
                                And mstrWriteCardTypeIDs <> ""
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Location, "定位")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F9, conMenu_File_CashCount
        .Add 0, VK_F11, conMenu_File_FeeCollect
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add FCONTROL, Asc("M"), conMenu_Edit_ClinicBalance
        .Add 0, VK_F2, conMenu_Edit_BatchBalance
        .Add 0, VK_F4, conMenu_Edit_UnitBalance
        .Add 0, VK_F8, conMenu_Edit_FeeManage
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add FCONTROL, Asc("G"), conMenu_View_Location
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ModifyStyle &H400000, 0
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")

        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicBalance, "门诊"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "门诊费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InHosBalance, "住院")
        mcbrControl.IconId = 3590
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "住院费用结帐") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))

        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "异常重结"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "异常作废")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrDelBalance, "异常重退"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelBalance, "作废"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 4114
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "结帐作废") And (zlStr.IsHavePrivs(mstrPrivs, "保险结算") Or zlStr.IsHavePrivs(mstrPrivs, "普通病人结算"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_View, "查阅")
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "轧帐"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "轧帐")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): mcbrControl.BeginGroup = True
    End With
    
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_Edit_UserType Then
            mcbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    zlDefCommandBars = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picCons.Left = 5000

    IDKind.Top = 30
    txtIdentify.Top = 30
    IDKind.Left = Me.Width - 3105
    txtIdentify.Left = IDKind.Left + IDKind.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If Not mfrmNormal Is Nothing Then Unload mfrmNormal: Set mfrmNormal = Nothing
    If Not mfrmErr Is Nothing Then Unload mfrmErr: Set mfrmErr = Nothing
    If Not mfrmRefund Is Nothing Then Unload mfrmRefund: Set mfrmRefund = Nothing
    
    '存储列表的个性化设置(本地)
    
    Call SaveRegInFor(g私有模块, Me.Name, "异常单据查询", cboDate.ListIndex)
    SaveWinState Me, App.ProductName, mstrTitle
    '卸载加载窗体和类
    Set mrsInfo = Nothing

    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
    
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, int来源 As Integer
    Dim blnFill As Boolean
    Dim strNO As String
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    Select Case tabMain.Selected.Index
    Case 0
        For i = IIf(blnHead, 1, mlngGo) To mfrmNormal.vsfMain.Rows - 1
            DoEvents
            
            '比较条件
            blnFill = True
            With frmBalanceGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("单据号")) = .txtNO.Text
                End If
                If .txtFact.Text <> "" Then
                    blnFill = blnFill And mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("票据号")) = .txtFact.Text
                End If
                If .txt住院号.Text <> "" Then
                    blnFill = blnFill And mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("住院号")) = .txt住院号.Text
                End If
                If .txt姓名.Text <> "" Then
                    blnFill = blnFill And UCase(mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
                End If
            End With
            
            '满足则退出
            If blnFill Then
                mlngGo = i + 1
                mfrmNormal.vsfMain.Row = i: mfrmNormal.vsfMain.TopRow = i
                mfrmNormal.vsfMain.Col = 0: mfrmNormal.vsfMain.ColSel = mfrmNormal.vsfMain.Cols - 1
                
                stbThis.Panels(2).Text = "找到一条记录"
                Screen.MousePointer = 0: Exit Sub
            End If
            
            '按ESC取消
            If mblnGo = False Then
                stbThis.Panels(2).Text = "用户取消定位操作"
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    Case 1
        For i = IIf(blnHead, 1, mlngGo) To mfrmErr.vsfMain.Rows - 1
            DoEvents
            
            '比较条件
            blnFill = True
            With frmBalanceGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("单据号")) = .txtNO.Text
                End If
                If .txtFact.Text <> "" Then
                    blnFill = blnFill And mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("票据号")) = .txtFact.Text
                End If
                If .txt住院号.Text <> "" Then
                    blnFill = blnFill And mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("住院号")) = .txt住院号.Text
                End If
                If .txt姓名.Text <> "" Then
                    blnFill = blnFill And UCase(mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
                End If
            End With
            
            '满足则退出
            If blnFill Then
                mlngGo = i + 1
                mfrmErr.vsfMain.Row = i: mfrmErr.vsfMain.TopRow = i
                mfrmErr.vsfMain.Col = 0: mfrmErr.vsfMain.ColSel = mfrmErr.vsfMain.Cols - 1
                
                stbThis.Panels(2).Text = "找到一条记录"
                Screen.MousePointer = 0: Exit Sub
            End If
            
            '按ESC取消
            If mblnGo = False Then
                stbThis.Panels(2).Text = "用户取消定位操作"
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    Case 2
        For i = IIf(blnHead, 1, mlngGo) To mfrmRefund.vsfMain.Rows - 1
            DoEvents
            
            '比较条件
            blnFill = True
            With frmBalanceGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("单据号")) = .txtNO.Text
                End If
                If .txtFact.Text <> "" Then
                    blnFill = blnFill And mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("票据号")) = .txtFact.Text
                End If
                If .txt住院号.Text <> "" Then
                    blnFill = blnFill And mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("住院号")) = .txt住院号.Text
                End If
                If .txt姓名.Text <> "" Then
                    blnFill = blnFill And UCase(mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
                End If
            End With
            
            '满足则退出
            If blnFill Then
                mlngGo = i + 1
                mfrmRefund.vsfMain.Row = i: mfrmRefund.vsfMain.TopRow = i
                mfrmRefund.vsfMain.Col = 0: mfrmRefund.vsfMain.ColSel = mfrmRefund.vsfMain.Cols - 1
                
                stbThis.Panels(2).Text = "找到一条记录"
                Screen.MousePointer = 0: Exit Sub
            End If
            
            '按ESC取消
            If mblnGo = False Then
                stbThis.Panels(2).Text = "用户取消定位操作"
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    End Select
    
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub

Private Sub cboDate_Click()
    Dim dtStartDate As Date, dtEndDate As Date
    lblSplit.Visible = cboDate.ListIndex = 6
    dtpStartDate.Visible = cboDate.ListIndex = 6
    dtpEndDate.Visible = cboDate.ListIndex = 6
    lblDateShow.Visible = cboDate.ListIndex <> 6 And cboDate.ListIndex <> 0
    Select Case cboDate.ListIndex
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
            dtStartDate = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    lblDateShow.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
    lblDateShow.Caption = lblDateShow.Caption & "~" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
    If cboDate.Visible = False Then Exit Sub
    If tabMain.Selected.Index = 1 Then
        Call mfrmErr.ReadData
    Else
        Call mfrmRefund.ReadData
    End If
End Sub

Private Sub SetCboDate()
    Dim i As Integer
    Dim strValue As String
    Call GetRegInFor(g私有模块, Me.Name, "异常单据查询", strValue)
    i = Val(strValue)
    With cboDate
        .Clear
        .AddItem "所有异常情况"
        .ListIndex = .NewIndex
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "今日"
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "最近两天"
        If i = 2 Then .ListIndex = .NewIndex
        .AddItem "最近三天"
        If i = 3 Then .ListIndex = .NewIndex
        .AddItem "最近一周"
        If i = 4 Then .ListIndex = .NewIndex
        .AddItem "本月"
        If i = 5 Then .ListIndex = .NewIndex
        .AddItem "自定义时间范围"
        dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpEndDate.MaxDate = dtpStartDate.MaxDate
        dtpEndDate.Value = dtpEndDate.MaxDate
        dtpStartDate.Value = DateAdd("d", -7, dtpEndDate.MaxDate)
    End With
    Call cboDate_Click
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtIdentify.Enabled And txtIdentify.Visible Then txtIdentify.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtIdentify.Locked Then Exit Sub
    txtIdentify.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, txtIdentify.Text)
End Sub

Private Sub tabMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    stbThis.Panels(2).Text = ""
    Select Case Item.Index
        Case 0
            picCons.Visible = False
            If mblnFirst Then Exit Sub
            Call mfrmNormal.ReadData(0, mstrPrivs)
        Case 1
            picCons.Visible = True
            Call mfrmErr.ReadData
        Case 2
            picCons.Visible = True
            Call mfrmRefund.ReadData
    End Select
End Sub

Private Sub txtIdentify_Change()
    txtIdentify.Tag = ""
    If Me.ActiveControl Is txtIdentify Then
        IDKind.SetAutoReadCard txtIdentify.Text = ""
    End If
End Sub

Private Sub txtIdentify_GotFocus()
    Call zlControl.TxtSelAll(txtIdentify)
    Call zlCommFun.OpenIme(True)
    If txtIdentify.Text = "" And ActiveControl Is txtIdentify Then IDKind.SetAutoReadCard True
End Sub

Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.名称
            Else
                If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
            End If
     End Select
End Function

Private Sub txtIdentify_KeyPress(KeyAscii As Integer)
  Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSql As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtIdentify.Locked Then Exit Sub
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "姓名") Then
        blnCard = zlCommFun.InputIsCard(txtIdentify, KeyAscii, IDKind.ShowPassText)
    ElseIf IsCardType(IDKind, "门诊号") Or IsCardType(IDKind, "住院号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtIdentify.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtIdentify.IMEMode = 0
    End If
    If blnCard And Len(txtIdentify.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtIdentify.Text) <> "" Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(txtIdentify.Tag) <> 0 Then    '存在
                 zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
        If KeyAscii <> 13 Then
            txtIdentify.Text = txtIdentify.Text & Chr(KeyAscii)
            txtIdentify.SelStart = Len(txtIdentify.Text)
            KeyAscii = 0
        End If
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtIdentify.Text))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog '
End Sub

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    IDKind.AllowAutoCommCard = True
    IDKind.AllowAutoICCard = True
    IDKind.AllowAutoIDCard = True
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtIdentify)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
End Function

'获取默认IDKind索引
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-09-03 09:32:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not GetPatient(objCard, strInput, blnCard) Then
        MsgBox "未找到病人，请重新输入！", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case tabMain.Selected.Index
    Case 0
        Call mfrmNormal.ReadData(1, mstrPrivs, Val(NVL(mrsInfo!ID)))
    End Select
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：blnCard=是否就诊卡刷卡
    '返回：查找成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-07-16 14:24:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim strSql As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    
    strSql = ""
    If blnCard And objCard.名称 Like "姓名*" And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        If lng病人ID <= 0 Then lng病人ID = 0
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSql = strSql & " And B.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSql = strSql & " And B.门诊号=[2]" & str非在院
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSql = strSql & " And B.病人ID=[2]" & str非在院
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strSql = strSql & " And B.病人ID = (Select Nvl(Max(病人ID),0) As 病人ID From 病案主页   Where  住院号=[1])"
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                '姓名
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtIdentify.Text = mrsInfo!姓名 Then blnSame = True
                End If
                
                If Not blnSame Then
                    'strSQL = strSQL & " And  B.姓名 Like [3]"
                    '问题号:50485
                     strPati = _
                         " Select /*+Rule */distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位,decode(b.卡号,Null,Null,'√') As 是否有医疗卡" & _
                         " From 病人信息 A, 病人医疗卡信息 B " & _
                         " Where Rownum <101 And a.病人ID=b.病人ID(+) And b.状态(+)=0 And B.卡类别ID(+)=[2]  And A.停用时间 is NULL And A.姓名 Like [1]" & str非在院
                         
                     vRect = zlControl.GetControlRect(txtIdentify.hWnd)
                     Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtIdentify.Height, blnCancel, False, True, strInput & "%", Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, 0)))
                     If Not rsTmp Is Nothing Then
                         If rsTmp!ID = 0 Then
                             Set mrsInfo = Nothing: Exit Function
                         Else
                             strInput = "-" & rsTmp!病人ID
                             strSql = strSql & " And B.病人ID=[2]"
                         End If
                     Else '取消选择
                         txtIdentify.Text = ""
                         Set mrsInfo = Nothing: Exit Function
                     End If
                Else
                    strSql = strSql & " And B.病人ID=[2]"
                    strInput = "-" & Val(mrsInfo!病人ID)
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strSql = strSql & " And B.医保号=[1]" & str非在院
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSql = strSql & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                ' strSQL = strSQL & " And B.身份证号=[1] " & str非在院
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSql = strSql & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And B.门诊号=[1]" & str非在院
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And B.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[1]) " & str非在院
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    If lng病人ID = 0 Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID <= 0 Then lng病人ID = 0
                strSql = strSql & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSql
    strSql = "    " & vbNewLine & " Select /*+Rule */distinct  B.病人id As ID, Decode(sign(nvl(ylkxx.病人id,0)),0,'','√') as 三方账户, B.病人id,B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位,"
    strSql = strSql & vbNewLine & "      A.名称 险类名称"
    strSql = strSql & vbNewLine & " From 病人信息 B, 保险类别 A,医疗卡类别 YLK,病人医疗卡信息 YLKXX"
    strSql = strSql & vbNewLine & " Where B.险类 = A.序号(+) and b.病人id=ylkxx.病人id(+) and ylkxx.状态(+)=0 and  ylkxx.卡类别id=ylk.id(+)  and ylk.是否自制(+)=0 And B.停用时间 Is Null   "
    strSql = strSql & vbNewLine & strTmp
    
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
        
    If mrsInfo Is Nothing Then GoTo ClearPati:
    If mrsInfo.State <> 1 Then GoTo ClearPati:
    If mrsInfo.RecordCount = 0 Then GoTo ClearPati:
    If Val(NVL(mrsInfo!ID)) = 0 Then GoTo ClearPati:
    
    txtIdentify.Text = NVL(mrsInfo!姓名)
    Me.txtIdentify.Tag = NVL(mrsInfo!ID)
    txtIdentify.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtIdentify.IMEMode = 0
    GetPatient = True
    Exit Function
ClearPati:
    txtIdentify.Text = ""
    txtIdentify.PasswordChar = ""
    Set mrsInfo = Nothing
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtIdentify.IMEMode = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub PrintDetail()
'功能：输入出列表
    Dim strNO As String
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    strNO = mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以打印证明！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    intRow = mfrmNormal.vsfDetail.Row
    
    '表头
    objOut.Title.Text = "病人结帐单据明细"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmBalanceFilter
        objRow.Add "单据号：" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("单据号"))
        objRow.Add "结帐范围：" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("开始日期")) & " 至 " & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结束日期"))
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "住院号：" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("住院号"))
        objRow.Add "姓名：" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("姓名"))
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mfrmNormal.vsfDetail.Redraw = False
    Set objOut.Body = mfrmNormal.vsfDetail
    
    bytR = zlPrintAsk(objOut)
    Me.Refresh
    If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    
    mfrmNormal.vsfDetail.Row = intRow
    mfrmNormal.vsfDetail.Col = 0: mfrmNormal.vsfDetail.ColSel = mfrmNormal.vsfDetail.Cols - 1
    mfrmNormal.vsfDetail.Redraw = True
End Sub

Private Sub PrintBill(bytMode As Byte)
'功能：当前收款记录重新打印一张票据
'bytMode=0-重打,1-补打
    Dim strNO As String, lng结帐ID As Long, blnMediCare As Boolean, bytFlag As Byte '门诊还是住院
    Dim intInsure As Integer
    Dim lng病人ID As Long, bytFunc As Byte
    
    With mfrmNormal.vsfMain
        strNO = .TextMatrix(.Row, .ColIndex("单据号"))
        If strNO = "" Then
            MsgBox "当前没有单据可以重打票据！", vbInformation, gstrSysName
            Exit Sub
        End If
        lng结帐ID = Val(.TextMatrix(.Row, .ColIndex("结帐ID")))
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        bytFunc = IIf(Val(.TextMatrix(.Row, .ColIndex("标志"))) = 1, 0, 1)
        
         '单据权限
        If bytMode = 0 Then
            If Not BillOperCheck(7, .TextMatrix(.Row, .ColIndex("操作员")), _
                CDate(.TextMatrix(.Row, .ColIndex("收费时间"))), "重打") Then Exit Sub
        Else
            If Trim(.TextMatrix(.Row, .ColIndex("票据号"))) <> "" Then
                MsgBox "当前单据已打印过票据,不能进行补打！", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        
        intInsure = BalanceExistInsure(strNO, bytFlag)
        If RePrintBalance(strNO, Me, lng结帐ID, intInsure) Then
            '银医一卡通写卡，85950
            Call WriteInforToCard(Me, mlngModule, mstrPrivs, gobjSquare.objSquareCard, 0, bytFunc, lng结帐ID, lng病人ID)
            Call RefreshData
        End If
    End With
End Sub

Private Sub txtIdentify_LostFocus()
    IDKind.SetAutoReadCard False
End Sub
