VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceManage 
   Caption         =   "保险补充结算管理"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmReplenishTheBalanceManage.frx":0000
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
         Format          =   146800643
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
         Format          =   146800643
         CurrentDate     =   40777
      End
      Begin VB.Label lbl缺省 
         AutoSize        =   -1  'True
         Caption         =   "缺省显示"
         Height          =   180
         Left            =   60
         TabIndex        =   10
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
      TabIndex        =   2
      Top             =   870
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
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
      SimpleText      =   $"frmReplenishTheBalanceManage.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceManage.frx":05D1
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
Attribute VB_Name = "frmReplenishTheBalanceManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mcbrPopupMain As CommandBar, mcbrMenuView As CommandBarPopup, mcbrRefresh As CommandBarControl
Private mcbrCmb As CommandBarComboBox, mstrPrivs As String, mlngModule As Long
Private mfrmNormal As New frmDoubleBalanceNormal
Private mfrmErr As New frmDoubleBalanceErr
Private mfrmRefund As New frmDoubleBalanceRefund
Private mblnCancel As Boolean   '外部卸载窗体标识
Private mstrTitle As String '用于窗体个性化保存的窗体名
Private mrsInfo As ADODB.Recordset, mstrPrivsRollingCurtain As String
Private mobjInvoice As clsInvoice, mstrInvoice As String, mlng领用ID As Long
Private mobjFactProperty As clsFactProperty
Private mblnFirst As Boolean

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    Select Case Control.ID
        Case conMenu_File_FeeCollect
            If zlCheckPrivs(mstrPrivsRollingCurtain, "轧帐") = False Then Exit Sub
            Call zlExecuteChargeRollingCurtain(Me)
        Case conMenu_File_SetInsure
            gclsInsure.InsureSupport
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
        Case conMenu_File_PrintSet
            zlPrintSet
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case conMenu_View_Refresh
            Select Case tabMain.Selected.Index
                Case 0
                    Call mfrmNormal.ReadData(0, mstrPrivs)
                Case 1
                    Call mfrmErr.ReadData(0)
                Case 2
                    Call mfrmRefund.ReadData(0)
            End Select
        Case conMenu_File_Parameter
'            If zlCheckPrivs(mstrPrivs, "参数设置") = False Then Exit Sub
            If frmSetReplenishTheBalance.zlSetPara(Me, mlngModule, mstrPrivs) Then
                Call InitLocPar(1124)
            End If
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
        Case conMenu_View_RefreshType_No
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked = False
            zlDatabase.SetPara "刷新方式", "0", glngSys, mlngModule, zlCheckPrivs(mstrPrivs, "参数设置")
        Case conMenu_View_RefreshType_Ask
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked = False
            zlDatabase.SetPara "刷新方式", "1", glngSys, mlngModule, zlCheckPrivs(mstrPrivs, "参数设置")
        Case conMenu_View_RefreshType_Auto
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            zlDatabase.SetPara "刷新方式", "2", glngSys, mlngModule, zlCheckPrivs(mstrPrivs, "参数设置")
        Case conMenu_Edit_RegistBalance
            If zlCheckPrivs(mstrPrivs, "医保结算") = False Then Exit Sub
            If frmReplenishTheBalanceBill.zlEditCard(Me, mlngModule, mstrPrivs, EM_Balance_Register) Then
                Call RefreshData
            End If
        Case conMenu_Edit_InsureBalance
            If zlCheckPrivs(mstrPrivs, "医保结算") = False Then Exit Sub
            If frmReplenishTheBalanceBill.zlEditCard(Me, mlngModule, mstrPrivs, EM_Balance_Charge) Then
                Call RefreshData
            End If
        Case conMenu_Edit_BalanceDel
            If zlCheckPrivs(mstrPrivs, "结算退费") = False Then Exit Sub
'            If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结算序号")) = "" Then Exit Sub
            If frmReplenishTheBalanceDel.zlShowMe _
            (Me, mlngModule, mstrPrivs, EM_RBDTY_退费, mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结算序号"))) Then
                Call RefreshData
            End If
        Case conMenu_Edit_ReDel
            If zlCheckPrivs(mstrPrivs, "结算退费") = False Then Exit Sub
            If mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("结算序号")) = "" Then Exit Sub
            If frmReplenishTheBalanceDel.zlShowMe _
            (Me, mlngModule, mstrPrivs, EM_RBDTY_异常重退, mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("结算序号")), , , , mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("结算时间"))) Then
                Call RefreshData
            End If
        Case conMenu_Edit_ReBalance
            If zlCheckPrivs(mstrPrivs, "医保结算") = False Then Exit Sub
            If mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("结算ID")) = "" Then Exit Sub
            If frmReplenishTheBalanceBill.zlEditCard _
            (Me, mlngModule, mstrPrivs, EM_Balance_Err_ReCharge, mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("结算ID"))) Then
                Call RefreshData
            End If
        Case conMenu_Edit_BalanceCancel
            If zlCheckPrivs(mstrPrivs, "医保结算") = False Then Exit Sub
            With mfrmErr.vsfMain
                If .TextMatrix(.Row, .ColIndex("结算ID")) = "" Then Exit Sub
                If BalanceErrCancelCheck(Val(.TextMatrix(.Row, .ColIndex("结算ID")))) = False Then Exit Sub
                If frmReplenishTheBalanceBill.zlEditCard _
                    (Me, mlngModule, mstrPrivs, EM_Balance_Err_Cancel, .TextMatrix(.Row, .ColIndex("结算ID"))) Then Call RefreshData
            End With
        Case conMenu_Edit_ViewBalance
            Call ViewBalance(tabMain.Selected.Index)
        Case conMenu_Edit_PrintAmend
            Call PrintBill(2)
        Case conMenu_Edit_ReprintBalanceReceipt
            Call PrintBill(1)
        Case conMenu_Edit_DelPrint
            Call PrintDelBill
        Case conMenu_Edit_PrintList '打印结算清单
           Call PrintList
        Case Else
            If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call zlOpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function BalanceErrCancelCheck(ByVal lng结算ID As Long) As Boolean
    '异常单据作废检查
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandler
    '存在非医保结算方式时不允许异常作废，114149
    strSQL = _
        "Select 1" & vbNewLine & _
        "From 病人预交记录 A, 费用补充记录 C, 结算方式 B" & vbNewLine & _
        "Where a.结算序号 = c.结算序号 And a.结算方式 = b.名称 And b.性质 Not In ('3', '4') And c.结算id = [1] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结算ID)
    If Not rsTmp.EOF Then
        MsgBox "本次补充结算已成功结算的结算方式中含有非医保的，因此不允许再作废，只能进行重新结算！", _
            vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    BalanceErrCancelCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CheckErrBill()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim dtStartDate As Date, dtEndDate As Date

    dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
    dtEndDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")

    strSQL = " Select A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, Sum(B.结帐金额), A.操作员姓名, A.登记时间, A.结算序号" & _
             " From 费用补充记录 A, 门诊费用记录 B " & _
             " Where A.登记时间 Between [1] And [2] And Nvl(A.费用状态,0)=1 And A.收费结帐ID=B.结帐ID And A.记录状态 = 2 And A.操作员姓名 = [3]" & _
             " Group By A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, A.操作员姓名, A.登记时间, A.结算序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
    
    If rsTmp.RecordCount <> 0 Then
        tabMain.Item(2).Caption = "异常退费记录(" & rsTmp.RecordCount & ")"
        If MsgBox("存在补充结算异常退费记录,是否处理异常记录?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            tabMain.Item(2).Selected = True
            Call mfrmRefund.ReadData(0)

            Exit Sub
        End If
    Else
        tabMain.Item(2).Caption = "异常退费记录"
    End If
    
    strSQL = " Select A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费') As 类型, B.姓名, B.性别, B.年龄, Sum(B.结帐金额), A.操作员姓名, A.登记时间, A.结算序号" & _
             " From 费用补充记录 A, 门诊费用记录 B " & _
             " Where A.登记时间 Between [1] And [2] And Nvl(A.费用状态,0)=1 And A.收费结帐ID=B.结帐ID And A.记录状态 In (1,3) And A.操作员姓名 = [3]" & _
             "       And Not Exists (Select 1 From 费用补充记录 Where 结算序号=A.结算序号 And 记录状态=2)" & _
             " Group By A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, A.操作员姓名, A.登记时间, A.结算序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
    
    If rsTmp.RecordCount <> 0 Then
        tabMain.Item(1).Caption = "异常结算记录(" & rsTmp.RecordCount & ")"
        If MsgBox("存在补充结算异常结算记录,是否处理异常记录?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            tabMain.Item(1).Selected = True
            Call mfrmErr.ReadData(0)

            Exit Sub
        End If
    Else
        tabMain.Item(1).Caption = "异常结算记录"
    End If

End Sub

Public Sub ViewBalance(intType As Integer)
    Select Case intType
        Case 0
            If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结算序号")) = "" Then Exit Sub
            If Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("退费标志"))) = 2 Then
                frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_查看, mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结算序号")), , , , mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结算时间"))
            Else
                frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_查看, mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("结算序号"))
            End If
        Case 1
            If mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("结算序号")) = "" Then Exit Sub
            frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_查看, mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("结算序号"))
        Case 2
            If mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("结算序号")) = "" Then Exit Sub
            frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_查看, mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("结算序号")), , , , mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("结算时间"))
    End Select
End Sub

Private Sub RefreshData()
    If mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Select Case tabMain.Selected.Index
                Case 0
                    Call mfrmNormal.ReadData(0, mstrPrivs)
                Case 1
                    Call mfrmErr.ReadData(0)
                Case 2
                    Call mfrmRefund.ReadData(0)
            End Select
        End If
    ElseIf mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked Then
        Select Case tabMain.Selected.Index
            Case 0
                Call mfrmNormal.ReadData(0, mstrPrivs)
            Case 1
                Call mfrmErr.ReadData(0)
            Case 2
                Call mfrmRefund.ReadData(0)
        End Select
    End If
End Sub

Private Sub zlOpenReport(ByVal lngSys As Long, ByVal strReportCode As String, Optional ByVal intType As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode-报表编号
    '     intType-报表操作类型:0-默认,1-直接预览,2-直接打印,3-输出到EXCEL
    '编制:刘尔旋
    '日期:2013-09-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, intType)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, blnCollect As Boolean
    Select Case Control.ID
        Case conMenu_Edit_BalanceDel
            If tabMain.Selected.Index = 0 Then
                Control.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
                '0-无记录,1-收费记录,2-退费记录,3-已被退费的收费记录
                Control.Enabled = mfrmNormal.zlGetFeeState <> 2
            Else
                Control.Visible = False
            End If
        Case conMenu_View_Filter
            If tabMain.Selected.Index = 0 Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
        Case conMenu_View_Refresh
            If tabMain.Selected.Index = 0 Then
                Control.BeginGroup = False
            Else
                Control.BeginGroup = True
            End If
        Case conMenu_Edit_ReBalance
            If tabMain.Selected.Index = 1 Then
                Control.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
                With mfrmErr.vsfMain
                    If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And .Row <> 0 Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_BalanceCancel
            If tabMain.Selected.Index = 1 Then
                Control.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
                With mfrmErr.vsfMain
                    If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And .Row <> 0 Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_ReDel
            If tabMain.Selected.Index = 2 Then
                Control.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
                With mfrmRefund.vsfMain
                    If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And .Row <> 0 Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_ReprintBalanceReceipt
            If tabMain.Selected.Index = 0 Then
                Control.Visible = zlCheckPrivs(mstrPrivs, "重打票据")
                With mfrmNormal.vsfMain
                    If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("退费标志"))) <> 2 And _
                        .TextMatrix(.Row, .ColIndex("实际票号")) <> "" Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_PrintAmend
            If tabMain.Selected.Index = 0 Then
                Control.Visible = zlCheckPrivs(mstrPrivs, "补打票据")
                With mfrmNormal.vsfMain
                    If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("退费标志"))) <> 2 And _
                        .TextMatrix(.Row, .ColIndex("实际票号")) = "" Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_DelPrint
            If tabMain.Selected.Index = 0 Then
                With mfrmNormal.vsfMain
                    If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("退费标志"))) = 2 Then
                        If Val(.TextMatrix(.Row, .ColIndex("红票已打印"))) = 1 Then
                            Control.Enabled = zlCheckPrivs(mstrPrivs, "重打票据")
                            Control.Caption = "重打退费票据" & IIf(InStr(Control.Caption, "(") > 0, "(&B)", "")
                        Else
                            Control.Enabled = zlCheckPrivs(mstrPrivs, "补打票据")
                            Control.Caption = "补打退费票据" & IIf(InStr(Control.Caption, "(") > 0, "(&B)", "")
                        End If
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_PrintList
            If tabMain.Selected.Index = 0 Then
                Control.Visible = zlCheckPrivs(mstrPrivs, "门诊结算清单")
                With mfrmNormal.vsfMain
                    If .TextMatrix(.Row, .ColIndex("结算单号")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("退费标志"))) <> 2 Then
                        Control.Enabled = True
                    Else
                        Control.Enabled = False
                    End If
                End With
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_ViewBalance
            Select Case tabMain.Selected.Index
                Case 0
                    With mfrmNormal.vsfMain
                        If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And .Row <> 0 Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End With
                Case 1
                    With mfrmErr.vsfMain
                        If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And .Row <> 0 Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End With
                Case 2
                    With mfrmRefund.vsfMain
                        If .TextMatrix(.Row, .ColIndex("结算序号")) <> "" And .Row <> 0 Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End With
            End Select
    End Select
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
    If mblnFirst Then
        mblnFirst = False
        Call CheckErrBill
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mlngModule = glngModul
    mstrPrivs = gstrPrivs
    mstrPrivsRollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mblnCancel = False
    mstrTitle = "保险补充结算管理"
    '打印部件初始化
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    Call mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Set mobjFactProperty = New clsFactProperty
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_收费收据, 0, 0, 0, mobjFactProperty)
    Call zlDefCommandBars
    '创建TAB信息
    Call SetTabControl
    Call InitIDKind
    Call SetCboDate
    stbThis.Panels(3).Text = UserInfo.姓名
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
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
    Err = 0: On Error GoTo Errhand
    If tabMain.Selected.Index = 0 Then
        Call mfrmNormal.zlRptPrint(bytFunc)
    ElseIf tabMain.Selected.Index = 1 Then
        Call mfrmErr.zlRptPrint(bytFunc)
    ElseIf tabMain.Selected.Index = 2 Then
        Call mfrmRefund.zlRptPrint(bytFunc)
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
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
    
    Err = 0: On Error GoTo Errhand:
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "收费员扎帐(&M)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlCheckPrivs(mstrPrivsRollingCurtain, "轧帐")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_SetInsure, "保险类别(&I)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)"): mcbrControl.BeginGroup = True
'        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "参数设置")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RegistBalance, "挂号结算(&J)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3019
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InsureBalance, "收费结算(&S)")
        mcbrControl.IconId = 3011
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceDel, "结算退费(&U)")
        mcbrControl.IconId = 3017
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBalance, "重新结算(&J)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3831
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceCancel, "结算作废(&C)")
        mcbrControl.IconId = 3832
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReDel, "重新退费(&D)")
        mcbrControl.IconId = 228
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ViewBalance, "查看单据(&V)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintBalanceReceipt, "重打结算票据(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "重打票据")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "补打结算票据(&R)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "补打票据")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DelPrint, "补打退费票据(&B)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "补打票据") Or zlCheckPrivs(mstrPrivs, "重打票据")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintList, "打印收费清单(&L)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "门诊结算清单")
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
        intPara = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModule, "0"))
        Set mcbrRefresh = .Add(xtpControlPopup, conMenu_View_RefreshType, "刷新方式(&O)"): mcbrControl.BeginGroup = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_No, "操作后不刷新数据(&1)", -1, False)
        If intPara = 0 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Ask, "操作后提示刷新数据(&2)", -1, False)
        If intPara = 1 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Auto, "操作后自动刷新数据(&3)", -1, False)
        If intPara = 2 Then cbrControl.Checked = True
        mcbrRefresh.Visible = zlCheckPrivs(mstrPrivs, "参数设置")
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "收费员轧帐"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlCheckPrivs(mstrPrivsRollingCurtain, "轧帐")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
'        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "参数设置")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RegistBalance, "挂号结算"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3019
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InsureBalance, "收费结算")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceDel, "结算退费")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
        mcbrControl.IconId = 3017
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBalance, "重新结算")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3831
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceCancel, "结算作废")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3832
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReDel, "重新退费")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ViewBalance, "查阅单据"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintBalanceReceipt, "重打结算票据"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "重打票据")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "补打结算票据")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "补打票据")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DelPrint, "补打退费票据")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "补打票据") Or zlCheckPrivs(mstrPrivs, "重打票据")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintList, "打印收费清单")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "门诊结算清单")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add 0, VK_F11, conMenu_File_FeeCollect
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F4, conMenu_Edit_RegistBalance
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_DELETE, conMenu_Edit_BalanceDel
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RegistBalance, "挂号结算"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3019
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InsureBalance, "收费结算")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceDel, "结算退费")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
        mcbrControl.IconId = 3017
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBalance, "重新结算")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3831
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceCancel, "结算作废")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "医保结算")
        mcbrControl.IconId = 3832
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReDel, "重新退费")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "结算退费")
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "收费员轧帐"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlCheckPrivs(mstrPrivsRollingCurtain, "轧帐")
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
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tabMain.Width = Me.Width - 225
    picCons.Left = 4500
    If cbsThis(2).Visible Then
        If cbsThis.Options.LargeIcons Then
            tabMain.Top = 900
            picCons.Top = 915
        Else
            tabMain.Top = 780
            picCons.Top = 795
        End If
    Else
        tabMain.Top = 400
        picCons.Top = 415
    End If
    IDKind.Top = 30
    txtIdentify.Top = 30
    IDKind.Left = Me.Width - 3105
    txtIdentify.Left = IDKind.Left + IDKind.Width
    
    '根据状态栏调整界面
    If stbThis.Visible Then
        tabMain.Height = Me.Height - 910 - tabMain.Top
    Else
        tabMain.Height = Me.Height - 910 - tabMain.Top + stbThis.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If Not mfrmNormal Is Nothing Then Unload mfrmNormal: Set mfrmNormal = Nothing
    If Not mfrmErr Is Nothing Then Unload mfrmErr: Set mfrmErr = Nothing
    If Not mfrmRefund Is Nothing Then Unload mfrmRefund: Set mfrmRefund = Nothing
    
    '存储列表的个性化设置(本地)
    
    
    SaveWinState Me, App.ProductName, mstrTitle
    '卸载加载窗体和类
    Set mrsInfo = Nothing
End Sub

Private Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的权限是否存在
    '参数:strPrivs-权限串
    '     strMyPriv-具体权限
    '返回,存在权限,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 14:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCheckPrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

Private Sub cboDate_Click()
    Dim dtStartDate As Date, dtEndDate As Date
    lblSplit.Visible = cboDate.ListIndex = 5
    dtpStartDate.Visible = cboDate.ListIndex = 5
    dtpEndDate.Visible = cboDate.ListIndex = 5
    lblDateShow.Visible = cboDate.ListIndex <> 5
    Select Case cboDate.ListIndex
        Case 0 '今日
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
        Case 1 '最近两天
            dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 2 '最近三天
            dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 3  '最近一周
            dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 4  '本月
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-01") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case Else
            dtStartDate = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    lblDateShow.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
    lblDateShow.Caption = lblDateShow.Caption & "~" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
    If cboDate.Visible = False Then Exit Sub
    If tabMain.Selected.Index = 1 Then
        Call mfrmErr.ReadData(0)
    Else
        Call mfrmRefund.ReadData(0)
    End If
End Sub

Private Sub SetCboDate()
    Dim i As Integer
    i = Val(zlDatabase.GetPara("异常单据查询", glngSys, mlngModule, 0, Array(lbl缺省, cboDate)))
    With cboDate
        .Clear
        .AddItem "今日"
        .ListIndex = .NewIndex
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "最近两天"
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "最近三天"
        If i = 2 Then .ListIndex = .NewIndex
        .AddItem "最近一周"
        If i = 3 Then .ListIndex = .NewIndex
        .AddItem "本月"
        If i = 4 Then .ListIndex = .NewIndex
        .AddItem "自定义"
        If i = 5 Then .ListIndex = .NewIndex
        dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpEndDate.MaxDate = dtpStartDate.MaxDate
        dtpEndDate.Value = dtpEndDate.MaxDate
        dtpStartDate.Value = DateAdd("d", -7, dtpEndDate.MaxDate)
    End With
    Call cboDate_Click
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtIdentify.Locked Then Exit Sub
    txtIdentify.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, txtIdentify.Text)
End Sub

Private Sub tabMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
        Case 0
            picCons.Visible = False
            If mblnFirst Then Exit Sub
            Call mfrmNormal.ReadData(0, mstrPrivs)
        Case 1
            picCons.Visible = True
            Call mfrmErr.ReadData(0)
        Case 2
            picCons.Visible = True
            Call mfrmRefund.ReadData(0)
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
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtIdentify.Locked Then Exit Sub
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "姓名") Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtIdentify.Text, 1)) > 0 And IsNumeric(Mid(txtIdentify.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtIdentify, KeyAscii, IDKind.ShowPassText)
        End If
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
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtIdentify)
    IDKind.AllowAutoCommCard = True
    IDKind.AllowAutoICCard = True
    IDKind.AllowAutoIDCard = True
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
        Call mfrmNormal.ReadData(1, mstrPrivs, Val(Nvl(mrsInfo!ID)))
    Case 1
        Call mfrmErr.ReadData(1, Val(Nvl(mrsInfo!ID)))
    Case 2
        Call mfrmRefund.ReadData(1, Val(Nvl(mrsInfo!ID)))
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
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    
    strSQL = ""
    If blnCard And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then    '103563
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        If lng病人ID <= 0 Then lng病人ID = 0
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And B.门诊号=[2]" & str非在院
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And B.病人ID=[2]" & str非在院
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strSQL = strSQL & " And B.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])" & str非在院
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                '姓名
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtIdentify.Text = mrsInfo!姓名 Then blnSame = True
                End If
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                        txtIdentify.Text = ""
                        Set mrsInfo = Nothing: Exit Function
                    Else
                       'strSQL = strSQL & " And  B.姓名 Like [3]"
                       '问题号:50485
                        strPati = _
                            " Select /*+Rule */distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位,decode(b.卡号,Null,Null,'√') As 是否有医疗卡" & _
                            " From 病人信息 A, 病人医疗卡信息 B " & _
                            " Where Rownum <101 And a.病人ID=b.病人ID(+) And b.状态(+)=0 And B.卡类别ID(+)=[3]  And A.停用时间 is NULL And A.姓名 Like [1]" & str非在院 & _
                            IIf(gintNameDays = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                            
                        vRect = zlControl.GetControlRect(txtIdentify.hWnd)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtIdentify.Height, blnCancel, False, True, strInput & "%", gintNameDays, Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, 0)))
                        If Not rsTmp Is Nothing Then
                            If rsTmp!ID = 0 Then
                                Set mrsInfo = Nothing: Exit Function
                            Else
                                strInput = "-" & rsTmp!病人ID
                                strSQL = strSQL & " And B.病人ID=[2]"
                            End If
                        Else '取消选择
                            txtIdentify.Text = ""
                            Set mrsInfo = Nothing: Exit Function
                        End If
                    End If
                Else
                    strSQL = strSQL & " And B.病人ID=[2]"
                    strInput = "-" & Val(mrsInfo!病人ID)
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And B.医保号=[1]" & str非在院
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                ' strSQL = strSQL & " And B.身份证号=[1] " & str非在院
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.门诊号=[1]" & str非在院
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])" & str非在院
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
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSQL
    strSQL = "    " & vbNewLine & " Select /*+Rule */distinct  B.病人id As ID, Decode(sign(nvl(ylkxx.病人id,0)),0,'','√') as 三方账户, B.病人id,B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位,"
    strSQL = strSQL & vbNewLine & "      A.名称 险类名称"
    strSQL = strSQL & vbNewLine & " From 病人信息 B, 保险类别 A,医疗卡类别 YLK,病人医疗卡信息 YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.险类 = A.序号(+) and b.病人id=ylkxx.病人id(+) and ylkxx.状态(+)=0 and  ylkxx.卡类别id=ylk.id(+)  and ylk.是否自制(+)=0 And B.停用时间 Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
    
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
        
    If mrsInfo Is Nothing Then GoTo ClearPati:
    If mrsInfo.State <> 1 Then GoTo ClearPati:
    If mrsInfo.RecordCount = 0 Then GoTo ClearPati:
    If Val(Nvl(mrsInfo!ID)) = 0 Then GoTo ClearPati:
    
    txtIdentify.Text = Nvl(mrsInfo!姓名)
    Me.txtIdentify.Tag = Nvl(mrsInfo!ID)
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

Private Function GetBalanceInsure(ByVal str结算序号 As String, _
    Optional ByRef str险类名称 As String, Optional ByRef lng病人ID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保险序事情
    '出参:str险类名称-险类名称
    '     lng病人ID-病人ID
    '返回:返回险类
    '编制:刘兴洪
    '日期:2014-09-22 13:57:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
    strSQL = "" & _
        "   Select /*+ rule */ b.记录id, b.险类, b.病人id, c.名称" & _
        "   From 费用补充记录 A, 保险结算记录 B, 保险类别 C" & _
        "   Where a.结算id = b.记录id And b.险类 = c.序号(+) And b.性质 = 1 And a.结算序号 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结算序号)
    If Not rsTmp.EOF Then
        lng病人ID = Nvl(rsTmp!病人ID, 0)
        str险类名称 = Nvl(rsTmp!名称)
        GetBalanceInsure = Nvl(rsTmp!险类, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 14:15:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.姓名, EM_收费收据, mobjFactProperty.使用类别, lng领用ID, mobjFactProperty.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mobjFactProperty.使用类别) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mobjFactProperty.使用类别 & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mobjFactProperty.使用类别) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mobjFactProperty.使用类别 & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新收费票据号
    '编制:刘兴洪
    '日期:2014-06-06 14:21:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjFactProperty Is Nothing Then Exit Sub
    If mobjFactProperty.打印方式 = 0 Then Exit Sub
    
    If mobjFactProperty.严格控制 Then
            
        If zlGetInvoiceGroupUseID(mlng领用ID) = False Then
            mstrInvoice = "": Exit Sub
        End If
        '严格：取下一个号码
        If mobjInvoice.zlGetNextBill(mlngModule, mlng领用ID, strFactNO) = False Then strFactNO = ""
        mstrInvoice = strFactNO
        
    Else
        '松散：取下一个号码
        mstrInvoice = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModule)))
    End If
End Sub

Private Function GetFeeNos(ByVal strNo As String) As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, strResult As String
    strSQL = _
        " Select Distinct NO" & vbNewLine & _
        " From 门诊费用记录" & vbNewLine & _
        " Where 结帐id In (Select Distinct 收费结帐id From 费用补充记录 Where NO = [1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    Do While Not rsTmp.EOF
        strResult = strResult & "," & rsTmp!NO
        rsTmp.MoveNext
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 2)
    GetFeeNos = strResult
End Function

Private Sub PrintBill(bytType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印
    '入参:bytType-操作类型 1:重打票据 2:补打票据
    '编制:刘尔旋
    '日期:2014-09-24 17:33:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVirtualPrint As Boolean, lng病人ID As Long
    Dim intPrint As Integer, dtDate As Date, strNos As String, str类别 As String
    Dim intInsure As Integer, i As Integer, strNo As String
    If tabMain.Selected.Index <> 0 Then Exit Sub
    If bytType = 2 Then
        With mfrmNormal.vsfInvoice
            If .TextMatrix(1, 1) <> "" Then
                MsgBox "选择的记录已经打印过票据,不能进行补打！", vbInformation, gstrSysName
                Exit Sub
            End If
        End With
    End If
    With mfrmNormal.vsfMain
        strNo = .TextMatrix(.Row, .ColIndex("结算单号"))
        strNos = GetFeeNos(strNo)
        str类别 = .TextMatrix(.Row, .ColIndex("类型"))
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        intInsure = GetBalanceInsure(Val(.TextMatrix(.Row, .ColIndex("结算序号"))))
        Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_收费收据, lng病人ID, 0, intInsure, mobjFactProperty)
        dtDate = .TextMatrix(.Row, .ColIndex("结算时间"))
        If strNo = "" Then Exit Sub
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(Val(.TextMatrix(.Row, .ColIndex("结算序号")))))
        If Not blnVirtualPrint And strNos <> "" Then
            If str类别 = "收费" Then
                If Not BillExistMoney(strNos, 1) Then
                    MsgBox "选择的记录已经全部退费,不能进行打印！", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                If Not BillExistMoney(strNos, 4) Then
                    MsgBox "选择的记录已经全部退费,不能进行打印！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        If zlRePrintReplenishTheBalanceBill(Me, mlngModule, bytType, strNo, intInsure, mobjInvoice, mobjFactProperty, , , blnVirtualPrint) Then
            Call RefreshData
        End If
    End With
End Sub

Private Sub PrintDelBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:票据打印
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVirtualPrint As Boolean, lng病人ID As Long
    Dim intPrint As Integer, strNos As String, str类别 As String
    Dim intInsure As Integer, i As Integer, lng结算序号 As Long
    
    Err = 0: On Error GoTo Errhand
    If tabMain.Selected.Index <> 0 Then Exit Sub
    With mfrmNormal.vsfMain
        lng结算序号 = Val(.TextMatrix(.Row, .ColIndex("结算序号")))
        If lng结算序号 = 0 Then Exit Sub
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        intInsure = GetBalanceInsure(lng结算序号)
        
        Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_退费收据, lng病人ID, 0, intInsure, mobjFactProperty)
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(Val(.TextMatrix(.Row, .ColIndex("结算序号")))))
        
        If zlPrintReplenishTheDelBalanceBill(Me, mlngModule, lng结算序号, intInsure, mobjInvoice, mobjFactProperty, , zlDatabase.Currentdate, blnVirtualPrint) Then
            Call RefreshData
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtIdentify_LostFocus()
    IDKind.SetAutoReadCard False
End Sub

Private Sub PrintList()
    '打印收费结算清单
    Dim strNo As String
    
    On Error GoTo Errhand
    With mfrmNormal
        strNo = .vsfMain.TextMatrix(.vsfMain.Row, .vsfMain.ColIndex("结算单号"))
        If strNo = "" Then
            MsgBox "当前没有单据可以打印清单！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '是否已转入后备数据表中
        If .mblnNOMoved Then
            If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
            .mblnNOMoved = False  '此时已转入在线数据表
        End If
    End With
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO='" & strNo & "'", "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
    End If
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
