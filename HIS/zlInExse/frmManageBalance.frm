VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBalance 
   Caption         =   "���˽��ʹ���"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmManageBalance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
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
      Begin VB.Label lblȱʡ 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡ��ʾ"
         Height          =   180
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblSplit 
         Caption         =   "��"
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
      IDKindStr       =   "��|��������￨|0|0|0|0|0|0;��|�����|0|0|0|0|0|0;ס|סԺ��|0|0|0|0|0|0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
Private mblnCancel As Boolean   '�ⲿж�ش����ʶ
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����
Private mrsInfo As ADODB.Recordset, mstrPrivsRollingCurtain As String
Private mblnFirst As Boolean, mstrWriteCardTypeIDs As String
Private mobjInPati As Object, mbln�������� As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mbln��ͳģʽ As Boolean

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnOk As Boolean
    On Error GoTo errHandle
    Select Case Control.ID
        Case conMenu_File_FeeCollect
            If zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "����") = False Then Exit Sub
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
            If zlStr.IsHavePrivs(mstrPrivs, "��������") = False Then Exit Sub
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
            zlDatabase.SetPara "ˢ�·�ʽ", "0", glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "��������")
        Case conMenu_View_RefreshType_Ask
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked = False
            zlDatabase.SetPara "ˢ�·�ʽ", "1", glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "��������")
        Case conMenu_View_RefreshType_Auto
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            zlDatabase.SetPara "ˢ�·�ʽ", "2", glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "��������")
        Case conMenu_Edit_ClinicBalance
            If zlStr.IsHavePrivs(mstrPrivs, "������ý���") = False Then Exit Sub
            '������ý���
            If mbln��ͳģʽ Then
                blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_�������, mstrPrivs)
            Else
                blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_�������, mstrPrivs)
            End If
            If blnOk Then Call RefreshData
        Case conMenu_Edit_InHosBalance
            If zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") = False Then Exit Sub
            'סԺ���ý���
            If mbln��ͳģʽ Then
                blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_סԺ����, mstrPrivs)
            Else
                blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_סԺ����, mstrPrivs)
            End If
            If blnOk Then Call RefreshData
        Case conMenu_Edit_ErrReBalance
            '�쳣�ؽ�
            With mfrmErr.vsfMain
                If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
                If mbln��ͳģʽ Then
                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_���½���, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")))
                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_���½���, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")))
                End If
                If blnOk Then Call RefreshData
            End With
        Case conMenu_Edit_ErrCancelBalance
            '�쳣����
            With mfrmErr.vsfMain
                If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
'                If mbln��ͳģʽ Then
'                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_ȡ������, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")))
'                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_ȡ������, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")))
'                End If
                If blnOk Then Call RefreshData
            End With
        Case conMenu_Edit_ErrDelBalance
            '�쳣����
            With mfrmRefund.vsfMain
                If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
'                If mbln��ͳģʽ Then
'                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_��������, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")), True)
'                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_��������, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")), True)
'                End If
                If blnOk Then Call RefreshData
                
            End With
        Case conMenu_Edit_CancelBalance
            '��������
            With mfrmNormal.vsfMain
                If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
'                If mbln��ͳģʽ Then
'                    blnOk = frmPatiBalanceTraditional.ShowMe(Me, g_Ed_��������, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")))
'                Else
                    blnOk = frmPatiBalanceSplit.ShowMe(Me, g_Ed_��������, mstrPrivs, , , .TextMatrix(.Row, .ColIndex("���ݺ�")))
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
            If InStr(1, mstrPrivs, ";�������תסԺ;") = 0 Then Exit Sub
            If mobjInPati Is Nothing Then
                Err = 0: On Error Resume Next
                Set mobjInPati = CreateObject("zl9InPatient.clsInPatient")
                
                If Err <> 0 Then
                    MsgBox "ע��:" & vbCrLf & "   סԺ���˲���(zl9InPatient)����ʧ��,����ϵͳ����Ա��ϵ!"
                    Exit Sub
                End If
            End If
            Call mobjInPati.zlOutFeeToInFee(Me, gcnOracle, glngSys, mlngModule, mstrPrivs, gstrDBUser, 0, 0)
        Case conMenu_Edit_ToHosCancel
            If InStr(mstrPrivs, ";תסԺ��������;") = 0 Or mbln�������� Then Exit Sub
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
            '�����˲���Ʊ��
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
    '����:����˿�
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
    Dim lngCardTypeID As Long, strExpend As String, lng����ID As Long
    Dim lng����ID As Long, strNO As String, lng��¼״̬ As Long
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim bytFunc As Byte
    
    With mfrmNormal.vsfMain
        strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        bytFunc = IIf(Val(.TextMatrix(.Row, .ColIndex("��־"))) = 1, 0, 1)
    End With
    '����:��סԺ��Ϣд�뿨��
    '����:56615
    If mstrWriteCardTypeIDs = "" Then Exit Sub
    If bytFunc = 0 Then '������ʷ���
        If Not zlStr.IsHavePrivs(mstrPrivs, "������Ϣд��") Then Exit Sub
    Else
        If Not zlStr.IsHavePrivs(mstrPrivs, "סԺ��Ϣд��") Then Exit Sub
    End If
    
     If strNO = "" Then
        MsgBox "��ǰû�е��ݿ�������д����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If lng����ID = 0 Or lng����ID = 0 Then Exit Sub
    If InStr(1, mstrWriteCardTypeIDs, ",") = 0 Then lngCardTypeID = Val(mstrWriteCardTypeIDs)
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, gobjSquare.objSquareCard, lngCardTypeID, bytFunc, lng����ID, lng����ID)
End Sub

Private Sub CheckErrBill()
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim dtStartDate As Date, dtEndDate As Date

    dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
    dtEndDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")

    strSql = " Select Count(1) As ��¼��" & vbNewLine & _
             " From ���˽��ʼ�¼" & vbNewLine & _
             " Where �շ�ʱ�� Between [1] And [2] And Nvl(����״̬, 0) = 1 And ��¼״̬ = 2 And ����Ա���� = [3]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
    
    If rsTmp.RecordCount <> 0 Then
        If Val(NVL(rsTmp!��¼��)) <> 0 Then
            tabMain.Item(2).Caption = "�쳣�˷Ѽ�¼(" & Val(NVL(rsTmp!��¼��)) & ")"
            If MsgBox("���ڽ����쳣�˷Ѽ�¼,�Ƿ����쳣��¼?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tabMain.Item(2).Selected = True
                Call mfrmRefund.ReadData
                Exit Sub
            End If
        End If
    Else
        tabMain.Item(2).Caption = "�쳣�˷Ѽ�¼"
    End If
    
    strSql = " Select Count(1) As ��¼��" & vbNewLine & _
             " From ���˽��ʼ�¼ A" & vbNewLine & _
             " Where a.�շ�ʱ�� Between [1] And [2] And Nvl(a.����״̬, 0) = 1 And a.��¼״̬ In (1, 3) And a.����Ա���� = [3] And" & vbNewLine & _
             "       Not Exists (Select 1 From ���˽��ʼ�¼ Where NO = a.No And ��¼״̬ = 2)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
    
    If rsTmp.RecordCount <> 0 Then
        If Val(NVL(rsTmp!��¼��)) <> 0 Then
            tabMain.Item(1).Caption = "�쳣�����¼(" & Val(NVL(rsTmp!��¼��)) & ")"
            If MsgBox("���ڽ����쳣�����¼,�Ƿ����쳣��¼?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tabMain.Item(1).Selected = True
                Call mfrmErr.ReadData
                Exit Sub
            End If
        End If
    Else
        tabMain.Item(1).Caption = "�쳣�����¼"
    End If

End Sub

Public Sub ViewBalance(intTYPE As Integer)
    '���ĵ���
    Select Case intTYPE
        Case 0
            If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("����ID")) = "" Then Exit Sub
            frmPatiBalanceSplit.ShowMe Me, g_Ed_���ݲ鿴, mstrPrivs, , , _
                                       mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("���ݺ�")), _
                                       Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("��¼״̬"))) = 2, _
                                       zlDatabase.NOMoved("���˽��ʼ�¼", mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("���ݺ�")))
        Case 1
            If mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("����ID")) = "" Then Exit Sub
            frmPatiBalanceSplit.ShowMe Me, g_Ed_���ݲ鿴, mstrPrivs, , , mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("���ݺ�")), False
        Case 2
            If mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("����ID")) = "" Then Exit Sub
            frmPatiBalanceSplit.ShowMe Me, g_Ed_���ݲ鿴, mstrPrivs, , , mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("���ݺ�")), True
    End Select
End Sub

Private Sub RefreshData()
    If mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked Then
        If MsgBox("��ǰ�ѷ�������,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode-������
    '     intType-�����������:0-Ĭ��,1-ֱ��Ԥ��,2-ֱ�Ӵ�ӡ,3-�����EXCEL
    '����:������
    '����:2013-09-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String
    Select Case tabMain.Selected.Index
    Case 0
        With mfrmNormal.vsfMain
            strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
            If strNO = "" Then
                Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, intTYPE)
            Else
                Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, _
                "����ID=" & .TextMatrix(.Row, .ColIndex("����ID")), _
                "סԺ��=" & .TextMatrix(.Row, .ColIndex("סԺ��")), _
                "����ID=" & .TextMatrix(.Row, .ColIndex("����ID")), _
                "NO=" & strNO, _
                "��¼״̬=" & .TextMatrix(.Row, .ColIndex("��¼״̬")), intTYPE)
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
        '��ͨ���ʼ�¼��ť����
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_Edit_View
                Control.Enabled = mfrmNormal.vsfMain.TextMatrix(1, mfrmNormal.vsfMain.ColIndex("����ID")) <> ""
            Case conMenu_Edit_ErrReBalance, conMenu_Edit_ErrCancelBalance, conMenu_Edit_ErrDelBalance
                Control.Enabled = False
                Control.Visible = False
            Case conMenu_Edit_CancelBalance
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
                If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("����ID")) <> "" Then
                    Control.Enabled = Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("��¼״̬"))) = 1
                    If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("ҽ��")) <> "" Then
                        If zlStr.IsHavePrivs(mstrPrivs, "���ս���") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���") = False Then Control.Enabled = False
                    End If
                Else
                    Control.Enabled = False
                End If
            Case conMenu_Edit_ReprintReceipt
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ش�Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
                If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("����ID")) <> "" Then
                    Control.Enabled = Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("��¼״̬"))) = 1 _
                                    And mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("Ʊ�ݺ�")) <> ""
                    If InStr(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("  ��������")), "����") > 0 Then
                        If zlStr.IsHavePrivs(mstrPrivs, "������ý���") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") = False Then Control.Enabled = False
                    End If
                    If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("ҽ��")) <> "" Then
                        If zlStr.IsHavePrivs(mstrPrivs, "���ս���") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���") = False Then Control.Enabled = False
                    End If
                Else
                    Control.Enabled = False
                End If
            Case conMenu_Edit_PrintAmend
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
                If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("����ID")) <> "" Then
                    Control.Enabled = Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("��¼״̬"))) = 1 _
                                    And mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("Ʊ�ݺ�")) = ""
                    If InStr(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("  ��������")), "����") > 0 Then
                        If zlStr.IsHavePrivs(mstrPrivs, "������ý���") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") = False Then Control.Enabled = False
                    End If
                    If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("ҽ��")) <> "" Then
                        If zlStr.IsHavePrivs(mstrPrivs, "���ս���") = False Then Control.Enabled = False
                    Else
                        If zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���") = False Then Control.Enabled = False
                    End If
                Else
                    Control.Enabled = False
                End If
            Case conMenu_Edit_WriteCard
                bytFunc = IIf(Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("��־"))) = 1, 0, 1)
                Control.Visible = (zlStr.IsHavePrivs(mstrPrivs, "סԺ��Ϣд��") Or zlStr.IsHavePrivs(mstrPrivs, "������Ϣд��")) _
                                And mstrWriteCardTypeIDs <> ""
                Control.Enabled = (bytFunc = 0 And zlStr.IsHavePrivs(mstrPrivs, "������Ϣд��")) _
                        Or (bytFunc = 1 And zlStr.IsHavePrivs(mstrPrivs, "סԺ��Ϣд��")) _
                        And mfrmNormal.vsfMain.TextMatrix(1, mfrmNormal.vsfMain.ColIndex("����ID")) <> ""
            Case conMenu_View_Filter
                Control.Visible = True
                Control.Enabled = True
        End Select
        IDKind.Visible = True
        txtIdentify.Visible = True
    ElseIf tabMain.Selected.Index = 1 Then
        '�쳣�����¼��ť����
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_Edit_View
                Control.Enabled = mfrmErr.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("����ID")) <> ""
            Case conMenu_Edit_ErrReBalance, conMenu_Edit_ErrCancelBalance
                Control.Visible = True
                Control.Enabled = mfrmErr.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("����ID")) <> ""
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
        '�쳣���ϼ�¼��ť����
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_Edit_View
                Control.Enabled = mfrmErr.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("����ID")) <> ""
            Case conMenu_Edit_ErrReBalance, conMenu_Edit_ErrCancelBalance, conMenu_Edit_CancelBalance
                Control.Visible = False
                Control.Enabled = False
            Case conMenu_Edit_ErrDelBalance
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
                Control.Enabled = mfrmRefund.vsfMain.TextMatrix(1, mfrmErr.vsfMain.ColIndex("����ID")) <> ""
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
    '����:�ⲿ�������ж�ش���,��ֵ������FORMLOAD���˳�
    '����:������
    '����:2013-10-11
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    mblnCancel = True
End Sub

Private Sub SetTabControl()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����TAB�ؼ�
    '����:������
    '����:2013-09-04
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With tabMain
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        .InsertItem 1, "���������¼", mfrmNormal.hWnd, 0
        .InsertItem 2, "�쳣�����¼", mfrmErr.hWnd, 0
        .InsertItem 3, "�쳣�˷Ѽ�¼", mfrmRefund.hWnd, 0
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
    mstrTitle = "���˽��ʹ���"
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
    mbln��ͳģʽ = Val(zlDatabase.GetPara("���ʽ�����", glngSys, mlngModule, "1")) = 0
    mstrWriteCardTypeIDs = ""
    If Not gobjSquare Is Nothing Then
        If Not gobjSquare.objSquareCard Is Nothing Then
            mstrWriteCardTypeIDs = gobjSquare.objSquareCard.zlGetAvailabilityWriteCardType
        End If
    End If
    Call zlDefCommandBars
    '����TAB��Ϣ
    Call SetTabControl
    Call InitIDKind
    Call SetCboDate
    stbThis.Panels(3).Text = UserInfo.����
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
    '���������˰�ش�ӡ����
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(2)
        End If
        On Error GoTo 0
    End If
    '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.���, UserInfo.����)
    End If
    On Error GoTo 0
End Sub

Public Sub ShowPopup()
    mcbrPopupMain.ShowPopup
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:������
    '����:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    If mfrmNormal.ActiveControl Is mfrmNormal.vsfMain And tabMain.Selected.Index = 0 Then
        With mfrmNormal.vsfMain
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("����ID"))) = 0 Then Exit Sub
        End With
        Call mfrmNormal.zlRptPrint(bytFunc)
    End If
    
    If mfrmErr.ActiveControl Is mfrmErr.vsfMain And tabMain.Selected.Index = 1 Then
        With mfrmErr.vsfMain
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("����ID"))) = 0 Then Exit Sub
        End With
        Call mfrmErr.zlRptPrint(bytFunc)
    End If
    
    If mfrmRefund.ActiveControl Is mfrmRefund.vsfMain And tabMain.Selected.Index = 2 Then
        With mfrmRefund.vsfMain
            If .Rows = 1 Then Exit Sub
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("����ID"))) = 0 Then Exit Sub
        End With
        Call mfrmRefund.zlRptPrint(bytFunc)
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:������
    '����:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim intPara As Integer
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    '��ʼ������
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&U)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_CashCount, "�ֽ�㳮(&D)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "�շ�����(&M)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_SetInsure, "�������(&I)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicBalance, "�������(&M)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InHosBalance, "סԺ����(&A)")
        mcbrControl.IconId = 3590
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BatchBalance, "������;����(&T)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "������;����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_UnitBalance, "��Լ��λ����(&U)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_FeeManage, "Ӧ�տ����(&Y)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "Ӧ�տ����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RefundDeposit, "����˿�(&R)")
        mcbrControl.IconId = 3017
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicToHos, "�������תסԺ(&Z)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������תסԺ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ToHosCancel, "תסԺ��������(&X)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "תסԺ��������") And Not mbln��������
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "�쳣�ؽ�(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "�쳣����(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrDelBalance, "�쳣����(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelBalance, "��������(&D)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 4114
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_View, "���ĵ���(&V)")
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "�ش����Ʊ��(&R)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ش�Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "�������Ʊ��(&B)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "����Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintDetail, "��ӡ������ϸ(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmendByPati, "�����˲������Ʊ��(&P)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "����Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_WriteCard, "������Ϣд��(&W)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = (zlStr.IsHavePrivs(mstrPrivs, "סԺ��Ϣд��") Or zlStr.IsHavePrivs(mstrPrivs, "������Ϣд��")) _
                                And mstrWriteCardTypeIDs <> ""
    End With
    
    Set mcbrMenuView = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuView.ID = conMenu_ViewPopup
    With mcbrMenuView.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrControl.Checked = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����(&F)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Location, "��λ(&G)")
        intPara = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModule, "0"))
        Set mcbrRefresh = .Add(xtpControlPopup, conMenu_View_RefreshType, "ˢ�·�ʽ(&O)"): mcbrControl.BeginGroup = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_No, "������ˢ������(&1)", -1, False)
        If intPara = 0 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Ask, "��������ʾˢ������(&2)", -1, False)
        If intPara = 1 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Auto, "�������Զ�ˢ������(&3)", -1, False)
        If intPara = 2 Then cbrControl.Checked = True
        mcbrRefresh.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�����")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&K)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '���������˵�
    Set mcbrPopupMain = cbsThis.Add("�����˵�1", xtpBarPopup)
    With mcbrPopupMain.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����Excel")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "�շ�����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicBalance, "�������(&M)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InHosBalance, "סԺ����(&A)")
        mcbrControl.IconId = 3590
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BatchBalance, "������;����(&T)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "������;����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_UnitBalance, "��Լ��λ����(&U)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_FeeManage, "Ӧ�տ����(&Y)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "Ӧ�տ����")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicToHos, "�������תסԺ(&Z)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������תסԺ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ToHosCancel, "תסԺ��������(&X)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "תסԺ��������") And Not mbln��������
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "�쳣�ؽ�(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "�쳣����(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrDelBalance, "�쳣����(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelBalance, "��������(&D)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 4114
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_View, "���ĵ���(&V)")
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "�ش����Ʊ��(&R)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "�ش�Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "�������Ʊ��(&B)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "����Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintDetail, "��ӡ������ϸ(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmendByPati, "�����˲������Ʊ��(&P)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "����Ʊ��") And _
                            ((zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or _
                              zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))) Or (zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And _
                             (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_WriteCard, "������Ϣд��(&W)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = (zlStr.IsHavePrivs(mstrPrivs, "סԺ��Ϣд��") Or zlStr.IsHavePrivs(mstrPrivs, "������Ϣд��")) _
                                And mstrWriteCardTypeIDs <> ""
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Location, "��λ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
    End With
    
    '�����
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
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ModifyStyle &H400000, 0
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")

        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicBalance, "����"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "������ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InHosBalance, "סԺ")
        mcbrControl.IconId = 3590
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))

        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "�쳣�ؽ�"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "�쳣����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ErrDelBalance, "�쳣����"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelBalance, "����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 4114
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "��������") And (zlStr.IsHavePrivs(mstrPrivs, "���ս���") Or zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���"))
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_View, "����")
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivsRollingCurtain, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): mcbrControl.BeginGroup = True
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
    
    '�洢�б�ĸ��Ի�����(����)
    
    Call SaveRegInFor(g˽��ģ��, Me.Name, "�쳣���ݲ�ѯ", cboDate.ListIndex)
    SaveWinState Me, App.ProductName, mstrTitle
    'ж�ؼ��ش������
    Set mrsInfo = Nothing

    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
    
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, int��Դ As Integer
    Dim blnFill As Boolean
    Dim strNO As String
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    Select Case tabMain.Selected.Index
    Case 0
        For i = IIf(blnHead, 1, mlngGo) To mfrmNormal.vsfMain.Rows - 1
            DoEvents
            
            '�Ƚ�����
            blnFill = True
            With frmBalanceGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("���ݺ�")) = .txtNO.Text
                End If
                If .txtFact.Text <> "" Then
                    blnFill = blnFill And mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("Ʊ�ݺ�")) = .txtFact.Text
                End If
                If .txtסԺ��.Text <> "" Then
                    blnFill = blnFill And mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("סԺ��")) = .txtסԺ��.Text
                End If
                If .txt����.Text <> "" Then
                    blnFill = blnFill And UCase(mfrmNormal.vsfMain.TextMatrix(i, mfrmNormal.vsfMain.ColIndex("����"))) Like "*" & UCase(.txt����.Text) & "*"
                End If
            End With
            
            '�������˳�
            If blnFill Then
                mlngGo = i + 1
                mfrmNormal.vsfMain.Row = i: mfrmNormal.vsfMain.TopRow = i
                mfrmNormal.vsfMain.Col = 0: mfrmNormal.vsfMain.ColSel = mfrmNormal.vsfMain.Cols - 1
                
                stbThis.Panels(2).Text = "�ҵ�һ����¼"
                Screen.MousePointer = 0: Exit Sub
            End If
            
            '��ESCȡ��
            If mblnGo = False Then
                stbThis.Panels(2).Text = "�û�ȡ����λ����"
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    Case 1
        For i = IIf(blnHead, 1, mlngGo) To mfrmErr.vsfMain.Rows - 1
            DoEvents
            
            '�Ƚ�����
            blnFill = True
            With frmBalanceGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("���ݺ�")) = .txtNO.Text
                End If
                If .txtFact.Text <> "" Then
                    blnFill = blnFill And mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("Ʊ�ݺ�")) = .txtFact.Text
                End If
                If .txtסԺ��.Text <> "" Then
                    blnFill = blnFill And mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("סԺ��")) = .txtסԺ��.Text
                End If
                If .txt����.Text <> "" Then
                    blnFill = blnFill And UCase(mfrmErr.vsfMain.TextMatrix(i, mfrmErr.vsfMain.ColIndex("����"))) Like "*" & UCase(.txt����.Text) & "*"
                End If
            End With
            
            '�������˳�
            If blnFill Then
                mlngGo = i + 1
                mfrmErr.vsfMain.Row = i: mfrmErr.vsfMain.TopRow = i
                mfrmErr.vsfMain.Col = 0: mfrmErr.vsfMain.ColSel = mfrmErr.vsfMain.Cols - 1
                
                stbThis.Panels(2).Text = "�ҵ�һ����¼"
                Screen.MousePointer = 0: Exit Sub
            End If
            
            '��ESCȡ��
            If mblnGo = False Then
                stbThis.Panels(2).Text = "�û�ȡ����λ����"
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    Case 2
        For i = IIf(blnHead, 1, mlngGo) To mfrmRefund.vsfMain.Rows - 1
            DoEvents
            
            '�Ƚ�����
            blnFill = True
            With frmBalanceGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("���ݺ�")) = .txtNO.Text
                End If
                If .txtFact.Text <> "" Then
                    blnFill = blnFill And mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("Ʊ�ݺ�")) = .txtFact.Text
                End If
                If .txtסԺ��.Text <> "" Then
                    blnFill = blnFill And mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("סԺ��")) = .txtסԺ��.Text
                End If
                If .txt����.Text <> "" Then
                    blnFill = blnFill And UCase(mfrmRefund.vsfMain.TextMatrix(i, mfrmRefund.vsfMain.ColIndex("����"))) Like "*" & UCase(.txt����.Text) & "*"
                End If
            End With
            
            '�������˳�
            If blnFill Then
                mlngGo = i + 1
                mfrmRefund.vsfMain.Row = i: mfrmRefund.vsfMain.TopRow = i
                mfrmRefund.vsfMain.Col = 0: mfrmRefund.vsfMain.ColSel = mfrmRefund.vsfMain.Cols - 1
                
                stbThis.Panels(2).Text = "�ҵ�һ����¼"
                Screen.MousePointer = 0: Exit Sub
            End If
            
            '��ESCȡ��
            If mblnGo = False Then
                stbThis.Panels(2).Text = "�û�ȡ����λ����"
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    End Select
    
    mlngGo = 1
    stbThis.Panels(2).Text = "�Ѷ�λ���嵥β��"
    Screen.MousePointer = 0
End Sub

Private Sub cboDate_Click()
    Dim dtStartDate As Date, dtEndDate As Date
    lblSplit.Visible = cboDate.ListIndex = 6
    dtpStartDate.Visible = cboDate.ListIndex = 6
    dtpEndDate.Visible = cboDate.ListIndex = 6
    lblDateShow.Visible = cboDate.ListIndex <> 6 And cboDate.ListIndex <> 0
    Select Case cboDate.ListIndex
        Case 0 '�����쳣
            
        Case 1 '����
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
        Case 2 '���2��
            dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 3 '���3��
            dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 4  '���һ��
            dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 5  '����
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
    Call GetRegInFor(g˽��ģ��, Me.Name, "�쳣���ݲ�ѯ", strValue)
    i = Val(strValue)
    With cboDate
        .Clear
        .AddItem "�����쳣���"
        .ListIndex = .NewIndex
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "����"
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "�������"
        If i = 2 Then .ListIndex = .NewIndex
        .AddItem "�������"
        If i = 3 Then .ListIndex = .NewIndex
        .AddItem "���һ��"
        If i = 4 Then .ListIndex = .NewIndex
        .AddItem "����"
        If i = 5 Then .ListIndex = .NewIndex
        .AddItem "�Զ���ʱ�䷶Χ"
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
    txtIdentify.Text = objPatiInfor.����
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
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.����
            Else
                If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
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
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "����") Then
        blnCard = zlCommFun.InputIsCard(txtIdentify, KeyAscii, IDKind.ShowPassText)
    ElseIf IsCardType(IDKind, "�����") Or IsCardType(IDKind, "סԺ��") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtIdentify.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtIdentify.IMEMode = 0
    End If
    If blnCard And Len(txtIdentify.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtIdentify.Text) <> "" Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(txtIdentify.Tag) <> 0 Then    '����
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

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    IDKind.AllowAutoCommCard = True
    IDKind.AllowAutoICCard = True
    IDKind.AllowAutoIDCard = True
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtIdentify)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
End Function

'��ȡĬ��IDKind����
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-09-03 09:32:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not GetPatient(objCard, strInput, blnCard) Then
        MsgBox "δ�ҵ����ˣ����������룡", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case tabMain.Selected.Index
    Case 0
        Call mfrmNormal.ReadData(1, mstrPrivs, Val(NVL(mrsInfo!ID)))
    End Select
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�blnCard=�Ƿ���￨ˢ��
    '���أ����ҳɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:24:14
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    
    strSql = ""
    If blnCard And objCard.���� Like "����*" And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        If lng����ID <= 0 Then lng����ID = 0
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSql = strSql & " And B.����ID=[2] " & str����Ժ
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSql = strSql & " And B.�����=[2]" & str����Ժ
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSql = strSql & " And B.����ID=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSql = strSql & " And B.����ID = (Select Nvl(Max(����ID),0) As ����ID From ������ҳ   Where  סԺ��=[1])"
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                '����
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtIdentify.Text = mrsInfo!���� Then blnSame = True
                End If
                
                If Not blnSame Then
                    'strSQL = strSQL & " And  B.���� Like [3]"
                    '�����:50485
                     strPati = _
                         " Select /*+Rule */distinct 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ,decode(b.����,Null,Null,'��') As �Ƿ���ҽ�ƿ�" & _
                         " From ������Ϣ A, ����ҽ�ƿ���Ϣ B " & _
                         " Where Rownum <101 And a.����ID=b.����ID(+) And b.״̬(+)=0 And B.�����ID(+)=[2]  And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & str����Ժ
                         
                     vRect = zlControl.GetControlRect(txtIdentify.hWnd)
                     Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtIdentify.Height, blnCancel, False, True, strInput & "%", Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, 0)))
                     If Not rsTmp Is Nothing Then
                         If rsTmp!ID = 0 Then
                             Set mrsInfo = Nothing: Exit Function
                         Else
                             strInput = "-" & rsTmp!����ID
                             strSql = strSql & " And B.����ID=[2]"
                         End If
                     Else 'ȡ��ѡ��
                         txtIdentify.Text = ""
                         Set mrsInfo = Nothing: Exit Function
                     End If
                Else
                    strSql = strSql & " And B.����ID=[2]"
                    strInput = "-" & Val(mrsInfo!����ID)
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSql = strSql & " And B.ҽ����=[1]" & str����Ժ
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSql = strSql & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                ' strSQL = strSQL & " And B.���֤��=[1] " & str����Ժ
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSql = strSql & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And B.�����=[1]" & str����Ժ
                '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And B.����ID = (Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[1]) " & str����Ժ
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                    If lng����ID = 0 Then lng����ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID <= 0 Then lng����ID = 0
                strSql = strSql & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSql
    strSql = "    " & vbNewLine & " Select /*+Rule */distinct  B.����id As ID, Decode(sign(nvl(ylkxx.����id,0)),0,'','��') as �����˻�, B.����id,B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ,"
    strSql = strSql & vbNewLine & "      A.���� ��������"
    strSql = strSql & vbNewLine & " From ������Ϣ B, ������� A,ҽ�ƿ���� YLK,����ҽ�ƿ���Ϣ YLKXX"
    strSql = strSql & vbNewLine & " Where B.���� = A.���(+) and b.����id=ylkxx.����id(+) and ylkxx.״̬(+)=0 and  ylkxx.�����id=ylk.id(+)  and ylk.�Ƿ�����(+)=0 And B.ͣ��ʱ�� Is Null   "
    strSql = strSql & vbNewLine & strTmp
    
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
        
    If mrsInfo Is Nothing Then GoTo ClearPati:
    If mrsInfo.State <> 1 Then GoTo ClearPati:
    If mrsInfo.RecordCount = 0 Then GoTo ClearPati:
    If Val(NVL(mrsInfo!ID)) = 0 Then GoTo ClearPati:
    
    txtIdentify.Text = NVL(mrsInfo!����)
    Me.txtIdentify.Tag = NVL(mrsInfo!ID)
    txtIdentify.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtIdentify.IMEMode = 0
    GetPatient = True
    Exit Function
ClearPati:
    txtIdentify.Text = ""
    txtIdentify.PasswordChar = ""
    Set mrsInfo = Nothing
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtIdentify.IMEMode = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub PrintDetail()
'���ܣ�������б�
    Dim strNO As String
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    strNO = mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Դ�ӡ֤����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    intRow = mfrmNormal.vsfDetail.Row
    
    '��ͷ
    objOut.Title.Text = "���˽��ʵ�����ϸ"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmBalanceFilter
        objRow.Add "���ݺţ�" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("���ݺ�"))
        objRow.Add "���ʷ�Χ��" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("��ʼ����")) & " �� " & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("��������"))
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "סԺ�ţ�" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("סԺ��"))
        objRow.Add "������" & mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("����"))
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
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
'���ܣ���ǰ�տ��¼���´�ӡһ��Ʊ��
'bytMode=0-�ش�,1-����
    Dim strNO As String, lng����ID As Long, blnMediCare As Boolean, bytFlag As Byte '���ﻹ��סԺ
    Dim intInsure As Integer
    Dim lng����ID As Long, bytFunc As Byte
    
    With mfrmNormal.vsfMain
        strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        If strNO = "" Then
            MsgBox "��ǰû�е��ݿ����ش�Ʊ�ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        bytFunc = IIf(Val(.TextMatrix(.Row, .ColIndex("��־"))) = 1, 0, 1)
        
         '����Ȩ��
        If bytMode = 0 Then
            If Not BillOperCheck(7, .TextMatrix(.Row, .ColIndex("����Ա")), _
                CDate(.TextMatrix(.Row, .ColIndex("�շ�ʱ��"))), "�ش�") Then Exit Sub
        Else
            If Trim(.TextMatrix(.Row, .ColIndex("Ʊ�ݺ�"))) <> "" Then
                MsgBox "��ǰ�����Ѵ�ӡ��Ʊ��,���ܽ��в���", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        
        intInsure = BalanceExistInsure(strNO, bytFlag)
        If RePrintBalance(strNO, Me, lng����ID, intInsure) Then
            '��ҽһ��ͨд����85950
            Call WriteInforToCard(Me, mlngModule, mstrPrivs, gobjSquare.objSquareCard, 0, bytFunc, lng����ID, lng����ID)
            Call RefreshData
        End If
    End With
End Sub

Private Sub txtIdentify_LostFocus()
    IDKind.SetAutoReadCard False
End Sub
