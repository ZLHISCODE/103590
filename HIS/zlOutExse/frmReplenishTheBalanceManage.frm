VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceManage 
   Caption         =   "���ղ���������"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmReplenishTheBalanceManage.frx":0000
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
      Begin VB.Label lblȱʡ 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡ��ʾ"
         Height          =   180
         Left            =   60
         TabIndex        =   10
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
      TabIndex        =   2
      Top             =   870
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
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
Private mblnCancel As Boolean   '�ⲿж�ش����ʶ
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����
Private mrsInfo As ADODB.Recordset, mstrPrivsRollingCurtain As String
Private mobjInvoice As clsInvoice, mstrInvoice As String, mlng����ID As Long
Private mobjFactProperty As clsFactProperty
Private mblnFirst As Boolean

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    Select Case Control.ID
        Case conMenu_File_FeeCollect
            If zlCheckPrivs(mstrPrivsRollingCurtain, "����") = False Then Exit Sub
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
'            If zlCheckPrivs(mstrPrivs, "��������") = False Then Exit Sub
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
            zlDatabase.SetPara "ˢ�·�ʽ", "0", glngSys, mlngModule, zlCheckPrivs(mstrPrivs, "��������")
        Case conMenu_View_RefreshType_Ask
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Auto).Checked = False
            zlDatabase.SetPara "ˢ�·�ʽ", "1", glngSys, mlngModule, zlCheckPrivs(mstrPrivs, "��������")
        Case conMenu_View_RefreshType_Auto
            Control.Checked = True
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked = False
            mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_No).Checked = False
            zlDatabase.SetPara "ˢ�·�ʽ", "2", glngSys, mlngModule, zlCheckPrivs(mstrPrivs, "��������")
        Case conMenu_Edit_RegistBalance
            If zlCheckPrivs(mstrPrivs, "ҽ������") = False Then Exit Sub
            If frmReplenishTheBalanceBill.zlEditCard(Me, mlngModule, mstrPrivs, EM_Balance_Register) Then
                Call RefreshData
            End If
        Case conMenu_Edit_InsureBalance
            If zlCheckPrivs(mstrPrivs, "ҽ������") = False Then Exit Sub
            If frmReplenishTheBalanceBill.zlEditCard(Me, mlngModule, mstrPrivs, EM_Balance_Charge) Then
                Call RefreshData
            End If
        Case conMenu_Edit_BalanceDel
            If zlCheckPrivs(mstrPrivs, "�����˷�") = False Then Exit Sub
'            If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("�������")) = "" Then Exit Sub
            If frmReplenishTheBalanceDel.zlShowMe _
            (Me, mlngModule, mstrPrivs, EM_RBDTY_�˷�, mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("�������"))) Then
                Call RefreshData
            End If
        Case conMenu_Edit_ReDel
            If zlCheckPrivs(mstrPrivs, "�����˷�") = False Then Exit Sub
            If mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("�������")) = "" Then Exit Sub
            If frmReplenishTheBalanceDel.zlShowMe _
            (Me, mlngModule, mstrPrivs, EM_RBDTY_�쳣����, mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("�������")), , , , mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("����ʱ��"))) Then
                Call RefreshData
            End If
        Case conMenu_Edit_ReBalance
            If zlCheckPrivs(mstrPrivs, "ҽ������") = False Then Exit Sub
            If mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("����ID")) = "" Then Exit Sub
            If frmReplenishTheBalanceBill.zlEditCard _
            (Me, mlngModule, mstrPrivs, EM_Balance_Err_ReCharge, mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("����ID"))) Then
                Call RefreshData
            End If
        Case conMenu_Edit_BalanceCancel
            If zlCheckPrivs(mstrPrivs, "ҽ������") = False Then Exit Sub
            With mfrmErr.vsfMain
                If .TextMatrix(.Row, .ColIndex("����ID")) = "" Then Exit Sub
                If BalanceErrCancelCheck(Val(.TextMatrix(.Row, .ColIndex("����ID")))) = False Then Exit Sub
                If frmReplenishTheBalanceBill.zlEditCard _
                    (Me, mlngModule, mstrPrivs, EM_Balance_Err_Cancel, .TextMatrix(.Row, .ColIndex("����ID"))) Then Call RefreshData
            End With
        Case conMenu_Edit_ViewBalance
            Call ViewBalance(tabMain.Selected.Index)
        Case conMenu_Edit_PrintAmend
            Call PrintBill(2)
        Case conMenu_Edit_ReprintBalanceReceipt
            Call PrintBill(1)
        Case conMenu_Edit_DelPrint
            Call PrintDelBill
        Case conMenu_Edit_PrintList '��ӡ�����嵥
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

Private Function BalanceErrCancelCheck(ByVal lng����ID As Long) As Boolean
    '�쳣�������ϼ��
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandler
    '���ڷ�ҽ�����㷽ʽʱ�������쳣���ϣ�114149
    strSQL = _
        "Select 1" & vbNewLine & _
        "From ����Ԥ����¼ A, ���ò����¼ C, ���㷽ʽ B" & vbNewLine & _
        "Where a.������� = c.������� And a.���㷽ʽ = b.���� And b.���� Not In ('3', '4') And c.����id = [1] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not rsTmp.EOF Then
        MsgBox "���β�������ѳɹ�����Ľ��㷽ʽ�к��з�ҽ���ģ���˲����������ϣ�ֻ�ܽ������½��㣡", _
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

    strSQL = " Select A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, Sum(B.���ʽ��), A.����Ա����, A.�Ǽ�ʱ��, A.�������" & _
             " From ���ò����¼ A, ������ü�¼ B " & _
             " Where A.�Ǽ�ʱ�� Between [1] And [2] And Nvl(A.����״̬,0)=1 And A.�շѽ���ID=B.����ID And A.��¼״̬ = 2 And A.����Ա���� = [3]" & _
             " Group By A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, A.����Ա����, A.�Ǽ�ʱ��, A.�������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
    
    If rsTmp.RecordCount <> 0 Then
        tabMain.Item(2).Caption = "�쳣�˷Ѽ�¼(" & rsTmp.RecordCount & ")"
        If MsgBox("���ڲ�������쳣�˷Ѽ�¼,�Ƿ����쳣��¼?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            tabMain.Item(2).Selected = True
            Call mfrmRefund.ReadData(0)

            Exit Sub
        End If
    Else
        tabMain.Item(2).Caption = "�쳣�˷Ѽ�¼"
    End If
    
    strSQL = " Select A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�') As ����, B.����, B.�Ա�, B.����, Sum(B.���ʽ��), A.����Ա����, A.�Ǽ�ʱ��, A.�������" & _
             " From ���ò����¼ A, ������ü�¼ B " & _
             " Where A.�Ǽ�ʱ�� Between [1] And [2] And Nvl(A.����״̬,0)=1 And A.�շѽ���ID=B.����ID And A.��¼״̬ In (1,3) And A.����Ա���� = [3]" & _
             "       And Not Exists (Select 1 From ���ò����¼ Where �������=A.������� And ��¼״̬=2)" & _
             " Group By A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, A.����Ա����, A.�Ǽ�ʱ��, A.�������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
    
    If rsTmp.RecordCount <> 0 Then
        tabMain.Item(1).Caption = "�쳣�����¼(" & rsTmp.RecordCount & ")"
        If MsgBox("���ڲ�������쳣�����¼,�Ƿ����쳣��¼?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            tabMain.Item(1).Selected = True
            Call mfrmErr.ReadData(0)

            Exit Sub
        End If
    Else
        tabMain.Item(1).Caption = "�쳣�����¼"
    End If

End Sub

Public Sub ViewBalance(intType As Integer)
    Select Case intType
        Case 0
            If mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("�������")) = "" Then Exit Sub
            If Val(mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("�˷ѱ�־"))) = 2 Then
                frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_�鿴, mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("�������")), , , , mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("����ʱ��"))
            Else
                frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_�鿴, mfrmNormal.vsfMain.TextMatrix(mfrmNormal.vsfMain.Row, mfrmNormal.vsfMain.ColIndex("�������"))
            End If
        Case 1
            If mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("�������")) = "" Then Exit Sub
            frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_�鿴, mfrmErr.vsfMain.TextMatrix(mfrmErr.vsfMain.Row, mfrmErr.vsfMain.ColIndex("�������"))
        Case 2
            If mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("�������")) = "" Then Exit Sub
            frmReplenishTheBalanceDel.zlShowMe Me, mlngModule, mstrPrivs, EM_RBDTY_�鿴, mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("�������")), , , , mfrmRefund.vsfMain.TextMatrix(mfrmRefund.vsfMain.Row, mfrmRefund.vsfMain.ColIndex("����ʱ��"))
    End Select
End Sub

Private Sub RefreshData()
    If mcbrRefresh.CommandBar.Controls.Find(, conMenu_View_RefreshType_Ask).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode-������
    '     intType-�����������:0-Ĭ��,1-ֱ��Ԥ��,2-ֱ�Ӵ�ӡ,3-�����EXCEL
    '����:������
    '����:2013-09-22
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
                Control.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
                '0-�޼�¼,1-�շѼ�¼,2-�˷Ѽ�¼,3-�ѱ��˷ѵ��շѼ�¼
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
                Control.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
                With mfrmErr.vsfMain
                    If .TextMatrix(.Row, .ColIndex("�������")) <> "" And .Row <> 0 Then
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
                Control.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
                With mfrmErr.vsfMain
                    If .TextMatrix(.Row, .ColIndex("�������")) <> "" And .Row <> 0 Then
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
                Control.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
                With mfrmRefund.vsfMain
                    If .TextMatrix(.Row, .ColIndex("�������")) <> "" And .Row <> 0 Then
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
                Control.Visible = zlCheckPrivs(mstrPrivs, "�ش�Ʊ��")
                With mfrmNormal.vsfMain
                    If .TextMatrix(.Row, .ColIndex("�������")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("�˷ѱ�־"))) <> 2 And _
                        .TextMatrix(.Row, .ColIndex("ʵ��Ʊ��")) <> "" Then
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
                Control.Visible = zlCheckPrivs(mstrPrivs, "����Ʊ��")
                With mfrmNormal.vsfMain
                    If .TextMatrix(.Row, .ColIndex("�������")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("�˷ѱ�־"))) <> 2 And _
                        .TextMatrix(.Row, .ColIndex("ʵ��Ʊ��")) = "" Then
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
                    If .TextMatrix(.Row, .ColIndex("�������")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("�˷ѱ�־"))) = 2 Then
                        If Val(.TextMatrix(.Row, .ColIndex("��Ʊ�Ѵ�ӡ"))) = 1 Then
                            Control.Enabled = zlCheckPrivs(mstrPrivs, "�ش�Ʊ��")
                            Control.Caption = "�ش��˷�Ʊ��" & IIf(InStr(Control.Caption, "(") > 0, "(&B)", "")
                        Else
                            Control.Enabled = zlCheckPrivs(mstrPrivs, "����Ʊ��")
                            Control.Caption = "�����˷�Ʊ��" & IIf(InStr(Control.Caption, "(") > 0, "(&B)", "")
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
                Control.Visible = zlCheckPrivs(mstrPrivs, "��������嵥")
                With mfrmNormal.vsfMain
                    If .TextMatrix(.Row, .ColIndex("���㵥��")) <> "" And _
                        .Row <> 0 And Val(.TextMatrix(.Row, .ColIndex("�˷ѱ�־"))) <> 2 Then
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
                        If .TextMatrix(.Row, .ColIndex("�������")) <> "" And .Row <> 0 Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End With
                Case 1
                    With mfrmErr.vsfMain
                        If .TextMatrix(.Row, .ColIndex("�������")) <> "" And .Row <> 0 Then
                            Control.Enabled = True
                        Else
                            Control.Enabled = False
                        End If
                    End With
                Case 2
                    With mfrmRefund.vsfMain
                        If .TextMatrix(.Row, .ColIndex("�������")) <> "" And .Row <> 0 Then
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
    mstrTitle = "���ղ���������"
    '��ӡ������ʼ��
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    Call mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Set mobjFactProperty = New clsFactProperty
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_�շ��վ�, 0, 0, 0, mobjFactProperty)
    Call zlDefCommandBars
    '����TAB��Ϣ
    Call SetTabControl
    Call InitIDKind
    Call SetCboDate
    stbThis.Panels(3).Text = UserInfo.����
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
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
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:������
    '����:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim intPara As Integer
    
    Err = 0: On Error GoTo Errhand:
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "�շ�Ա����(&M)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlCheckPrivs(mstrPrivsRollingCurtain, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_SetInsure, "�������(&I)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)"): mcbrControl.BeginGroup = True
'        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "��������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RegistBalance, "�ҺŽ���(&J)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3019
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InsureBalance, "�շѽ���(&S)")
        mcbrControl.IconId = 3011
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceDel, "�����˷�(&U)")
        mcbrControl.IconId = 3017
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBalance, "���½���(&J)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3831
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceCancel, "��������(&C)")
        mcbrControl.IconId = 3832
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReDel, "�����˷�(&D)")
        mcbrControl.IconId = 228
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ViewBalance, "�鿴����(&V)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintBalanceReceipt, "�ش����Ʊ��(&R)"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�ش�Ʊ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "�������Ʊ��(&R)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "����Ʊ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DelPrint, "�����˷�Ʊ��(&B)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "����Ʊ��") Or zlCheckPrivs(mstrPrivs, "�ش�Ʊ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintList, "��ӡ�շ��嵥(&L)")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "��������嵥")
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
        intPara = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModule, "0"))
        Set mcbrRefresh = .Add(xtpControlPopup, conMenu_View_RefreshType, "ˢ�·�ʽ(&O)"): mcbrControl.BeginGroup = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_No, "������ˢ������(&1)", -1, False)
        If intPara = 0 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Ask, "��������ʾˢ������(&2)", -1, False)
        If intPara = 1 Then cbrControl.Checked = True
        Set cbrControl = mcbrRefresh.CommandBar.Controls.Add(xtpControlButton, conMenu_View_RefreshType_Auto, "�������Զ�ˢ������(&3)", -1, False)
        If intPara = 2 Then cbrControl.Checked = True
        mcbrRefresh.Visible = zlCheckPrivs(mstrPrivs, "��������")
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "�շ�Ա����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlCheckPrivs(mstrPrivsRollingCurtain, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
'        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "��������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RegistBalance, "�ҺŽ���"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3019
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InsureBalance, "�շѽ���")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceDel, "�����˷�")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
        mcbrControl.IconId = 3017
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBalance, "���½���")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3831
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceCancel, "��������")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3832
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReDel, "�����˷�")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ViewBalance, "���ĵ���"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 221
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReprintBalanceReceipt, "�ش����Ʊ��"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�ش�Ʊ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintAmend, "�������Ʊ��")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "����Ʊ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DelPrint, "�����˷�Ʊ��")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "����Ʊ��") Or zlCheckPrivs(mstrPrivs, "�ش�Ʊ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PrintList, "��ӡ�շ��嵥")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "��������嵥")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
    End With
    
    '�����
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RegistBalance, "�ҺŽ���"): mcbrControl.BeginGroup = True
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3019
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_InsureBalance, "�շѽ���")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceDel, "�����˷�")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
        mcbrControl.IconId = 3017
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBalance, "���½���")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3831
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BalanceCancel, "��������")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "ҽ������")
        mcbrControl.IconId = 3832
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReDel, "�����˷�")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "�����˷�")
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "�շ�Ա����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3588
        mcbrControl.Visible = zlCheckPrivs(mstrPrivsRollingCurtain, "����")
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
    
    '����״̬����������
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
    
    '�洢�б�ĸ��Ի�����(����)
    
    
    SaveWinState Me, App.ProductName, mstrTitle
    'ж�ؼ��ش������
    Set mrsInfo = Nothing
End Sub

Private Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ����Ȩ���Ƿ����
    '����:strPrivs-Ȩ�޴�
    '     strMyPriv-����Ȩ��
    '����,����Ȩ��,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-19 14:34:46
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
        Case 0 '����
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
        Case 1 '�������
            dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 2 '�������
            dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 3  '���һ��
            dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 4  '����
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
    i = Val(zlDatabase.GetPara("�쳣���ݲ�ѯ", glngSys, mlngModule, 0, Array(lblȱʡ, cboDate)))
    With cboDate
        .Clear
        .AddItem "����"
        .ListIndex = .NewIndex
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "�������"
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "�������"
        If i = 2 Then .ListIndex = .NewIndex
        .AddItem "���һ��"
        If i = 3 Then .ListIndex = .NewIndex
        .AddItem "����"
        If i = 4 Then .ListIndex = .NewIndex
        .AddItem "�Զ���"
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
    txtIdentify.Text = objPatiInfor.����
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
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtIdentify.Locked Then Exit Sub
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "����") Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtIdentify.Text, 1)) > 0 And IsNumeric(Mid(txtIdentify.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtIdentify, KeyAscii, IDKind.ShowPassText)
        End If
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
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtIdentify)
    IDKind.AllowAutoCommCard = True
    IDKind.AllowAutoICCard = True
    IDKind.AllowAutoIDCard = True
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
        Call mfrmNormal.ReadData(1, mstrPrivs, Val(Nvl(mrsInfo!ID)))
    Case 1
        Call mfrmErr.ReadData(1, Val(Nvl(mrsInfo!ID)))
    Case 2
        Call mfrmRefund.ReadData(1, Val(Nvl(mrsInfo!ID)))
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
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    
    strSQL = ""
    If blnCard And objCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then    '103563
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        If lng����ID <= 0 Then lng����ID = 0
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.����ID=[2] " & str����Ժ
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And B.�����=[2]" & str����Ժ
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strSQL = strSQL & " And B.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])" & str����Ժ
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                '����
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtIdentify.Text = mrsInfo!���� Then blnSame = True
                End If
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                        txtIdentify.Text = ""
                        Set mrsInfo = Nothing: Exit Function
                    Else
                       'strSQL = strSQL & " And  B.���� Like [3]"
                       '�����:50485
                        strPati = _
                            " Select /*+Rule */distinct 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ,decode(b.����,Null,Null,'��') As �Ƿ���ҽ�ƿ�" & _
                            " From ������Ϣ A, ����ҽ�ƿ���Ϣ B " & _
                            " Where Rownum <101 And a.����ID=b.����ID(+) And b.״̬(+)=0 And B.�����ID(+)=[3]  And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & str����Ժ & _
                            IIf(gintNameDays = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                            
                        vRect = zlControl.GetControlRect(txtIdentify.hWnd)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtIdentify.Height, blnCancel, False, True, strInput & "%", gintNameDays, Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, 0)))
                        If Not rsTmp Is Nothing Then
                            If rsTmp!ID = 0 Then
                                Set mrsInfo = Nothing: Exit Function
                            Else
                                strInput = "-" & rsTmp!����ID
                                strSQL = strSQL & " And B.����ID=[2]"
                            End If
                        Else 'ȡ��ѡ��
                            txtIdentify.Text = ""
                            Set mrsInfo = Nothing: Exit Function
                        End If
                    End If
                Else
                    strSQL = strSQL & " And B.����ID=[2]"
                    strInput = "-" & Val(mrsInfo!����ID)
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And B.ҽ����=[1]" & str����Ժ
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                ' strSQL = strSQL & " And B.���֤��=[1] " & str����Ժ
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.�����=[1]" & str����Ժ
                '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])" & str����Ժ
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
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSQL
    strSQL = "    " & vbNewLine & " Select /*+Rule */distinct  B.����id As ID, Decode(sign(nvl(ylkxx.����id,0)),0,'','��') as �����˻�, B.����id,B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ,"
    strSQL = strSQL & vbNewLine & "      A.���� ��������"
    strSQL = strSQL & vbNewLine & " From ������Ϣ B, ������� A,ҽ�ƿ���� YLK,����ҽ�ƿ���Ϣ YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.���� = A.���(+) and b.����id=ylkxx.����id(+) and ylkxx.״̬(+)=0 and  ylkxx.�����id=ylk.id(+)  and ylk.�Ƿ�����(+)=0 And B.ͣ��ʱ�� Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
    
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
        
    If mrsInfo Is Nothing Then GoTo ClearPati:
    If mrsInfo.State <> 1 Then GoTo ClearPati:
    If mrsInfo.RecordCount = 0 Then GoTo ClearPati:
    If Val(Nvl(mrsInfo!ID)) = 0 Then GoTo ClearPati:
    
    txtIdentify.Text = Nvl(mrsInfo!����)
    Me.txtIdentify.Tag = Nvl(mrsInfo!ID)
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

Private Function GetBalanceInsure(ByVal str������� As String, _
    Optional ByRef str�������� As String, Optional ByRef lng����ID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������
    '����:str��������-��������
    '     lng����ID-����ID
    '����:��������
    '����:���˺�
    '����:2014-09-22 13:57:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
    strSQL = "" & _
        "   Select /*+ rule */ b.��¼id, b.����, b.����id, c.����" & _
        "   From ���ò����¼ A, ���ս����¼ B, ������� C" & _
        "   Where a.����id = b.��¼id And b.���� = c.���(+) And b.���� = 1 And a.������� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�������)
    If Not rsTmp.EOF Then
        lng����ID = Nvl(rsTmp!����ID, 0)
        str�������� = Nvl(rsTmp!����)
        GetBalanceInsure = Nvl(rsTmp!����, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-30 14:15:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.����, EM_�շ��վ�, mobjFactProperty.ʹ�����, lng����ID, mobjFactProperty.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    If lng����ID <= 0 Then
        Select Case lng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(mobjFactProperty.ʹ�����) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & mobjFactProperty.ʹ����� & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mobjFactProperty.ʹ�����) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & mobjFactProperty.ʹ����� & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
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
    '����:ˢ���շ�Ʊ�ݺ�
    '����:���˺�
    '����:2014-06-06 14:21:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjFactProperty Is Nothing Then Exit Sub
    If mobjFactProperty.��ӡ��ʽ = 0 Then Exit Sub
    
    If mobjFactProperty.�ϸ���� Then
            
        If zlGetInvoiceGroupUseID(mlng����ID) = False Then
            mstrInvoice = "": Exit Sub
        End If
        '�ϸ�ȡ��һ������
        If mobjInvoice.zlGetNextBill(mlngModule, mlng����ID, strFactNO) = False Then strFactNO = ""
        mstrInvoice = strFactNO
        
    Else
        '��ɢ��ȡ��һ������
        mstrInvoice = zlStr.Increase(UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, mlngModule)))
    End If
End Sub

Private Function GetFeeNos(ByVal strNo As String) As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, strResult As String
    strSQL = _
        " Select Distinct NO" & vbNewLine & _
        " From ������ü�¼" & vbNewLine & _
        " Where ����id In (Select Distinct �շѽ���id From ���ò����¼ Where NO = [1])"
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
    '����:Ʊ�ݴ�ӡ
    '���:bytType-�������� 1:�ش�Ʊ�� 2:����Ʊ��
    '����:������
    '����:2014-09-24 17:33:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVirtualPrint As Boolean, lng����ID As Long
    Dim intPrint As Integer, dtDate As Date, strNos As String, str��� As String
    Dim intInsure As Integer, i As Integer, strNo As String
    If tabMain.Selected.Index <> 0 Then Exit Sub
    If bytType = 2 Then
        With mfrmNormal.vsfInvoice
            If .TextMatrix(1, 1) <> "" Then
                MsgBox "ѡ��ļ�¼�Ѿ���ӡ��Ʊ��,���ܽ��в���", vbInformation, gstrSysName
                Exit Sub
            End If
        End With
    End If
    With mfrmNormal.vsfMain
        strNo = .TextMatrix(.Row, .ColIndex("���㵥��"))
        strNos = GetFeeNos(strNo)
        str��� = .TextMatrix(.Row, .ColIndex("����"))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        intInsure = GetBalanceInsure(Val(.TextMatrix(.Row, .ColIndex("�������"))))
        Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_�շ��վ�, lng����ID, 0, intInsure, mobjFactProperty)
        dtDate = .TextMatrix(.Row, .ColIndex("����ʱ��"))
        If strNo = "" Then Exit Sub
        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(Val(.TextMatrix(.Row, .ColIndex("�������")))))
        If Not blnVirtualPrint And strNos <> "" Then
            If str��� = "�շ�" Then
                If Not BillExistMoney(strNos, 1) Then
                    MsgBox "ѡ��ļ�¼�Ѿ�ȫ���˷�,���ܽ��д�ӡ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                If Not BillExistMoney(strNos, 4) Then
                    MsgBox "ѡ��ļ�¼�Ѿ�ȫ���˷�,���ܽ��д�ӡ��", vbInformation, gstrSysName
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
    '����:Ʊ�ݴ�ӡ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVirtualPrint As Boolean, lng����ID As Long
    Dim intPrint As Integer, strNos As String, str��� As String
    Dim intInsure As Integer, i As Integer, lng������� As Long
    
    Err = 0: On Error GoTo Errhand
    If tabMain.Selected.Index <> 0 Then Exit Sub
    With mfrmNormal.vsfMain
        lng������� = Val(.TextMatrix(.Row, .ColIndex("�������")))
        If lng������� = 0 Then Exit Sub
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        intInsure = GetBalanceInsure(lng�������)
        
        Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_�˷��վ�, lng����ID, 0, intInsure, mobjFactProperty)
        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(Val(.TextMatrix(.Row, .ColIndex("�������")))))
        
        If zlPrintReplenishTheDelBalanceBill(Me, mlngModule, lng�������, intInsure, mobjInvoice, mobjFactProperty, , zlDatabase.Currentdate, blnVirtualPrint) Then
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
    '��ӡ�շѽ����嵥
    Dim strNo As String
    
    On Error GoTo Errhand
    With mfrmNormal
        strNo = .vsfMain.TextMatrix(.vsfMain.Row, .vsfMain.ColIndex("���㵥��"))
        If strNo = "" Then
            MsgBox "��ǰû�е��ݿ��Դ�ӡ�嵥��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�Ƿ���ת������ݱ���
        If .mblnNOMoved Then
            If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
            .mblnNOMoved = False  '��ʱ��ת���������ݱ�
        End If
    End With
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO='" & strNo & "'", "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
    End If
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
