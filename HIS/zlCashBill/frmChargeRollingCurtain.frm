VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeRollingCurtain 
   BorderStyle     =   0  'None
   Caption         =   "�շ�Ա����"
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
   StartUpPosition =   3  '����ȱʡ
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "����(&Z)"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "������ȡ����(&R)"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "�ݴ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�Ͻɽ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����˵��(&M)"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ϴ�����ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ֹ����ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�տ��(&D)"
         BeginProperty Font 
            Name            =   "����"
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
    EM_PN_Filter = 260101   '��������
    EM_PN_ChargeBillTotal = 260102  '�տƱ�ݻ���
End Enum
Private mobjChargeBill As clsChargeBill, mfrmMain As Object
Private mlngModule As Long, mstrPrivs As String, mdatEnd As Date, mdatBegin As Date
Private mstrPreDate As String, mstrOperatorName As String, mblnNotClick As Boolean
Private mdblDefaultHandIn As Double
Private mstrRollingType As String    '���
Private mblnChangeEndDate As Boolean
Private mblnNotChange As Boolean, mstrDefaultTime As String

Public Sub RefreshPage()
    Call cmdRefresh_Click
End Sub

Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mobjChargeBill.zlPrint(bytMode, "", txtMemo.Text)
End Sub

Public Sub zlInitVar(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal strPreDate As String, ByVal strOperatorName As String, ByVal strRollingType As String, Optional ByVal strDefaultTime As String = "")
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '���:lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '     strPreDate-�ϴ�����ʱ��
    '     strOperatorName-����Ա����
    '     strRollingType-�������
    '����:
    '����:���˺�
    '����:2013-09-09 14:41:46
    '˵��:���ش����,��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mlngModule = lngModule: mstrPrivs = strPrivs: mstrOperatorName = strOperatorName
    mstrPreDate = strPreDate: mstrRollingType = strRollingType
    mstrDefaultTime = strDefaultTime
    Call InitFace: Call SetPopedom
End Sub
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2013-09-11 14:05:08
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date
    
    '��ȡ�ϴ�����ʱ��
    dtDate = zlDatabase.Currentdate
    
    dtpStartDate.Enabled = mstrPreDate = "" '�ϴ�����ʱ��Ϊ��ʱ,��Ҫ�ֹ�ѡ��ȷ��ʱ��
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
    '����:����Ȩ�޿���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 12:02:06
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdOK.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngFilterHeight As Long, lngBillHeight As Long
    lngFilterHeight = 1215 / Screen.TwipsPerPixelX
    lngBillHeight = 1000 / Screen.TwipsPerPixelX
    With dkpMan
        'Set .ImageList = zlCommFun.GetPubIcons
        Set objPane = .CreatePane(EM_PN_Filter, 100, lngFilterHeight, DockLeftOf, Nothing)
        objPane.Title = "��������": objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngFilterHeight: objPane.MaxTrackSize.Height = lngFilterHeight
        objPane.Handle = picFilter.hWnd
        Set objPane = .CreatePane(EM_PN_ChargeBillTotal, 400, 400, DockBottomOf, objPane)
        objPane.Title = "�տƱ�ݻ���"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
        objPane.Handle = mobjChargeBill.GetChargeAndBillTotalForm.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function
Private Function CheckValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ĺϷ���
    '����:��������Ϸ�������true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 11:45:04
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
'    If cboDept.ListIndex < 0 Then
'        MsgBox "ע��:" & vbCrLf & "   δѡ���տ��,��ѡ���տ��!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
'        If cboDept.Enabled And cboDept.Visible Then cboDept.SetFocus
'        Exit Function
'    End If
    If InStr(Trim(txtMemo.Text), "'") > 0 Then
        MsgBox "ע��:" & vbCrLf & "   ����˵���������е�����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    
    '�����:110281,����,2017/08/11,������˵�������޴�50���ַ�����Ϊ500���ַ�
    If zlCommFun.ActualLen(txtMemo.Text) > 500 Then
        MsgBox "ע��:" & vbCrLf & "   ����˵�����ֻ������500���ַ���250������,����������!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
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
    '����:����Ʊ�ݴ�ӡ
    '����:���˺�
    '����:2013-09-11 11:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    blnPrint = False
    If Not zlStr.IsHavePrivs(mstrPrivs, "�ɿ����ӡ") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("�ɿ����ӡ��ʽ", glngSys, mlngModule))     'ʹ��ҽ��վ����ز���
    Case 0    '����ӡ
        Exit Sub
    Case 1    '��������ӡ
        blnPrint = True
    Case 2    'ѡ���ӡ
        If MsgBox("���Ƿ�Ҫ��ӡ�ɿ��飿", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            blnPrint = True
        End If
    End Select
    If blnPrint = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1506", Me, "NO=" & strNO, 2)
End Sub

Public Function SaveData(ByRef lngID As Long, ByRef strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '����:lngId-����ID
    '       strNo-���ʵ���
    '����:�������ݱ���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 11:39:42
    '˵��:
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
        'Ʊ�ݴ�ӡ
        dtpStartDate.Value = dtpEndDate.Value: dtpStartDate.Enabled = False
        dtpEndDate.MaxDate = DateAdd("d", 1, zlDatabase.Currentdate)
        dtpEndDate.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        Call BillPrint(strNO)
        '���¼�������
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
        MsgBox "���ʽ���ʱ�䳬���˵�ǰ��ϵͳʱ��(" & Format(datTemp, "yyyy-mm-dd hh:mm:ss") & "),���������ʣ�", vbCritical, gstrSysName
        Exit Sub
    End If
    If mdatEnd <> dtpEndDate Then
        If MsgBox("��ǰ������ʱ������ȡ���ݵ�����ʱ�䲻һ�£��Ƿ����µ�����ʱ������ˢ�����ݣ�", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
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
        If MsgBox("��ֹʱ�䷢���˱仯,��������ȡ��Ҫ���ʵ�����,�Ƿ�������������?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
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
        '�޶���������
        If (KeyAscii < Asc(".") Or KeyAscii = Asc("/") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
            KeyAscii = 0
            Exit Sub
        End If
        'С������ж�
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
    '��ȡ�ֽ���
    GetCashMoney = mobjChargeBill.GetChargeAndBillTotalForm.GetCashMoney
End Property
Public Sub zlRefresh()
    '���½�������ˢ��
    Call cmdRefresh_Click
End Sub
Public Sub ShowChargeList(ByVal frmMain As Object, Optional strRollingType As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ�տ�����
    '���:frmMain-���õ�������
    '    strRollingType-�������,bytType=1ʱ��Ч�ֱ�Ϊ:
    '               0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�
    '����:���˺�
    '����:2013-09-16 17:33:21
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
    '����:�����Զ��屨��
    '���:lngSys-ϵͳ��
    '        strRptCode-������
    '����:���˺�
    '����:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'Dim lngDeptID As Long
    'If cboDept.ListIndex >= 0 Then lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "��ʼ��������=" & Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM:SS"), _
        "��ֹ��������=" & Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM:SS"))
End Sub

Public Sub zlDefaultSetFocus()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ����
    '����:���˺�
    '����:2013-10-16 14:23:12
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
    '����:������ת��Ĺ��ܼ�(��������)
    '����:���˺�
    '����:2013-10-16 15:14:14
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
