VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillView 
   BorderStyle     =   0  'None
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   -285
   ClientWidth     =   9075
   Icon            =   "frmBillView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   1  'Minimized
   Begin VB.PictureBox picDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   8415
      TabIndex        =   14
      Top             =   1080
      Width           =   8415
      Begin VB.PictureBox picFile 
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   840
         ScaleHeight     =   6495
         ScaleWidth      =   6735
         TabIndex        =   27
         Top             =   2040
         Width           =   6735
         Begin zl9CISCore.ctrlPatientFile ProFile1 
            Height          =   5175
            Index           =   0
            Left            =   480
            TabIndex        =   13
            Top             =   120
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   9128
            Border_Width    =   0
         End
      End
      Begin VB.PictureBox picAdvice 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   9255
         TabIndex        =   15
         Top             =   0
         Width           =   9255
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   6440
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk��ʼʱ�� 
            BackColor       =   &H80000005&
            Caption         =   "Ҫ��ʱ��"
            Height          =   225
            Left            =   315
            TabIndex        =   4
            ToolTipText     =   "�Ƿ���ʱ��"
            Top             =   420
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   10
            Top             =   1080
            Width           =   1380
         End
         Begin VB.TextBox txtƵ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1350
            TabIndex        =   8
            Top             =   1080
            Width           =   2500
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4725
            MaxLength       =   3
            TabIndex        =   9
            Top             =   1080
            Width           =   1380
         End
         Begin VB.CheckBox chk���� 
            BackColor       =   &H80000005&
            Caption         =   "����(&J)"
            Height          =   225
            Left            =   4200
            TabIndex        =   6
            Top             =   405
            Width           =   945
         End
         Begin VB.CommandButton cmdExt 
            Height          =   285
            Left            =   8040
            Picture         =   "frmBillView.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "ѡ�����걾"
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   285
            Left            =   5280
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   0
            Width           =   285
         End
         Begin VB.ComboBox cboִ�п��� 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "frmBillView.frx":0102
            Left            =   1350
            List            =   "frmBillView.frx":0104
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1440
            Width           =   2500
         End
         Begin VB.TextBox txtҽ������ 
            Height          =   300
            Left            =   1350
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   0
            Width           =   3945
         End
         Begin VB.ComboBox cboҽ�� 
            Height          =   300
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1425
            Width           =   1590
         End
         Begin VB.TextBox txtҽ������ 
            Height          =   300
            Left            =   1350
            MaxLength       =   100
            TabIndex        =   7
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmdƵ�� 
            Enabled         =   0   'False
            Height          =   240
            Left            =   3575
            Picture         =   "frmBillView.frx":0106
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(F4)"
            Top             =   1110
            Width           =   270
         End
         Begin MSComCtl2.DTPicker txt��ʼʱ�� 
            Height          =   300
            Left            =   1350
            TabIndex        =   5
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   62193667
            CurrentDate     =   38022
         End
         Begin VB.Line lineTitleSplit 
            BorderColor     =   &H80000000&
            X1              =   400
            X2              =   1440
            Y1              =   320
            Y2              =   320
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��鲿λ"
            Height          =   180
            Left            =   5640
            TabIndex        =   28
            Top             =   45
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ��"
            Height          =   180
            Left            =   6660
            TabIndex        =   26
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl������λ 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   8460
            TabIndex        =   25
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lblƵ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ƶ��"
            Height          =   180
            Left            =   960
            TabIndex        =   24
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl������λ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   6150
            TabIndex        =   23
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   4335
            TabIndex        =   22
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lblִ�п��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ�п���"
            Height          =   180
            Left            =   600
            TabIndex        =   21
            Top             =   1500
            Width           =   720
         End
         Begin VB.Label lblҽ������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ŀ"
            Height          =   180
            Left            =   600
            TabIndex        =   20
            Top             =   45
            Width           =   720
         End
         Begin VB.Label lbl��ʼʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ҫ��ʱ��"
            Height          =   180
            Left            =   600
            TabIndex        =   19
            Top             =   435
            Width           =   720
         End
         Begin VB.Label lbl����ҽ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽ��"
            Height          =   180
            Left            =   5175
            TabIndex        =   18
            Top             =   1485
            Width           =   720
         End
         Begin VB.Label lblҽ������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ������"
            Height          =   180
            Left            =   585
            TabIndex        =   17
            Top             =   795
            Width           =   720
         End
         Begin VB.Line lineSplit 
            X1              =   0
            X2              =   1080
            Y1              =   1800
            Y2              =   1800
         End
      End
   End
End
Attribute VB_Name = "frmBillView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PatientID As String '����ID
Private CheckID As String '����ID��Һŵ�ID
Private PatientType As Integer '0=���ﲡ�� 1=סԺ����
Private FileTypeID As String '����ģ���ļ�ID
Private bSample As Boolean '�Ƿ�ʾ��
Private blnMoved As Boolean
Private prbRefresh As Object
Attribute prbRefresh.VB_VarHelpID = -1

Private AdviceID As Long 'ҽ��ID
Private sCheckNo As String '���͵��ݺ�
Private iRecordType As Integer '��¼����
Private alngFileID(1) As Long '����ͱ���ID
Private intType As Integer '�������:-1=������0=�����ϡ�1=������2=��ҩ��3=����
Private iTabIndex As Integer

'ҽ���༭
Private rsRelativeAdvice As ADODB.Recordset '���ҽ��
Private strExtData As String '������Ŀ

Private iCurrElementIndex As Integer '��ǰԪ��˳���

Public Sub ShowMe(ByVal lngҽ��ID As Long, Optional objPrbRefresh As Object, Optional DataMoved As Boolean = False)
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '������Ŀ����
    Dim strDrAdvice As String 'ҽ������
    Dim bAllowEdit As Boolean
    Dim strSQL As String
    
    If AdviceID = lngҽ��ID Then Exit Sub
    
    AdviceID = lngҽ��ID
    blnMoved = DataMoved
    On Error Resume Next
    '��ʼ��
    Set prbRefresh = objPrbRefresh
    ClearForm
    
    strSQL = "Select a.����ID,a.��ҳID,a.�Һŵ�,Decode(a.��ҳID,Null,0,1),b.ID,b.����,a.ҽ������," + _
        "ҽ������,��ʼִ��ʱ��,������־,ִ��Ƶ��,�ܸ�����,��������,c.���� As ���ұ���,c.���� As ��������,����ҽ��,nvl(b.���㵥λ,' ') As ���㵥λ,b.���,nvl(a.�걾��λ,' ') As �걾��λ,Nvl(a.����ID,0) As ����ID,d.�����ļ�ID " + _
        "From ����ҽ����¼ a,������ĿĿ¼ b,���ű� c,���Ƶ���Ӧ�� d Where (a.ID=[1] Or a.���ID=[1]) And a.������ĿID=b.ID And a.ִ�п���ID=c.ID " + _
        "And b.ID=d.������ĿID And d.Ӧ�ó���=Decode(a.��ҳID,Null,1,2) Order By nvl(a.���ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, lngҽ��ID)
    If rsTmp.EOF Then Exit Sub
    prbRefresh.Value = 5
    
    strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
    '���츽����Ŀ��
    rsTmp.MoveNext: strExtData = ""
    Do While Not rsTmp.EOF
        strExtData = strExtData & "," & rsTmp(4)
    
        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    rsTmp.MoveFirst
    
    intType = -1
    Me.txtҽ������ = strDiagName
    If rsTmp!��� = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "�����Ŀ"), 0) = 1 Then
        '��������Ŀ
        intType = 0
        Call AdviceSet�������(1, strExtData)
        txtҽ������.Text = Get�����������(1, strDiagName)
        Me.txt���� = Get��λ����
    ElseIf rsTmp!��� = "F" Then
        '��������Ҫ����������Ŀ������ѡ�񸽼�����
        intType = 1
        Call AdviceSet�������(2, strExtData)
        txtҽ������.Text = Get�����������(2, strDiagName)
        Me.txt���� = Get��������
    ElseIf InStr(",7,8,", rsTmp!���) > 0 Then
        '��ҩ�䷽(��ζ��ҩ���䷽����)
        intType = 2
    ElseIf rsTmp!��� = "C" Then
        '������Ŀѡ�����걾
        intType = 3
        Me.txt���� = rsTmp("�걾��λ")
    End If
    
    alngFileID(0) = rsTmp("����ID"): PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    PatientType = rsTmp(3): FileTypeID = rsTmp("�����ļ�ID"): bSample = False
    
    '��ʾҽ������
    If IsNull(rsTmp("��ʼִ��ʱ��")) Then
        Me.chk��ʼʱ��.Visible = True: Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ�� = CDate(Date & " " & Time): Me.txt��ʼʱ��.Enabled = False
    Else
        Me.txt��ʼʱ�� = rsTmp("��ʼִ��ʱ��"): Me.txt��ʼʱ��.Enabled = True
    End If
    Me.chk����.Value = rsTmp("������־")
    If Not IsNull(rsTmp("ҽ������")) Then Me.txtҽ������ = rsTmp("ҽ������")
    Me.txtƵ�� = rsTmp("ִ��Ƶ��"): Me.txtƵ��.Enabled = True: Me.cmdƵ��.Enabled = True
    Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    If Not IsNull(rsTmp("�ܸ�����")) Then Me.txt���� = rsTmp("�ܸ�����"): Me.txt����.Enabled = True
    If Not IsNull(rsTmp("��������")) Then Me.txt���� = rsTmp("��������"): Me.txt����.Enabled = True: Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    Me.cboִ�п���.Clear: Me.cboִ�п���.AddItem rsTmp("���ұ���") & "-" & rsTmp("��������")
    Me.cboִ�п���.Text = rsTmp("���ұ���") & "-" & rsTmp("��������"): Me.cboִ�п���.Enabled = True
    Me.cboҽ��.Clear: Me.cboҽ��.AddItem rsTmp("����ҽ��")
    Me.cboҽ��.Text = rsTmp("����ҽ��"): Me.cboҽ��.Enabled = True
    Me.picAdvice.Enabled = False
    
    SetItemFormat
    prbRefresh.Value = 15
    '��ʼ������
    
    '�ж��ܷ�༭����
    bAllowEdit = False
    iCurrElementIndex = 1
    
    Me.MousePointer = vbHourglass
'    ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 1, prbRefresh, , , , blnMoved
'    ProFile1(0).SetActiveElement 1
    Me.MousePointer = vbDefault
End Sub

Public Sub ShowMe_Report(ByVal strNO As String, ByVal int��¼���� As Integer, Optional objPrbRefresh As Object, Optional DataMoved As Boolean = False)
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '������Ŀ����
    Dim strDrAdvice As String 'ҽ������
    Dim bAllowEdit As Boolean
    Dim strSQL As String

    If sCheckNo = strNO And iRecordType = int��¼���� Then Exit Sub
    
    sCheckNo = strNO: iRecordType = int��¼����
    blnMoved = DataMoved
    On Error Resume Next
    '��ʼ��
    Set prbRefresh = objPrbRefresh
    ClearForm

    strSQL = "Select a.����ID,a.��ҳID,a.�Һŵ�,Decode(a.��ҳID,Null,0,1),b.ID,b.����,a.ҽ������,a.ID,a.����ID," + _
        "ҽ������,��ʼִ��ʱ��,������־,ִ��Ƶ��,�ܸ�����,��������,d.���� As ���ұ���,d.���� As ��������,����ҽ��,b.���,nvl(a.�걾��λ,' ') As �걾��λ,nvl(c.����ID,0) As ����ID,e.�����ļ�ID  " + _
        "From ����ҽ����¼ a,������ĿĿ¼ b,����ҽ������ c,���ű� d,���Ƶ���Ӧ�� e Where" & _
        " c.NO=[1] and c.��¼����=[2] And a.������ĿID=b.ID(+) And a.ID=c.ҽ��ID And a.ִ�п���ID=d.ID(+) " + _
        "And b.ID=e.������ĿID And e.Ӧ�ó���=Decode(a.��ҳID,Null,1,2) Order By nvl(a.���ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, strNO, int��¼����)
    If rsTmp.EOF Then Exit Sub
    prbRefresh.Value = 5
    
    strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
     '���츽����Ŀ��
    rsTmp.MoveNext: strExtData = ""
    Do While Not rsTmp.EOF
        strExtData = strExtData & "," & rsTmp(4)

        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    rsTmp.MoveFirst

    intType = -1
    Me.txtҽ������ = strDiagName
    If rsTmp!��� = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "�����Ŀ"), 0) = 1 Then
        '��������Ŀ
        intType = 0
        Call AdviceSet�������(1, strExtData)
        txtҽ������.Text = Get�����������(1, strDiagName)
        Me.txt���� = Get��λ����
    ElseIf rsTmp!��� = "F" Then
        '��������Ҫ����������Ŀ������ѡ�񸽼�����
        intType = 1
        Call AdviceSet�������(2, strExtData)
        txtҽ������.Text = Get�����������(2, strDiagName)
        Me.txt���� = Get��������
    ElseIf InStr(",7,8,", rsTmp!���) > 0 Then
        '��ҩ�䷽(��ζ��ҩ���䷽����)
        intType = 2
    ElseIf rsTmp!��� = "C" Then
        '������Ŀѡ�����걾
        intType = 3
        Me.txt���� = rsTmp("�걾��λ")
    End If

    alngFileID(0) = rsTmp("����ID"): PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    PatientType = rsTmp(3): FileTypeID = 0: bSample = False: AdviceID = rsTmp(7)

    '��ʾҽ������
    If IsNull(rsTmp("��ʼִ��ʱ��")) Then
        Me.chk��ʼʱ��.Visible = True: Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ�� = CDate(Date & " " & Time): Me.txt��ʼʱ��.Enabled = False
    Else
        Me.txt��ʼʱ�� = rsTmp("��ʼִ��ʱ��"): Me.txt��ʼʱ��.Enabled = True
    End If
    Me.chk����.Value = rsTmp("������־")
    If Not IsNull(rsTmp("ҽ������")) Then Me.txtҽ������ = rsTmp("ҽ������")
    Me.txtƵ�� = rsTmp("ִ��Ƶ��"): Me.txtƵ��.Enabled = True: Me.cmdƵ��.Enabled = True
    Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    If Not IsNull(rsTmp("�ܸ�����")) Then Me.txt���� = rsTmp("�ܸ�����"): Me.txt����.Enabled = True
    If Not IsNull(rsTmp("��������")) Then Me.txt���� = rsTmp("��������"): Me.txt����.Enabled = True: Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    Me.cboִ�п���.Clear: Me.cboִ�п���.AddItem rsTmp("���ұ���") & "-" & rsTmp("��������")
    Me.cboִ�п���.Text = rsTmp("���ұ���") & "-" & rsTmp("��������"): Me.cboִ�п���.Enabled = True
    Me.cboҽ��.Clear: Me.cboҽ��.AddItem rsTmp("����ҽ��")
    Me.cboҽ��.Text = rsTmp("����ҽ��"): Me.cboҽ��.Enabled = True
    Me.picAdvice.Enabled = False

    SetItemFormat
    prbRefresh.Value = 15
    '��ʼ������

    '�ж��ܷ�༭����
    bAllowEdit = False
    iCurrElementIndex = 1
    
    Me.MousePointer = vbHourglass
    ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 2, prbRefresh, , , , blnMoved
    ProFile1(0).SetActiveElement 1
    Me.MousePointer = vbDefault
End Sub

Private Sub ClearForm()
    On Error Resume Next
    Me.txt���� = "": Me.txt���� = "": Me.txt��ʼʱ�� = "": Me.txtƵ�� = "": Me.txtҽ������ = ""
    Me.txt���� = "": Me.txtҽ������ = "": Me.chk���� = 0: Me.chk��ʼʱ�� = 0: Me.cboҽ��.ListIndex = -1: Me.cboִ�п���.ListIndex = -1
    
    Me.MousePointer = vbHourglass
    ProFile1(0).ShowFile "", "", "", 10, "0", False
    Me.MousePointer = vbDefault
End Sub

Private Sub SetItemFormat()   '����������Ŀ������ʾ��ʽ
    Select Case intType
        Case 0
            Me.lblҽ������.Caption = "�����Ŀ": Me.lbl����.Caption = "��鲿λ": Me.cmdExt.ToolTipText = "ѡ���鲿λ"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
        Case 1
            Me.lblҽ������.Caption = "������Ŀ": Me.lbl����.Caption = "����ʽ": Me.cmdExt.ToolTipText = "ѡ������ʽ"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
        Case 3
            Me.lblҽ������.Caption = "������Ŀ": Me.lbl����.Caption = "����걾": Me.cmdExt.ToolTipText = "ѡ�����걾"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
        Case Else
            Me.lbl����.Visible = False: Me.txt����.Visible = False: Me.cmdExt.Visible = False
    End Select
End Sub

Private Sub Form_Load()
    ProFile1(0).ifShowDiagItem = False
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    Dim lngTxtWidth As Single
    Dim lngDistance As Single
    
    If WindowState = 1 Then Exit Sub
    lngDistance = 300
    
    On Error Resume Next
    With picDoc
        .Left = 0: .Top = 0
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - .Top
    End With
    With picAdvice
        .Left = 0: .Top = 0
        .Width = picDoc.ScaleWidth
    End With
    With lineSplit
        .X2 = picAdvice.Width + .X1
    End With
    With Me.chk����
        .Left = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Width
        If .Left < Me.txt��ʼʱ��.Left + Me.txt��ʼʱ��.Width + lngDistance Then .Left = Me.txt��ʼʱ��.Left + Me.txt��ʼʱ��.Width + lngDistance
    End With
    
    lngTxtWidth = (picAdvice.ScaleWidth - Me.lbl��ʼʱ��.Left - Me.cmdSel.Width - Me.txtҽ������.Left - lngDistance - _
        Me.lbl����.Width - Me.cmdExt.Width - 60) / 2
    With Me.txtҽ������
        .Width = lngTxtWidth
        Me.cmdSel.Left = .Left + .Width
        Me.lbl����.Left = Me.cmdSel.Left + Me.cmdSel.Width + lngDistance
    End With
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
        Me.cmdExt.Left = .Left + .Width
    End With
    Me.lineTitleSplit.X2 = Me.cmdExt.Left + Me.cmdExt.Width + 200

    With Me.txtҽ������
        .Width = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Left
    End With
    
    lngTxtWidth = (picAdvice.Width - Me.lbl��ʼʱ��.Left - Me.txtƵ��.Left - Me.txtƵ��.Width - _
        (Me.lbl������λ.Width + Me.lbl����.Width + lngDistance + 2 * 30) - _
        (Me.lbl������λ.Width + Me.lbl����.Width + lngDistance + 2 * 30)) / 2
    If lngTxtWidth < 1000 Then lngTxtWidth = 1000
    Me.lbl����.Left = Me.txtƵ��.Left + Me.txtƵ��.Width + lngDistance
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl������λ.Left = Me.txt����.Left + Me.txt����.Width + 30
    Me.lbl����.Left = Me.lbl������λ.Left + Me.lbl������λ.Width + lngDistance
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl������λ.Left = Me.txt����.Left + Me.txt����.Width + 30
    
    With Me.cboҽ��
        .Left = Me.txt����.Left
        .Width = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Left
    End With
    Me.lbl����ҽ��.Left = Me.cboҽ��.Left - Me.lbl����ҽ��.Width
    
    With picFile
        .Left = 0: .Top = picAdvice.Top + picAdvice.Height
        .Width = picDoc.ScaleWidth
        .Height = picDoc.ScaleHeight - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub picFile_Resize()
    On Error Resume Next
    With ProFile1(iTabIndex)
        .Left = 0: .Top = 0
        .Width = picFile.ScaleWidth
        .Height = picFile.ScaleHeight
        
        If .Width > picFile.ScaleWidth Then Me.Width = .Width
        If .Height > picFile.ScaleHeight Then Me.Height = .Height + picFile.Top
    End With
End Sub

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal strDataIDs As String)
'���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
'      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
'      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '���¼��벿λ�л򸽼������м�������Ŀ��
    If int���� = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    End If
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,����,����,nvl(�걾��λ,' ') As �걾��λ," + _
        "���,nvl(�Ƽ�����,0) As �Ƽ�����,nvl(ִ�п���,0) As ִ�п��� From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
        OpenRecord rsRelativeAdvice, strSQL, Me.Caption
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int���� = 1 Then Get����������� = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") <> "G" Then
                strTmp = strTmp & "," & rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get����������� = txtMainAdvice & " �� " & Mid(strTmp, 2)
    Else
        Get����������� = txtMainAdvice
    End If
End Function

Private Function Get��������() As String
    If rsRelativeAdvice Is Nothing Then Get�������� = "": Exit Function
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") = "G" Then
                Get�������� = rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
End Function

Private Function Get��λ����() As String
    If rsRelativeAdvice Is Nothing Then Get��λ���� = "": Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("�걾��λ"))) > 0 Then
            Get��λ���� = Get��λ���� & "," & rsRelativeAdvice("�걾��λ")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    If Len(Get��λ����) > 0 Then Get��λ���� = Mid(Get��λ����, 2)
End Function

Private Function GetItemField(ByVal lng��ĿID As Long, ByVal strField As String) As Variant
'���ܣ���ȡָ��������Ŀ��ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
