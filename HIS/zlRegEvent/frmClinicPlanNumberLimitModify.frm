VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanNumberLimitModify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Ӻ�"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10335
   Icon            =   "frmClinicPlanNumberLimitModify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   9030
      TabIndex        =   39
      Top             =   6810
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9030
      TabIndex        =   37
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9030
      TabIndex        =   38
      Top             =   810
      Width           =   1100
   End
   Begin VB.Frame fra��Դ��Ϣ 
      Caption         =   "��Դ������Ϣ"
      Height          =   1035
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   8835
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   10
         Top             =   645
         Width           =   3105
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4950
         TabIndex        =   6
         Top             =   285
         Width           =   1605
      End
      Begin VB.TextBox txtDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7290
         TabIndex        =   8
         Top             =   285
         Width           =   1365
      End
      Begin VB.TextBox txt���տ��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4950
         TabIndex        =   12
         Top             =   645
         Width           =   1605
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�Һ�ʱ���뽨��"
         Enabled         =   0   'False
         Height          =   180
         Left            =   6900
         TabIndex        =   13
         Top             =   705
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtSignalNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   2
         Top             =   285
         Width           =   1035
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2670
         TabIndex        =   4
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label lbl���տ��� 
         AutoSize        =   -1  'True
         Caption         =   "���տ���"
         Height          =   180
         Left            =   4200
         TabIndex        =   11
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   6900
         TabIndex        =   7
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   4560
         TabIndex        =   5
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         Height          =   180
         Left            =   450
         TabIndex        =   9
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2280
         TabIndex        =   3
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   450
         TabIndex        =   1
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ"
      Height          =   1035
      Left            =   30
      TabIndex        =   14
      Top             =   1140
      Width           =   8835
      Begin VB.TextBox txt�ϰ�ʱ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   18
         Top             =   285
         Width           =   1065
      End
      Begin VB.TextBox txtԤԼ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4950
         TabIndex        =   20
         Top             =   285
         Width           =   1605
      End
      Begin VB.CheckBox chkʱ�� 
         Caption         =   "����ʱ��"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4230
         TabIndex        =   24
         Top             =   698
         Width           =   1035
      End
      Begin VB.CheckBox chk��ſ��� 
         Caption         =   "������ſ���"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2130
         TabIndex        =   23
         Top             =   698
         Width           =   1395
      End
      Begin VB.TextBox txt����ҽ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   22
         Top             =   660
         Width           =   1065
      End
      Begin VB.TextBox txt�������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   16
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label lblԤԼ���� 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����"
         Height          =   180
         Left            =   4200
         TabIndex        =   19
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl�ϰ�ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "�ϰ�ʱ��"
         Height          =   180
         Left            =   2130
         TabIndex        =   17
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl����ҽ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ҽ��"
         Height          =   180
         Left            =   90
         TabIndex        =   21
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   345
         Width           =   720
      End
   End
   Begin zl9RegEvent.ClinicPlanWorkTimeNum cpWorkTimeNum 
      Height          =   4485
      Left            =   30
      TabIndex        =   36
      Top             =   2940
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   7911
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IsDataChanged   =   -1  'True
   End
   Begin VB.Frame fraLimitInfo 
      Caption         =   "�޺���Ϣ"
      Height          =   675
      Left            =   30
      TabIndex        =   25
      Top             =   2220
      Width           =   8835
      Begin VB.TextBox txtAdd��Լ�� 
         Height          =   300
         Left            =   7170
         MaxLength       =   6
         TabIndex        =   34
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txtAdd�޺��� 
         Height          =   300
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   29
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txt��Լ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   32
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txt�޺��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   27
         Top             =   285
         Width           =   1095
      End
      Begin MSComCtl2.UpDown upd�޺� 
         Height          =   285
         Left            =   3960
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   293
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtAdd�޺���"
         BuddyDispid     =   196640
         OrigLeft        =   1200
         OrigTop         =   120
         OrigRight       =   1455
         OrigBottom      =   420
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upd��Լ 
         Height          =   285
         Left            =   8250
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   293
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtAdd��Լ��"
         BuddyDispid     =   196639
         OrigLeft        =   1200
         OrigTop         =   120
         OrigRight       =   1455
         OrigBottom      =   420
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblAdd��Լ�� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   6420
         TabIndex        =   33
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lblAdd�޺��� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   2130
         TabIndex        =   28
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl��Լ�� 
         AutoSize        =   -1  'True
         Caption         =   "��Լ��"
         Height          =   180
         Left            =   4590
         TabIndex        =   31
         Top             =   345
         Width           =   540
      End
      Begin VB.Label lbl�޺��� 
         AutoSize        =   -1  'True
         Caption         =   "�޺���"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   345
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmClinicPlanNumberLimitModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As Byte '1-�Ӻţ�2-����
Private mobj��Դ As �����Դ, mobj�����¼ As �����¼

Private mblnOk As Boolean
Private mlngMinSN As Long '����ɾ����ʱ�ε���С��ţ���ʱ�Σ�������ţ�
Private mblnNotChanged As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As Byte, _
    ByVal obj��Դ As �����Դ, ByVal obj�����¼ As �����¼) As Boolean
    
    If obj��Դ Is Nothing Then Exit Function
    If obj�����¼ Is Nothing Then Exit Function
    
    mbytFun = bytFun
    Set mobj��Դ = obj��Դ: Set mobj�����¼ = obj�����¼
    
    If CheckDepend = False Then Exit Function
    mblnOk = False
    On Error Resume Next
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Function CheckDepend() As Boolean
    '����:�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    '���ܶ���ʷ�İ��Ž��в���
    If DateDiff("s", mobj�����¼.��ֹʱ��, zlDatabase.Currentdate) >= 0 Then
        MsgBox "��ǰϵͳʱ���Ѵ����˰���ʱ�ε���ֹʱ�䣬���ܽ���" & IIf(mbytFun = 1, "�Ӻ�", "����") & "������", vbInformation, gstrSysName
        Exit Function
    End If
    '���޺����Ĳ��ܵ���
    If mobj�����¼.�޺��� = 0 Then
        MsgBox "��ǰ����ʱ��Ϊ���޺ţ����ܽ���" & IIf(mbytFun = 1, "�Ӻ�", "����") & "������", vbInformation, gstrSysName
        Exit Function
    End If
    '�Ѿ�ͣ���δ���ﰲ�ŵģ�������Ӻ�/����
    strSQL = "Select 1 from �ٴ������¼ Where ID=[1] and �ϰ�ʱ��=[2] And ͣ�￪ʼʱ�� Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������¼", mobj�����¼.��¼ID, mobj�����¼.ʱ���)
    If rsTemp.EOF Then
        MsgBox "��ǰ����ʱ�β����ڻ���ͣ����ܽ���" & IIf(mbytFun = 1, "�Ӻ�", "����") & "������", vbInformation, gstrSysName
        Exit Function
    End If
    '�޺����Ѿ���ȫ��ʹ�õģ����������
    If mobj�����¼.�ѹ��� >= mobj�����¼.�޺��� And mbytFun = 2 Then
        MsgBox "��ǰ����ʱ����ȫ���Һţ����ܽ��м��Ų�����", vbInformation, gstrSysName
        Exit Function
    End If
    CheckDepend = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    If cpWorkTimeNum.�޺��� < 1 Then
        MsgBox IIf(mbytFun = 1, "", "ʣ��") & "�޺���(" & cpWorkTimeNum.�޺��� & ")С����1��", vbInformation, gstrSysName
        Exit Sub
    End If
    If cpWorkTimeNum.�޺��� < mlngMinSN Then
        MsgBox IIf(mbytFun = 1, "", "ʣ��") & "�޺���(" & cpWorkTimeNum.�޺��� & ")С������ʹ��ʱ�ε�������(" & mlngMinSN & ")��", vbInformation, gstrSysName
        Exit Sub
    End If
    If cpWorkTimeNum.��Լ�� > cpWorkTimeNum.�޺��� Then
        MsgBox IIf(mbytFun = 1, "", "ʣ��") & "��Լ��(" & cpWorkTimeNum.��Լ�� & ")������" & IIf(mbytFun = 1, "", "ʣ��") & "�޺���(" & cpWorkTimeNum.�޺��� & ")��", vbInformation, gstrSysName
        Exit Sub
    End If
    If cpWorkTimeNum.�޺��� < mobj�����¼.�ѹ��� Then
        MsgBox IIf(mbytFun = 1, "", "ʣ��") & "�޺���(" & cpWorkTimeNum.�޺��� & ")С�����ѹ���(" & mobj�����¼.�ѹ��� & ")��", vbInformation, gstrSysName
        Exit Sub
    End If
    'ԤԼ����:0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
    If mobj�����¼.ԤԼ���� <> 1 And cpWorkTimeNum.��Լ�� < mobj�����¼.��Լ�� Then
        MsgBox IIf(mbytFun = 1, "", "ʣ��") & "��Լ��(" & cpWorkTimeNum.��Լ�� & ")С������Լ��(" & mobj�����¼.��Լ�� & ")��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtAdd�޺���.Text) = 0 And Val(txtAdd��Լ��.Text) = 0 And cpWorkTimeNum.IsDataChanged = False Then
        If MsgBox("����δ�����κε���������Ҫ���棡Ҫ�˳�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Unload Me
        End If
        Exit Sub
    End If
    If mbytFun = 2 And mobj�����¼.��Լ�� > 0 And cpWorkTimeNum.��Լ�� = 0 Then
        If MsgBox("ʣ����Լ��Ϊ0��ʾ��ֹԤԼ����ȷ��Ҫ�Ըó��ﰲ�Ž��н�ֹԤԼ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    If cpWorkTimeNum.IsValied() = False Then Exit Sub
    
    If SaveData() = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    Call InitData
    Call SetEnabledBackColor(Me.Controls)
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function InitData() As Boolean
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    Select Case mbytFun
    Case 1 '�Ӻ�
        Me.Caption = "�Ӻ�"
        lblAdd��Լ��.Caption = "��������"
        lblAdd�޺���.Caption = "��������"
    Case 2 '����
        Me.Caption = "����"
        lblAdd��Լ��.Caption = "���μ���"
        lblAdd�޺���.Caption = "���μ���"
    End Select
    
    '��Դ��Ϣ
    txtSignalNO.Text = mobj��Դ.����
    txt����.Text = mobj��Դ.����
    txtDept.Text = mobj��Դ.��������
    txtItem.Text = mobj��Դ.��Ŀ����
    txtDoctor.Text = mobj��Դ.ҽ������
    txt���տ���.Text = Decode(mobj��Դ.���տ���״̬, 1, "����ԤԼ", 2, "��ֹԤԼ", 3, "�ܽڼ������ÿ���", "���ϰ�")
    chk����.Value = IIf(mobj��Դ.�Ƿ񽨲���, vbChecked, vbUnchecked)
    
    If IsDate(mobj�����¼.��������) Then
        txt��������.Text = Format(mobj�����¼.��������, "yyyy-mm-dd")
    Else
        txt��������.Text = mobj�����¼.��������
    End If
    txt�ϰ�ʱ��.Text = mobj�����¼.ʱ���
    txt����ҽ��.Text = mobj�����¼.����ҽ��
    txtԤԼ����.Text = Choose(mobj�����¼.ԤԼ���� + 1, "����ԤԼ", "��ֹԤԼ", "����ֹ��������ԤԼ")
    chk��ſ���.Value = IIf(mobj�����¼.�Ƿ���ſ���, vbChecked, vbUnchecked)
    chkʱ��.Value = IIf(mobj�����¼.�Ƿ��ʱ��, vbChecked, vbUnchecked)
    txt�޺���.Text = IIf(mobj�����¼.�޺��� <> 0, mobj�����¼.�޺���, "")
    txt��Լ��.Text = IIf(mobj�����¼.��Լ�� <> 0, mobj�����¼.��Լ��, "")
    
    '��ֹԤԼ�Ͳ�����ԤԼ�Ĳ������޸�ԤԼ��
    txtAdd��Լ��.Enabled = Not (mobj�����¼.ԤԼ���� = 1 Or mobj�����¼.��Լ�� = 0)
    upd��Լ.Enabled = txtAdd��Լ��.Enabled
    
    cpWorkTimeNum.CanReCalic = mobj�����¼.��Լ�� = 0 And mobj�����¼.�ѹ��� = 0 And mobj�����¼.�Ƿ��ʱ��
    cpWorkTimeNum.EditMode = ED_RegistPlan_NumLimitModify

    If cpWorkTimeNum.CanReCalic = False Or mobj�����¼.�Ƿ��ʱ�� Then
        '�����Щ�������޸�
        Dim strSQL As String, rsTemp As ADODB.Recordset
        Dim cllFixedSN As New Collection
        strSQL = "Select a.���, Nvl(Sum(a.����), 0) As ��Լ��" & vbNewLine & _
                " From �ٴ�������ſ��� A" & vbNewLine & _
                " Where a.��¼id = [1] And Nvl(a.�Һ�״̬, 0) <> 0" & vbNewLine & _
                " Group By a.���" & vbNewLine & _
                " Order By a.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobj�����¼.��¼ID)
        Do While Not rsTemp.EOF
            cllFixedSN.Add Array(Val(Nvl(rsTemp!���)), Val(Nvl(rsTemp!��Լ��)))
            If mobj�����¼.�Ƿ���ſ��� Then
                If Val(Nvl(rsTemp!���)) > mlngMinSN Then mlngMinSN = Val(Nvl(rsTemp!���))
            End If
            rsTemp.MoveNext
        Loop
    End If
    InitData = cpWorkTimeNum.LoadData(mobj�����¼.������Ϣ��.Clone, mobj�����¼.�ϰ�ʱ��, cllFixedSN)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtAdd�޺���_Change()
    If mblnNotChanged = False Then
        If mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) < 1 Then
            MsgBox IIf(mbytFun = 1, "", "ʣ��") & "�޺���(" & mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) & ")����С��1��", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd�޺���.Text = Val(txtAdd�޺���.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
        If mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) < mlngMinSN Then
            MsgBox IIf(mbytFun = 1, "", "ʣ��") & "�޺���(" & mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) & ")����С����ʹ��ʱ�ε�������(" & mlngMinSN & ")��", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd�޺���.Text = Val(txtAdd�޺���.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
        If mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) < mobj�����¼.�ѹ��� Then
            MsgBox IIf(mbytFun = 1, "", "ʣ��") & "�޺���(" & mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) & ")����С���ѹ���(" & mobj�����¼.�ѹ��� & ")��", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd�޺���.Text = Val(txtAdd�޺���.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
    End If
    mblnNotChanged = True
    txtAdd�޺���.Text = IIf(Val(txtAdd�޺���.Text) = 0, "", txtAdd�޺���.Text)
    mblnNotChanged = False
    
    If Val(txtAdd�޺���.Tag) <> Val(txtAdd�޺���.Text) Then
        txtAdd�޺���.Tag = Val(txtAdd�޺���.Text)
        cpWorkTimeNum.SetNewSN mobj�����¼.������Ϣ��.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text), _
            IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text), mbytFun = 1
    End If
    
    If mobj�����¼.������Ϣ��.��Լ�� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd��Լ��.Text) > mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) Then
        mblnNotChanged = True
        txtAdd��Լ��.Text = Abs(mobj�����¼.�޺��� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd�޺���.Text) - mobj�����¼.������Ϣ��.��Լ��)
        txtAdd��Լ��.Text = IIf(Val(txtAdd��Լ��.Text) = 0, "", txtAdd��Լ��.Text)
        mblnNotChanged = False
    End If
End Sub

Private Sub txtAdd�޺���_GotFocus()
    zlControl.TxtSelAll txtAdd�޺���
End Sub

Private Sub txtAdd�޺���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtAdd��Լ��_Change()
    If mblnNotChanged = False Then
        If mobj�����¼.������Ϣ��.��Լ�� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd��Լ��.Text) > cpWorkTimeNum.�޺��� Then
            MsgBox IIf(mbytFun = 1, "", "ʣ��") & "��Լ��(" & mobj�����¼.��Լ�� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd��Լ��.Text) & ")���ܴ����޺���(" & cpWorkTimeNum.�޺��� & ")��", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd��Լ��.Text = Val(txtAdd��Լ��.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
        If mobj�����¼.��Լ�� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd��Լ��.Text) < mobj�����¼.��Լ�� Then
            MsgBox IIf(mbytFun = 1, "", "ʣ��") & "��Լ��(" & mobj�����¼.��Լ�� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd��Լ��.Text) & ")����С����Լ��(" & mobj�����¼.��Լ�� & ")��", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd��Լ��.Text = Val(txtAdd��Լ��.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
    End If
    mblnNotChanged = True
    txtAdd��Լ��.Text = IIf(Val(txtAdd��Լ��.Text) = 0, "", txtAdd��Լ��.Text)
    mblnNotChanged = False
    
    If Val(txtAdd��Լ��.Tag) <> Val(txtAdd��Լ��.Text) Then
        txtAdd��Լ��.Tag = Val(txtAdd��Լ��.Text)
        cpWorkTimeNum.��Լ�� = mobj�����¼.������Ϣ��.��Լ�� + IIf(mbytFun = 1, 1, -1) * Val(txtAdd��Լ��.Text)
    End If
End Sub

Private Sub txtAdd��Լ��_GotFocus()
    zlControl.TxtSelAll txtAdd��Լ��
End Sub

Private Sub txtAdd��Լ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, cllPro As New Collection, i As Integer
    Dim obj���� As ������Ϣ
    Dim cll���� As Collection, str���� As String, strTemp As String
    Dim blnTrans As Boolean
    
    Err = 0: On Error GoTo errHandler
    Set mobj�����¼.������Ϣ�� = cpWorkTimeNum.Get����
    '����䶯��¼
    'Zl_�ٴ�������ſ��Ʊ䶯(
    strSQL = "Zl_�ٴ�������ſ��Ʊ䶯("
    '��¼id_In     �ٴ�����䶯��¼.��¼id%Type,
    strSQL = strSQL & "" & mobj�����¼.��¼ID & ","
    '�޺���_In     �ٴ������¼.�޺���%Type,
    strSQL = strSQL & "" & mobj�����¼.������Ϣ��.�޺��� & ","
    '��Լ��_In     �ٴ������¼.��Լ��%Type,
    strSQL = strSQL & "" & mobj�����¼.������Ϣ��.��Լ�� & ","
    'ԭ�ѹ���_In   �ٴ������¼.�ѹ���%Type,
    strSQL = strSQL & "" & mobj�����¼.�ѹ��� & ","
    'ԭ��Լ��_In   �ٴ������¼.��Լ��%Type,
    strSQL = strSQL & "" & mobj�����¼.��Լ�� & ","
    '����Ա����_In �ٴ�����䶯��¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '�Ǽ�ʱ��_In   �ٴ�����䶯��¼.�Ǽ�ʱ��%Type := Null
    strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    cllPro.Add strSQL
    
    Set cll���� = New Collection
    For Each obj���� In mobj�����¼.������Ϣ��
        strTemp = obj����.��� & "," & _
            GetWorkTrueDate(mobj�����¼.��ʼʱ��, ZDate(obj����.��ʼʱ��, mobj�����¼.��ʼʱ��, False), , False) & "," & _
            GetWorkTrueDate(mobj�����¼.��ʼʱ��, ZDate(obj����.��ֹʱ��, mobj�����¼.��ֹʱ��, False)) & "," & _
            obj����.���� & "," & IIf(obj����.�Ƿ�ԤԼ, 1, 0)
        If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
            'ʱ��_In:���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־|...
            str���� = Mid(str����, 2)
            cll����.Add str����
            str���� = ""
        End If
        str���� = str���� & "|" & strTemp
    Next
    If str���� <> "" Then
        str���� = Mid(str����, 2)
        cll����.Add str����
    End If
    For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
        'Zl_�ٴ�������ſ���_Update(
        strSQL = "Zl_�ٴ�������ſ���_Update("
        '��¼id_In   �ٴ������¼.Id%Type,
        strSQL = strSQL & "" & mobj�����¼.��¼ID & ","
        'ʱ��_In     Varchar2 := Null,--���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־|...
        str���� = ""
        If cll����.Count > 0 Then str���� = cll����(i)
        strSQL = strSQL & "'" & str���� & "',"
        'ɾ�����_In Number:=0 --�Ƿ�ɾ���������ʱ��
        strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
        cllPro.Add strSQL
    Next
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    SaveData = True
    Exit Function
errHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

