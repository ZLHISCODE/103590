VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMedicalStationPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   Icon            =   "frmMedicalStationPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5265
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   5085
      Left            =   60
      TabIndex        =   38
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   8969
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&1.���"
      TabPicture(0)   =   "frmMedicalStationPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2.�Ƿ�"
      TabPicture(1)   =   "frmMedicalStationPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "fraAction"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "Frame4"
      Tab(1).Control(5)=   "lst�շ����"
      Tab(1).Control(6)=   "Label1"
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame7 
         Caption         =   "������췶Χ"
         Height          =   1095
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   5475
         Begin VB.CheckBox chk 
            Caption         =   "������ʱ���(&8)"
            Height          =   240
            Index           =   2
            Left            =   3795
            TabIndex        =   50
            Top             =   750
            Width           =   1650
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   1875
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   690
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   1875
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   330
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���������ʱ��(&7)"
            Height          =   180
            Index           =   6
            Left            =   285
            TabIndex        =   52
            Top             =   750
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��������ʱ��(&6)"
            Height          =   180
            Index           =   2
            Left            =   285
            TabIndex        =   51
            Top             =   375
            Width           =   1530
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "����"
         Height          =   1695
         Left            =   120
         TabIndex        =   46
         Top             =   3225
         Width           =   5475
         Begin VB.CheckBox chk 
            Caption         =   "�����Ա���ܻ򱨵�ʱ�Զ���ӡָ����(&D)"
            Height          =   270
            Index           =   1
            Left            =   315
            TabIndex        =   12
            Top             =   1125
            Width           =   3810
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������������ʱ��ʾ���С��(&A)"
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   8
            Top             =   540
            Width           =   3375
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   1
            Left            =   3135
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "5"
            Top             =   780
            Width           =   345
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   1845
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   210
            Width           =   1920
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   1
            Left            =   3480
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   780
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txt(1)"
            BuddyDispid     =   196616
            BuddyIndex      =   1
            OrigLeft        =   4320
            OrigTop         =   1065
            OrigRight       =   4560
            OrigBottom      =   1365
            Max             =   30
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����Ա���ܻ򱨵�ʱ�Զ���ӡ���뵥(&E)"
            Height          =   255
            Index           =   3
            Left            =   300
            TabIndex        =   13
            Top             =   1410
            Width           =   3810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����졢��������Զ�ˢ�¼��(&B)         ��"
            Height          =   180
            Index           =   4
            Left            =   300
            TabIndex        =   9
            Top             =   840
            Width           =   3780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����÷ѱ�(&9)"
            Height          =   180
            Index           =   5
            Left            =   330
            TabIndex        =   6
            Top             =   285
            Width           =   1350
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����"
         Height          =   1440
         Left            =   -71685
         TabIndex        =   44
         Top             =   3540
         Width           =   2280
         Begin VB.CheckBox chkPay 
            Caption         =   "��ҩ�������븶��(&P)"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   540
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkTime 
            Caption         =   "���������������(&N)"
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   255
            Width           =   2055
         End
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "��ʾ����ҩ����(&T)"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   1095
            Width           =   2055
         End
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "��ʾ����ҩ�����(&R)"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   810
            Width           =   2055
         End
      End
      Begin VB.Frame fraAction 
         Caption         =   "ִ������ "
         Height          =   840
         Left            =   -74865
         TabIndex        =   43
         Top             =   4140
         Width           =   3135
         Begin VB.CheckBox chkFinish 
            Caption         =   "�������δ�շѲ��˵���Ŀ(&L)"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   540
            Width           =   2805
         End
         Begin VB.CheckBox chkActLog 
            Caption         =   "�������˴���ִ�м�¼(&K)"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   285
            Width           =   2745
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "סԺҩ������ "
         Height          =   1305
         Left            =   -74880
         TabIndex        =   42
         Top             =   1830
         Width           =   3135
         Begin VB.ComboBox cboס��ҩ 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   930
            Width           =   1950
         End
         Begin VB.ComboBox cboס��ҩ 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   240
            Width           =   1950
         End
         Begin VB.ComboBox cboס��ҩ 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   585
            Width           =   1950
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�в�ҩ(&G)"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   990
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ(&E)"
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�г�ҩ(&F)"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   645
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "����ҩ������ "
         Height          =   1320
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   3135
         Begin VB.ComboBox cbo�ų�ҩ 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   600
            Width           =   1950
         End
         Begin VB.ComboBox cbo����ҩ 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   255
            Width           =   1950
         End
         Begin VB.ComboBox cbo����ҩ 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   945
            Width           =   1950
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�г�ҩ(&B)"
            Height          =   180
            Left            =   105
            TabIndex        =   16
            Top             =   660
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ(&A)"
            Height          =   180
            Left            =   105
            TabIndex        =   14
            Top             =   315
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�в�ҩ(&D)"
            Height          =   180
            Left            =   105
            TabIndex        =   18
            Top             =   1005
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "ҩƷ��λ "
         Height          =   900
         Left            =   -74865
         TabIndex        =   40
         Top             =   3195
         Width           =   3135
         Begin VB.OptionButton optҩƷ��λ 
            Caption         =   "�ۼ۵�λ(I)"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   26
            Top             =   315
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optҩƷ��λ 
            Caption         =   "����/סԺ��λ(&J)"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   27
            Top             =   600
            Width           =   1845
         End
      End
      Begin VB.ListBox lst�շ���� 
         Height          =   2790
         Left            =   -71670
         Style           =   1  'Checkbox
         TabIndex        =   31
         ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
         Top             =   660
         Width           =   2265
      End
      Begin VB.Frame Frame5 
         Caption         =   "ʱ�䷶Χ"
         Height          =   1440
         Left            =   120
         TabIndex        =   39
         Top             =   570
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1035
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   675
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   330
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������췶Χ(&5)"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   1110
            Width           =   1350
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����ȱʡ��Χ(&4)"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   2
            Top             =   750
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����ȱʡ��Χ(&3)"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   0
            Top             =   405
            Width           =   1530
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������(&M)"
         Height          =   180
         Left            =   -71685
         TabIndex        =   30
         Top             =   435
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   36
      Top             =   5265
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   37
      Top             =   5265
      Width           =   1100
   End
End
Attribute VB_Name = "frmMedicalStationPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mlngLoop As Long
Private mfrmMain As Object

Public Function ShowPara(ByVal frmMain As Object) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objCbo As ComboBox, lngҩ��ID As Long
    Dim strSQL As String, strPar As String, i As Long
    Dim rs As New ADODB.Recordset
    
    mblnOK = False
    
    Set mfrmMain = frmMain
    '��ʼ��
    
    For mlngLoop = 0 To 5
        If mlngLoop <> 4 Then
            cbo(mlngLoop).AddItem "��  ��"
            cbo(mlngLoop).AddItem "��  ��"
            cbo(mlngLoop).AddItem "��  ��"
            cbo(mlngLoop).AddItem "��  ��"
            cbo(mlngLoop).AddItem "��  ��"
            cbo(mlngLoop).AddItem "������"
            cbo(mlngLoop).AddItem "��  ��"
            cbo(mlngLoop).AddItem "ǰ����"
            cbo(mlngLoop).AddItem "ǰһ��"
            cbo(mlngLoop).AddItem "ǰ����"
            cbo(mlngLoop).AddItem "ǰһ��"
            cbo(mlngLoop).AddItem "ǰ����"
            cbo(mlngLoop).AddItem "ǰ����"
            cbo(mlngLoop).AddItem "ǰ����"
            cbo(mlngLoop).AddItem "ǰһ��"
            cbo(mlngLoop).AddItem "ǰ����"
        End If
    Next
    
    cbo(4).Clear
    cbo(4).AddItem ""
    gstrSQL = "Select ����,1 As ID From �ѱ�"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Call AddComboData(cbo(4), rs, False)
    End If
    
    On Error Resume Next
    
    '������������õķѱ�����
    
    cbo(4).Text = Trim(zlDatabase.GetPara(130, glngSys, , ""))
    chk(0).Value = Val(zlDatabase.GetPara(131, glngSys, , "0"))
    
    cbo(0).Text = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�����ʱ�䷶Χ", "��  ��")
    cbo(1).Text = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�����ʱ�䷶Χ", "��  ��")
    
    cbo(5).Text = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "���������ʱ�䷶Χ", "��  ��")
    chk(2).Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "������ѯ����", "0"))
    
    cbo(2).Text = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�������ʱ�䷶Χ", "��  ��")
    cbo(3).Text = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "������췶Χ", "��  ��")
    txt(1).Text = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ�ˢ�¼��", 5))
    chk(1).Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡָ����", 0))
    chk(3).Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡ���뵥", 0))
    
    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    If cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
    
    
    chkPay.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9CISWork", "��ҩ����", 1))
    chkTime.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9CISWork", "�������", 0))
    chkҩ��.Value = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "��ʾ����ҩ�����", 0))
    chkҩ��.Value = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "��ʾ����ҩ����", 0))
    
    'ҩƷ��λ
    i = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "ҩƷ��λ", 0))
    optҩƷ��λ(IIf(i = 0, 0, 1)).Value = True
    
    'ȱʡҩ��
    cbo����ҩ.AddItem "ϵͳ����": cbo����ҩ.ListIndex = 0
    cbo�ų�ҩ.AddItem "ϵͳ����": cbo�ų�ҩ.ListIndex = 0
    cbo����ҩ.AddItem "ϵͳ����": cbo����ҩ.ListIndex = 0
    cboס��ҩ.AddItem "ϵͳ����": cboס��ҩ.ListIndex = 0
    cboס��ҩ.AddItem "ϵͳ����": cboס��ҩ.ListIndex = 0
    cboס��ҩ.AddItem "ϵͳ����": cboס��ҩ.ListIndex = 0
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.��������,B.�������" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " Order by A.����"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo����ҩ, IIf(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo�ų�ҩ, IIf(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo����ҩ, IIf(rsTmp!������� = 2, cboס��ҩ, Nothing))
        End If
        If objCbo Is Nothing Then
            If rsTmp!�������� = "��ҩ��" Then
                cbo����ҩ.AddItem rsTmp!����
                cbo����ҩ.ItemData(cbo����ҩ.NewIndex) = rsTmp!ID
                cboס��ҩ.AddItem rsTmp!����
                cboס��ҩ.ItemData(cboס��ҩ.NewIndex) = rsTmp!ID
            ElseIf rsTmp!�������� = "��ҩ��" Then
                cbo�ų�ҩ.AddItem rsTmp!����
                cbo�ų�ҩ.ItemData(cbo�ų�ҩ.NewIndex) = rsTmp!ID
                cboס��ҩ.AddItem rsTmp!����
                cboס��ҩ.ItemData(cboס��ҩ.NewIndex) = rsTmp!ID
            ElseIf rsTmp!�������� = "��ҩ��" Then
                cbo����ҩ.AddItem rsTmp!����
                cbo����ҩ.ItemData(cbo����ҩ.NewIndex) = rsTmp!ID
                cboס��ҩ.AddItem rsTmp!����
                cboס��ҩ.ItemData(cboס��ҩ.NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "����ȱʡ��ҩ��", 0))
    Call zlControl.CboLocate(cbo����ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "����ȱʡ��ҩ��", 0))
    Call zlControl.CboLocate(cbo�ų�ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "����ȱʡ��ҩ��", 0))
    Call zlControl.CboLocate(cbo����ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "סԺȱʡ��ҩ��", 0))
    Call zlControl.CboLocate(cboס��ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "סԺȱʡ��ҩ��", 0))
    Call zlControl.CboLocate(cboס��ҩ, lngҩ��ID, True)
    lngҩ��ID = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "סԺȱʡ��ҩ��", 0))
    Call zlControl.CboLocate(cboס��ҩ, lngҩ��ID, True)
    
    '�շ����
    strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst�շ����.AddItem rsTmp!���
        lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
        rsTmp.MoveNext
    Loop
    strPar = GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "�շ����", "")
    If strPar = "" Then
        For i = 0 To lst�շ����.ListCount - 1
            lst�շ����.Selected(i) = True
        Next
    Else
        For i = 0 To lst�շ����.ListCount - 1
            If InStr(strPar, Chr(lst�շ����.ItemData(i))) Then lst�շ����.Selected(i) = True
        Next
    End If
    If lst�շ����.ListCount > 0 Then lst�շ����.TopIndex = 0: lst�շ����.ListIndex = 0
    
    '�Ƿ��������ִ�м�¼
    chkActLog.Value = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "����ִ�м�¼", 0))
    
    '�Ƿ��������δ�շѲ��˵���Ŀ
    chkFinish.Value = Val(GetSetting("ZLSOFT", "����ģ��\zl9CISWork", "δ�շ����", 0))

    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo�ų�ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo����ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo����ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboס��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboס��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboס��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkActLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkFinish_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�����ʱ�䷶Χ", cbo(0).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�����ʱ�䷶Χ", cbo(1).Text)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "���������ʱ�䷶Χ", cbo(5).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "������ѯ����", chk(2).Value)

    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�������ʱ�䷶Χ", cbo(2).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "������췶Χ", cbo(3).Text)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ�ˢ�¼��", Val(txt(1).Text))
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡָ����", chk(1).Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡ���뵥", chk(3).Value)
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9CISWork", "��ҩ����", chkPay.Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9CISWork", "�������", chkTime.Value
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "��ʾ����ҩ�����", chkҩ��.Value
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "��ʾ����ҩ����", chkҩ��.Value
    
    'ҩƷ��λ
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "ҩƷ��λ", IIf(optҩƷ��λ(0).Value, 0, 1)
    
    'ȱʡҩ��
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "����ȱʡ��ҩ��", cbo�ų�ҩ.ItemData(cbo�ų�ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "����ȱʡ��ҩ��", cbo����ҩ.ItemData(cbo����ҩ.ListIndex)
    
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "סԺȱʡ��ҩ��", cboס��ҩ.ItemData(cboס��ҩ.ListIndex)
    
    '�շ����
    For i = lst�շ����.ListCount - 1 To 0 Step -1
        If lst�շ����.Selected(i) Then strPar = strPar & "'" & Chr(lst�շ����.ItemData(i)) & "',"
    Next
    If strPar <> "" Then strPar = Left(strPar, Len(strPar) - 1)
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "�շ����", strPar
    
    '�Ƿ��������ִ�м�¼
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "����ִ�м�¼", chkActLog.Value

    '�Ƿ��������δ�շѲ��˵���Ŀ
    SaveSetting "ZLSOFT", "����ģ��\zl9CISWork", "δ�շ����", chkFinish.Value
    
    Call zlDatabase.SetPara(130, cbo(4).Text, glngSys)
    Call zlDatabase.SetPara(131, chk(0).Value, glngSys)
    
    mblnOK = True

    
    Unload Me
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub

Private Sub lst�շ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optҩƷ��λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tbs.Tab = 1
        cbo����ҩ.SetFocus
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub


