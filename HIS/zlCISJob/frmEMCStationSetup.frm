VERSION 5.00
Begin VB.Form frmEMCStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   Icon            =   "frmEMCStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk�ϴ���� 
      Caption         =   "��������ȡ�����ϴ����"
      Height          =   240
      Left            =   4485
      TabIndex        =   51
      Top             =   1270
      Width           =   2300
   End
   Begin VB.CheckBox chkȱʡҩ�� 
      Caption         =   "�´�ҽ��ʱǿ��ȱʡҩ��"
      Height          =   240
      Left            =   4485
      TabIndex        =   50
      Top             =   1595
      Width           =   2580
   End
   Begin VB.Frame fra�������Ƶ���ӡ 
      Caption         =   "ҽ�����ͺ�,���Ƶ���"
      Height          =   1875
      Index           =   0
      Left            =   7320
      TabIndex        =   47
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton opt�������Ƶ���ӡ 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         Width           =   1080
      End
      Begin VB.OptionButton opt�������Ƶ���ӡ 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   870
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton opt�������Ƶ���ӡ 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   420
         Width           =   1560
      End
   End
   Begin VB.Frame fra����ָ������ӡ 
      Caption         =   "ҽ�����ͺ�,ָ����"
      Height          =   1635
      Left            =   7320
      TabIndex        =   43
      Top             =   2280
      Width           =   2295
      Begin VB.OptionButton opt����ָ������ӡ 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   1200
      End
      Begin VB.OptionButton opt����ָ������ӡ 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   810
         Width           =   1560
      End
      Begin VB.OptionButton opt����ָ������ӡ 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   420
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.CheckBox chkYYBR 
      Caption         =   "�����б�����ʾԤԼ����"
      Height          =   240
      Left            =   150
      TabIndex        =   42
      Top             =   2640
      Width           =   2340
   End
   Begin VB.CheckBox chkCanPay 
      Caption         =   "���֧������ʹ��Ԥ����"
      Height          =   250
      Left            =   4485
      TabIndex        =   41
      Top             =   600
      Width           =   2310
   End
   Begin VB.CheckBox chkAutoClose 
      Caption         =   "ҽ�����ͺ��Զ��رմ���"
      Height          =   195
      Left            =   4485
      TabIndex        =   40
      Top             =   1920
      Width           =   2745
   End
   Begin VB.CheckBox chkAutoFinish 
      Caption         =   "���ﲡ��ʱ�Զ�����һ�����˴���Ϊ��ɾ���������"
      Height          =   195
      Left            =   150
      TabIndex        =   37
      Top             =   3600
      Width           =   4665
   End
   Begin VB.Frame fraEPR 
      Caption         =   "��������"
      Height          =   1410
      Left            =   135
      TabIndex        =   24
      Top             =   4000
      Width           =   9480
      Begin VB.CheckBox chkΣ��ֵ 
         Caption         =   "Σ��ֵ��������"
         Height          =   240
         Left            =   7725
         TabIndex        =   52
         Top             =   277
         Width           =   1620
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ��Ӧ"
         Height          =   195
         Index           =   6
         Left            =   7725
         TabIndex        =   49
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   5
         Left            =   6630
         TabIndex        =   48
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   4
         Left            =   5535
         TabIndex        =   36
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "����������ʾ"
         Height          =   195
         Left            =   4350
         TabIndex        =   39
         Top             =   300
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   5850
         TabIndex        =   38
         Top             =   240
         Width           =   1365
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ⱦ��"
         Height          =   195
         Index           =   3
         Left            =   4605
         TabIndex        =   35
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�������"
         Height          =   195
         Index           =   2
         Left            =   3500
         TabIndex        =   34
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "ҽ������"
         Height          =   195
         Index           =   1
         Left            =   2330
         TabIndex        =   33
         Top             =   975
         Width           =   1035
      End
      Begin VB.Frame fraNotifyEPRDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   690
         TabIndex        =   30
         Top             =   840
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   675
         TabIndex        =   27
         Top             =   470
         Width           =   300
      End
      Begin VB.TextBox txtNotifyEPR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   690
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "10"
         Top             =   270
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Σ��ֵ"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkNotifyEPR 
         Caption         =   "ÿ    �����Զ�ˢ�����������е�����"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   300
         Width           =   3450
      End
      Begin VB.TextBox txtNotifyEPRDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   705
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "1"
         Top             =   630
         Width           =   300
      End
      Begin VB.Label lblNotifyEPRDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ���ڲ�������Ϣ��ʾ����������"
         Height          =   180
         Left            =   480
         TabIndex        =   32
         Top             =   660
         Width           =   3060
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   465
         TabIndex        =   29
         Top             =   975
         Width           =   810
      End
   End
   Begin VB.Frame fraReceive 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   105
      TabIndex        =   20
      Top             =   3200
      Width           =   4560
      Begin VB.OptionButton optAdd 
         Caption         =   "����ҽ��"
         Enabled         =   0   'False
         Height          =   260
         Index           =   0
         Left            =   1290
         TabIndex        =   23
         Top             =   60
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.CheckBox chkAutoAdd 
         Caption         =   "���ﲡ�˺�"
         Height          =   195
         Left            =   45
         TabIndex        =   22
         Top             =   90
         Width           =   2055
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   260
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         Top             =   60
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdPBPSet 
      Caption         =   "֧��Ʊ�ݴ�ӡ����"
      Height          =   300
      Left            =   4485
      TabIndex        =   19
      Top             =   210
      Width           =   1620
   End
   Begin VB.CheckBox chkStaKB 
      Caption         =   "������Ļ����"
      Height          =   250
      Left            =   4485
      TabIndex        =   18
      Top             =   935
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   1410
      TabIndex        =   17
      Top             =   2490
      Width           =   465
   End
   Begin VB.TextBox txtQueuePatis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   1365
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "3"
      ToolTipText     =   "��ʾ����ҽ������ܺ��ж��ٸ�����������,�����󣬾Ͳ����ٴκ���;�˲�����Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪҽ������������Ч"
      Top             =   2280
      Width           =   465
   End
   Begin VB.CheckBox chk�Զ����� 
      Caption         =   "���ҵ����ﲡ��֮���Զ�����"
      Height          =   255
      Left            =   150
      TabIndex        =   10
      Top             =   2950
      Width           =   3135
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   135
      TabIndex        =   11
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Caption         =   " ������� "
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4155
      Begin VB.CommandButton cmdYS 
         Caption         =   "��"
         Height          =   255
         Left            =   3645
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   1635
         Width           =   255
      End
      Begin VB.TextBox txt����ҽ�� 
         Height          =   300
         Left            =   1020
         TabIndex        =   8
         Top             =   1605
         Width           =   2910
      End
      Begin VB.ComboBox cbo���� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2910
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   255
         Left            =   3645
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   690
         Width           =   255
      End
      Begin VB.ComboBox cbo��Χ 
         ForeColor       =   &H80000012&
         Height          =   300
         ItemData        =   "frmEMCStationSetup.frx":000C
         Left            =   1020
         List            =   "frmEMCStationSetup.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "����Ĳ��˷�Χ"
         Top             =   1005
         Width           =   2910
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   3
         Top             =   660
         Width           =   2910
      End
      Begin VB.Label lblEditDept 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   255
         TabIndex        =   0
         Top             =   360
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   45
         X2              =   4090
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label lblҽ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ��"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   255
         TabIndex        =   2
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl��Χ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ﷶΧ"
         Height          =   180
         Left            =   225
         TabIndex        =   5
         Top             =   1065
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   45
         X2              =   4090
         Y1              =   1460
         Y2              =   1460
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8565
      TabIndex        =   13
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   12
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Label lblQueuePatis 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������ܺ���      ��"
      Height          =   180
      Left            =   135
      TabIndex        =   15
      ToolTipText     =   "��ʾ����ҽ������ܺ��ж��ٸ�����������,�����󣬾Ͳ����ٴκ���;�˲�����Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪҽ������������Ч"
      Top             =   2300
      Width           =   1980
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   -240
      X2              =   10455
      Y1              =   5535
      Y2              =   5535
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   10320
      Y1              =   5565
      Y2              =   5565
   End
End
Attribute VB_Name = "frmEMCStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Private mstrLike As String
Private mbln���֧�� As Boolean

Private Enum Enum_chkWarn
    chkDΣ��ֵ = 0
    chkDҽ������ = 1
    chkD������� = 2
    chkD��Ⱦ�� = 3
    chkD��Ѫ��� = 4
    chkD��Ѫ��� = 5
    chkD��Ѫ��Ӧ = 6
End Enum

Private Sub chkAutoAdd_Click()
    If chkAutoAdd.Value = 1 Then
        optAdd(0).Enabled = True
        optAdd(1).Enabled = True
    Else
        optAdd(0).Enabled = False
        optAdd(1).Enabled = False
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call ZLCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdOK_Click()
    Dim blnHavePara As Boolean  '�Ƿ��в�������Ȩ��
    Dim i As Integer
    Dim strTmp As String
    
    
    If txt����.Text = "" Then
        MsgBox "������ҽ�������ҡ�", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If txt����ҽ��.Text = "" Then
        MsgBox "�����ҽ����", vbInformation, gstrSysName
        txt����ҽ��.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex < 0 Then
        MsgBox "������ұ���ѡ��,����", vbInformation + vbOKOnly, gstrSysName
        cbo����.SetFocus
        Exit Sub
    End If
    
    If chkNotifyEPR.Value = 1 And Val(txtNotifyEPR.Text) = 0 Then
        If txtNotifyEPR.Text = "" Then
            MsgBox "��������Ϣ���ѵ��Զ�ˢ�¼����", vbInformation, gstrSysName
        Else
            MsgBox "��Ϣ���ѵ��Զ�ˢ�¼������ӦΪ1���ӡ�", vbInformation, gstrSysName
        End If
        txtNotifyEPR.SetFocus: Exit Sub
    End If
    
    If Val(txtNotifyEPRDay.Text) = 0 Then
        If txtNotifyEPRDay.Text = "" Then
            MsgBox "������Ҫ������Ϣ�����������", vbInformation, gstrSysName
        Else
            MsgBox "Ҫ���ѵ���Ϣ�����������ӦΪ1�졣", vbInformation, gstrSysName
        End If
        txtNotifyEPRDay.SetFocus: Exit Sub
    End If
        
    blnHavePara = InStr(1, ";" & mstrPrivs & ";", ";��������;") > 0
    
    Call zlDatabase.SetPara("��������", Me.txt����.Text, glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("���ﷶΧ", Me.cbo��Χ.ItemData(Me.cbo��Χ.ListIndex), glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("����ҽ��", Me.txt����ҽ��.Text, glngSys, p����ҽ��վ, blnHavePara)
    
    '���˺�:Ӧ�����Ŷӽкŵĺ����˴�:��Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪ���������ŶӺ��ж���=2ʱ��Ч
    If txtQueuePatis.Enabled Then
        Call zlDatabase.SetPara("ҽ����������", Val(Me.txtQueuePatis.Text), glngSys, p����ҽ��վ, blnHavePara)
    End If
    '�������
    Call zlDatabase.SetPara("�������", cbo����.ItemData(cbo����.ListIndex), glngSys, p����ҽ��վ, blnHavePara)
    
    '������ɺ�ر�ҽ������
    Call zlDatabase.SetPara("������ɺ�ر�ҽ������", chkAutoClose.Value, glngSys, p����ҽ���´�, blnHavePara)
    
    '�ҵ����˺��Զ�����
    Call zlDatabase.SetPara("�ҵ����˺��Զ�����", chk�Զ�����.Value, glngSys, p����ҽ��վ, blnHavePara)
    
    '������Զ�����
    If optAdd(1).Value And optAdd(1).Enabled Then
        Call zlDatabase.SetPara("������Զ�����", 2, glngSys, p����ҽ��վ, blnHavePara)
    Else
        Call zlDatabase.SetPara("������Զ�����", chkAutoAdd.Value, glngSys, p����ҽ��վ, blnHavePara)
    End If

    '������Ļ����
    Call zlDatabase.SetPara("������Ļ����", chkStaKB.Value, glngSys, p����ҽ��վ, blnHavePara)
    
    Call zlDatabase.SetPara("�Զ�ˢ�²������ļ��", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("�Զ�ˢ�²�����������", Val(txtNotifyEPRDay.Text), glngSys, p����ҽ��վ, blnHavePara)
    strTmp = ""
    For i = chkDΣ��ֵ To chkD��Ѫ��Ӧ
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zlDatabase.SetPara("�Զ�ˢ������", strTmp, glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("����ʱ�Զ�������ɾ���", chkAutoFinish.Value, glngSys, p����ҽ��վ, blnHavePara)
    Call zlDatabase.SetPara("����������ʾ", chkSound.Value, glngSys, p����ҽ��վ, blnHavePara)
    
    Call zlDatabase.SetPara("���֧������ʹ��Ԥ����", chkCanPay.Value, glngSys, p����ҽ���´�, blnHavePara)
    
    Call zlDatabase.SetPara("��ʾԤԼ����", chkYYBR.Value, glngSys, p����ҽ��վ, blnHavePara)
    
    Call zlDatabase.SetPara("����ҽ���´�ǿ��ȱʡҩ��", chkȱʡҩ��.Value, glngSys, p����ҽ���´�, blnHavePara)
    
    Call zlDatabase.SetPara("����Σ��ֵ��������", chkΣ��ֵ.Value, glngSys, p����ҽ��վ, blnHavePara)
    
    Call zlDatabase.SetPara("�ϴ���ϰ�������ȡ", chk�ϴ����.Value, glngSys, p����ҽ���´�, blnHavePara)
    
    'ҽ��վ�Ƿ��ӡ���Ƶ���
    Call zlDatabase.SetPara("���﷢�͵��ݴ�ӡ", IIf(opt�������Ƶ���ӡ(0).Value = True, 0, IIf(opt�������Ƶ���ӡ(1).Value = True, 1, 2)), glngSys, p����ҽ���´�, blnHavePara)
    'ҽ��վ�Ƿ��ӡָ����
    Call zlDatabase.SetPara("ָ������ӡ��ʽ", IIf(opt����ָ������ӡ(0).Value = True, 0, IIf(opt����ָ������ӡ(1).Value = True, 1, 2)), glngSys, p����ҽ���´�, blnHavePara)

    gblnOK = True
    Unload Me
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub


Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 0)
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlcontrol.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlcontrol.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cmdPBPSet_Click()
    If InitObjPublicExpense Then
        Call gobjPublicExpense.zlCliniqueRoomPayPrintSet(Me)
    End If
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, blnCancel As Boolean
    
    
    If txt����.Tag <> txt���� Then Exit Sub '��txt���ҵ�Validate�¼�����
    On Error GoTo errH
    If gbln�ҺŰ��� Then
        strSql = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
            " From �������� A, �����������ÿ��� B, ������Ա C, �ϻ���Ա�� D,�ٴ����� F" & vbNewLine & _
            " Where a.Id = b.����id And b.����id = c.����id And c.��Աid = d.��Աid" & vbNewLine & _
            "       And d.�û��� = User And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null) And C.����ID = F.����ID And F.�������� = '20'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Else
        Set rsTmp = GetRs�ҺŰ��������б�(1, "", 1, 0, "", p����ҽ��վ)
    End If
    
    '���û�в��ҵ����ݣ����ȡ�����е��������ҹ�ѡ��
    If rsTmp.EOF Then
        strSql = "Select a.Id, a.����, a.���� From �������� A Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    End If
    
    vRect = zlcontrol.GetControlRect(txt����.hwnd)
    Set rsTmp = zlDatabase.ShowRecSelect(Me, rsTmp, 0, "��������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt����.Height, blnCancel, False, True)
        
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            txt����.Tag = rsTmp("����"): txt���� = txt����.Tag
            If cbo��Χ.Enabled And cbo��Χ.Visible Then cbo��Χ.SetFocus
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdYS_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnCanle As Boolean
    If txt����ҽ��.Tag <> txt����ҽ�� Then Exit Sub '��txtҽ����Validate�¼�����
    On Error GoTo errH
    strSql = "Select Distinct A.��� as ID,A.���� as ����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D,�ٴ����� E" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And B.����ID=D.����ID" & _
        " And C.��Ա����||''='ҽ��' And D.������� IN(1,3) And D.��������||''='�ٴ�'" & _
        " And B.����ID In (Select ����ID From ������Ա Where ��ԱID=[1])" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And B.����ID = E.����ID And E.�������� = '20'" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "����ҽ��", False, "", "", False, False, False, 0, 0, txt����ҽ��.Height, blnCanle, False, True, UserInfo.ID)
    If blnCanle Then Exit Sub
    If Not rsTmp Is Nothing Then txt����ҽ��.Tag = rsTmp("����"): txt����ҽ�� = txt����ҽ��.Tag: Me.cmdOK.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strPar As String
    Dim blnSetup As Boolean
    Dim i As Long
    Dim intType As Integer
    Dim strNotify As String
    Dim str���� As String
    
    blnSetup = InStr(1, ";" & mstrPrivs & ";", ";��������;") > 0
    gblnOK = False
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    On Error Resume Next
    str���� = zlDatabase.GetPara("��������", glngSys, p����ҽ��վ, "", Array(lbl����, txt����, cmdSel), blnSetup)
    On Error GoTo 0
    
    On Error GoTo errH
    mbln���֧�� = Val(zlDatabase.GetPara("����ҽ�����ͺ��������֧��", glngSys, p����ҽ���´�)) = 1
    cmdPBPSet.Enabled = mbln���֧��
    '��ȡ����ȱʡ���ҷ�Χ
    strPar = zlDatabase.GetPara("�������", glngSys, p����ҽ��վ, "", Array(lblEditDept, cbo����), blnSetup)
    
    strSql = "Select Distinct B.ID,B.����,B.����,A.ȱʡ" & _
        " From ������Ա A,���ű� B,��������˵�� C,�ٴ����� D" & _
        " Where A.����ID=B.ID And B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null) And D.�������� = '20' And D.����ID = A.����ID" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And A.��ԱID=[1]" & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        cbo����.AddItem rsTmp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
        If rsTmp!ID = Val(strPar) Then
            cbo����.ListIndex = cbo����.NewIndex
        ElseIf Nvl(rsTmp!ȱʡ, 0) = 1 And cbo����.ListIndex = -1 Then
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTmp.MoveNext
    Next
    Me.cbo��Χ.ListIndex = Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2", Array(lbl��Χ, cbo��Χ), blnSetup)) - 1
    
    strSql = "Select 1 From �������� E where e.����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str����)
    If Not rsTmp.EOF Then
        txt����.Text = str����
        txt����.Tag = str����
    End If
    
    '����ѡ�������ҽ������Ĳ��˽��о���
    If InStr(mstrPrivs, "���ﲡ��") > 0 And InStr(mstrPrivs, "�������ý���ҽ��") > 0 Then
        '����ѡ�񱾿����µ�ҽ��
        cmdYS.Enabled = True
        txt����ҽ��.Enabled = True
    Else
        cmdYS.Enabled = False
        txt����ҽ��.Enabled = False
    End If
    txt����ҽ��.Tag = zlDatabase.GetPara("����ҽ��", glngSys, p����ҽ��վ, UserInfo.����, Array(lblҽ��, txt����ҽ��, cmdYS), blnSetup)
    txt����ҽ��.Text = txt����ҽ��.Tag
    
   
    '���˺�:Ӧ�����Ŷӽкŵĺ����˴�:��Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪ���������ŶӺ���վ��=1ʱ��Ч
    txtQueuePatis.Text = Val(zlDatabase.GetPara("ҽ����������", glngSys, p����ҽ��վ, 3, Array(lblQueuePatis, txtQueuePatis), blnSetup))
    If txtQueuePatis.Enabled Then
        If Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, p����������)) = 1 And Val(zlDatabase.GetPara("�ŶӺ���վ��", glngSys, p����������, "0")) = 1 Then
            txtQueuePatis.Enabled = True
        Else
            txtQueuePatis.Enabled = False
        End If
    End If
    
    '������ɺ�ر�ҽ������
    chkAutoClose.Value = Val(zlDatabase.GetPara("������ɺ�ر�ҽ������", glngSys, p����ҽ���´�, , Array(chkAutoClose), blnSetup))
    
    '�ҵ����˺��Զ�����
    chk�Զ�����.Value = Val(zlDatabase.GetPara("�ҵ����˺��Զ�����", glngSys, p����ҽ��վ, , Array(chk�Զ�����), blnSetup))
    
    '���֧������ʹ��Ԥ����
    chkCanPay.Value = Val(zlDatabase.GetPara("���֧������ʹ��Ԥ����", glngSys, p����ҽ���´�, , Array(chkCanPay), blnSetup))
    
    '������Զ�����
    strPar = Val(zlDatabase.GetPara("������Զ�����", glngSys, p����ҽ��վ, , Array(chkAutoAdd, optAdd(0), optAdd(1)), blnSetup))
    
    If strPar = 2 Then
        chkAutoAdd.Value = 1
        optAdd(1).Value = True
    Else
        chkAutoAdd.Value = strPar
    End If
    
    '������Ļ����
    chkStaKB.Value = Val(zlDatabase.GetPara("������Ļ����", glngSys, p����ҽ��վ, , Array(chkStaKB), blnSetup))
    
    '��Ϣ����ˢ��
    strPar = zlDatabase.GetPara("�Զ�ˢ�²������ļ��", glngSys, p����ҽ��վ, , Array(chkNotifyEPR), blnSetup, intType)
    If Val(strPar) > 0 Then
        chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
    End If
 
    If (intType = 3 Or intType = 15) And Not blnSetup Then
        txtNotifyEPR.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("�Զ�ˢ�²�����������", glngSys, p����ҽ��վ, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), blnSetup)
    txtNotifyEPRDay.Text = IIf(0 = Val(strPar), 1, Val(strPar))
        
    strNotify = zlDatabase.GetPara("�Զ�ˢ������", glngSys, p����ҽ��վ, , Array(chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), lblArea), blnSetup)
    chkWarn(chkDΣ��ֵ).Value = Val(Mid(strNotify, 1, 1))
    chkWarn(chkDҽ������).Value = Val(Mid(strNotify, 2, 1))
    chkWarn(chkD�������).Value = Val(Mid(strNotify, 3, 1))
    chkWarn(chkD��Ⱦ��).Value = Val(Mid(strNotify, 4, 1))
    chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 5, 1))
    chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 6, 1))
    chkWarn(chkD��Ѫ��Ӧ).Value = Val(Mid(strNotify, 7, 1))
    chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
    chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
    chkWarn(chkD��Ѫ��Ӧ).Visible = gblnѪ��ϵͳ
    If InStr(mstrPrivs, "��������") = 0 Then
        chkWarn(chkDΣ��ֵ).Enabled = False
        chkWarn(chkDҽ������).Enabled = False
        chkWarn(chkD�������).Enabled = False
        chkWarn(chkD��Ⱦ��).Enabled = False
        chkWarn(chkD��Ѫ���).Enabled = False
        chkWarn(chkD��Ѫ���).Enabled = False
        chkWarn(chkD��Ѫ��Ӧ).Enabled = False
    End If
    chkAutoFinish.Value = Val(zlDatabase.GetPara("����ʱ�Զ�������ɾ���", glngSys, p����ҽ��վ, , Array(chkAutoFinish), blnSetup))
    chkSound.Value = Val(zlDatabase.GetPara("����������ʾ", glngSys, p����ҽ��վ, , Array(chkSound, cmdSoundSet), blnSetup))
    
    '��ʾԤԼ����
    chkYYBR.Value = Val(zlDatabase.GetPara("��ʾԤԼ����", glngSys, p����ҽ��վ, "1", Array(chkYYBR), blnSetup))
    
    '����ҽ���´�ǿ��ȱʡҩ��
    chkȱʡҩ��.Value = Val(zlDatabase.GetPara("����ҽ���´�ǿ��ȱʡҩ��", glngSys, p����ҽ���´�, "1", Array(chkȱʡҩ��), blnSetup))
    
    '����Σ��ֵ��������
    chkΣ��ֵ.Value = Val(zlDatabase.GetPara("����Σ��ֵ��������", glngSys, p����ҽ��վ, "1", Array(chkΣ��ֵ), blnSetup))
    
    chk�ϴ����.Value = Val(zlDatabase.GetPara("�ϴ���ϰ�������ȡ", glngSys, p����ҽ���´�, , Array(chk�ϴ����), blnSetup))
    
    'ҽ��վ�Ƿ��ӡ���Ƶ���
    strPar = Val(zlDatabase.GetPara("���﷢�͵��ݴ�ӡ", glngSys, p����ҽ���´�, , Array(opt�������Ƶ���ӡ(0), opt�������Ƶ���ӡ(1), opt�������Ƶ���ӡ(2)), blnSetup))
    opt�������Ƶ���ӡ(Val(strPar)) = True
    'ҽ��վ�Ƿ��ӡָ����
    strPar = Val(zlDatabase.GetPara("ָ������ӡ��ʽ", glngSys, p����ҽ���´�, , Array(opt����ָ������ӡ(0), opt����ָ������ӡ(1), opt����ָ������ӡ(2)), blnSetup))
    opt����ָ������ӡ(Val(strPar)) = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
End Sub

Private Sub txt����ҽ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, blnCancel As Boolean

    If txt����ҽ��.Tag = txt����ҽ�� Then Exit Sub

    strSql = "Select Distinct A.��� as ID,A.���� as ����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D,�ٴ����� E" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And B.����ID=D.����ID" & _
        " And C.��Ա����||''='ҽ��' And D.������� IN(1,3) And D.��������||''='�ٴ�' And B.����ID = E.����ID And E.�������� = '20'" & _
        " And B.����ID In(Select ����ID From ������Ա Where ��ԱID=[1])" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (Upper(A.���) Like [2] Or Upper(A.����) Like [3] Or A.���� Like [3])" & _
        " Order by A.����"
        
    vRect = zlcontrol.GetControlRect(txt����ҽ��.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "����ҽ��", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt����ҽ��.Height, blnCancel, False, True, UserInfo.ID, UCase(txt����ҽ��.Text) & "%", mstrLike & UCase(txt����ҽ��.Text) & "%")
    If Not rsTmp Is Nothing Then
        txt����ҽ��.Tag = rsTmp("����")
        txt����ҽ�� = txt����ҽ��.Tag
    Else
        txt����ҽ��.Tag = ""
        txt����ҽ�� = ""
        Cancel = blnCancel
    End If
End Sub

Private Sub txt����ҽ��_GotFocus()
    Call zlcontrol.TxtSelAll(txt����ҽ��)
End Sub

Private Sub txt����ҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt����ҽ�� = "" Then txt����ҽ��.Tag = "1"
        ZLCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlcontrol.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt���� = "" Then txt����.Tag = "1"
        ZLCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If txt����.Tag = txt���� Then Exit Sub
    On Error GoTo errH
    If gbln�ҺŰ��� Then
        strSql = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
            " From �������� A, �����������ÿ��� B, ������Ա C, �ϻ���Ա�� D,�ٴ����� F" & vbNewLine & _
            " Where a.Id = b.����id And b.����id = c.����id And c.��Աid = d.��Աid And C.����ID = F.����ID And F.�������� = '20'" & vbNewLine & _
            " And (Upper(a.����) Like [1] Or Upper(a.����) Like [2] Or Upper(a.����) Like [2])" & _
            "       And d.�û��� = User And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)"
            
            '���û�в��ҵ����ݣ����ȡ�����е��������ҹ�ѡ��
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(txt����.Text) & "%", mstrLike & UCase(txt����.Text) & "%")
    Else
        Set rsTmp = GetRs�ҺŰ��������б�(1, UCase(txt����.Text), 1, 0, "", p����ҽ��վ)
    End If
        

    If rsTmp.EOF Then
        strSql = "Select a.Id, a.����, a.���� From �������� A Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null)" & _
            " And (Upper(a.����) Like [1] Or Upper(a.����) Like [2] Or Upper(a.����) Like [2])"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(txt����.Text) & "%", mstrLike & UCase(txt����.Text) & "%")
    End If
        
    vRect = zlcontrol.GetControlRect(txt����.hwnd)
    Set rsTmp = zlDatabase.ShowRecSelect(Me, rsTmp, 0, "��������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt����.Height, blnCancel, False, True)
    If Not rsTmp Is Nothing Then
        txt����.Tag = rsTmp("����")
        txt���� = txt����.Tag
    Else
        txt����.Tag = ""
        txt���� = ""
        Cancel = blnCancel
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




