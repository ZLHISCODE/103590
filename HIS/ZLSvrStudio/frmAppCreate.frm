VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ��ϵͳ��װ"
   ClientHeight    =   5085
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmAppCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   1
      Left            =   1305
      TabIndex        =   9
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.ComboBox cmbEnjoy 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1440
         Width           =   2610
      End
      Begin VB.Frame fraOwner 
         Caption         =   "�½�������"
         Height          =   1755
         Left            =   585
         TabIndex        =   29
         Top             =   1935
         Width           =   4530
         Begin VB.CheckBox chkDBA 
            Caption         =   "����DBA��ɫ"
            Height          =   255
            Left            =   3030
            TabIndex        =   50
            Top             =   1215
            Width           =   1320
         End
         Begin VB.TextBox txtOwnerUsr 
            Height          =   300
            Left            =   810
            TabIndex        =   30
            Top             =   360
            Width           =   1890
         End
         Begin VB.TextBox txtOwnerPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   810
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   780
            Width           =   1890
         End
         Begin VB.TextBox txtOwnerLab 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   810
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   32
            Top             =   1200
            Width           =   1890
         End
         Begin VB.Label lblDBA 
            Caption         =   "���Ը��ݹ���ϰ�߾����Ƿ�����DBA��ɫ"
            Height          =   660
            Left            =   3030
            TabIndex        =   49
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label lblNewUser 
            AutoSize        =   -1  'True
            Caption         =   "�û���"
            Height          =   180
            Left            =   210
            TabIndex        =   35
            Top             =   420
            Width           =   540
         End
         Begin VB.Label lblNewPwd 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   390
            TabIndex        =   34
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lblNewLab 
            AutoSize        =   -1  'True
            Caption         =   "��֤"
            Height          =   180
            Left            =   390
            TabIndex        =   33
            Top             =   1260
            Width           =   360
         End
      End
      Begin VB.CheckBox chkEnjoy 
         Caption         =   "ѡ���蹲��ϵͳ(&S)"
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   1500
         Width           =   1830
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "�ڶ��� ���ñ�ϵͳ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   225
         Width           =   2595
      End
      Begin VB.Label lblNote 
         Caption         =   "    ��ϵͳ���Ѿ���װ��������Ʒ����ѡ������ϵͳ����(������֪�������ߵ�����)��Ҳ�����½������ߣ������κβ�Ʒ����"
         Height          =   585
         Index           =   1
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   5250
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   0
      Left            =   1305
      TabIndex        =   4
      Top             =   -120
      Width           =   6075
      Begin VB.Frame fraSys 
         Height          =   1005
         Left            =   570
         TabIndex        =   46
         Top             =   2340
         Width           =   4545
         Begin VB.Label lblVersion 
            AutoSize        =   -1  'True
            Caption         =   "�汾�ţ�"
            Height          =   180
            Left            =   210
            TabIndex        =   48
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblSysName 
            AutoSize        =   -1  'True
            Caption         =   "ϵͳ����"
            Height          =   180
            Left            =   210
            TabIndex        =   47
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdSetupFile 
         Caption         =   "ѡ��(&S)��"
         Height          =   350
         Left            =   570
         TabIndex        =   5
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblSetupFile 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   570
         TabIndex        =   28
         Top             =   1650
         Width           =   4545
      End
      Begin VB.Label lbliniFile 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�ð�װ�����ļ�"
         Height          =   180
         Left            =   570
         TabIndex        =   27
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label lblNote 
         Caption         =   "    Ӧ��ϵͳ�İ�װ�����������ļ�����֮��صķ����������ű��ļ�������ȷָ����װ�����ļ���"
         Height          =   450
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   720
         Width           =   5250
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "��һ�� ָ����װ�����ļ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   2595
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   2
      Left            =   1305
      TabIndex        =   14
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraSpace 
         Height          =   2070
         Left            =   495
         TabIndex        =   55
         Top             =   1770
         Width           =   5055
         Begin VB.CheckBox chkLogin 
            Caption         =   "������ռ��¼��־"
            Height          =   270
            Index           =   0
            Left            =   3000
            TabIndex        =   61
            ToolTipText     =   "������ռ��Ƿ������־��Ĭ�ϲ�������־"
            Top             =   1155
            Width           =   1920
         End
         Begin VB.ComboBox cboSpaceExtentType 
            Height          =   300
            Index           =   0
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   60
            ToolTipText     =   "AUTOALLOCATE �� UNIFORM Size nM"
            Top             =   1620
            Width           =   1815
         End
         Begin VB.CheckBox chkSpaceExtd 
            Caption         =   "�Զ���չ"
            Height          =   270
            Index           =   0
            Left            =   1815
            TabIndex        =   59
            ToolTipText     =   "AUTOEXTEND ON Next (��ռ��С/10)M"
            Top             =   1155
            Width           =   1065
         End
         Begin VB.TextBox txtSpaceExtentSize 
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   58
            Text            =   "1"
            Top             =   1620
            Width           =   255
         End
         Begin VB.TextBox txtSpaceFile 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   705
            TabIndex        =   57
            Top             =   675
            Width           =   4005
         End
         Begin VB.TextBox txtSpaceSize 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   705
            MaxLength       =   10
            TabIndex        =   56
            Top             =   1125
            Width           =   750
         End
         Begin VB.Label lblSpaceExtend 
            AutoSize        =   -1  'True
            Caption         =   "���ߴ�"
            Height          =   180
            Left            =   105
            TabIndex        =   67
            Top             =   1680
            Width           =   540
         End
         Begin VB.Label lblTBS 
            Caption         =   "M"
            Height          =   255
            Left            =   2970
            TabIndex        =   66
            Top             =   1695
            Width           =   135
         End
         Begin VB.Label txtSpaceName 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   0
            Left            =   705
            TabIndex        =   65
            Top             =   225
            Width           =   2145
         End
         Begin VB.Label lblSpaceName 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   285
            TabIndex        =   64
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblSpaceFile 
            AutoSize        =   -1  'True
            Caption         =   "�ļ�"
            Height          =   180
            Left            =   285
            TabIndex        =   63
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lblSpaceSize 
            AutoSize        =   -1  'True
            Caption         =   "��С          M"
            Height          =   180
            Left            =   285
            TabIndex        =   62
            Top             =   1185
            Width           =   1350
         End
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   465
         Width           =   5800
      End
      Begin MSComctlLib.TabStrip tbsSpace 
         Height          =   2520
         Left            =   405
         TabIndex        =   54
         ToolTipText     =   "�����ı�ռ�����Ϊ���ع����ռ�(��ASSM)"
         Top             =   1440
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   4445
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.Label lblNote 
         Caption         =   "    ϵͳ��Ҫ�������±�ռ䣬Ϊ���ʹ���I/O��ͻ����ø��ݷ������������������ռ�ֱ����ڲ�ͬ�Ĵ����ϡ�"
         Height          =   405
         Index           =   2
         Left            =   225
         TabIndex        =   17
         Top             =   675
         Width           =   5610
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "������ ���ݴ洢�ռ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   225
         Width           =   2145
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5775
      TabIndex        =   0
      Top             =   4260
      Width           =   1100
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4704
      Left            =   0
      ScaleHeight     =   4650
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.Image imgSetup 
         Height          =   3315
         Left            =   60
         Picture         =   "frmAppCreate.frx":058A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   1545
      TabIndex        =   44
      Top             =   4260
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   2490
      TabIndex        =   43
      Top             =   4860
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ��(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4695
      TabIndex        =   3
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3585
      TabIndex        =   1
      Top             =   4260
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   42
      Top             =   4704
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppCreate.frx":5B70
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9340
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "13:59"
            Key             =   "STANUM"
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
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   3
      Left            =   1305
      TabIndex        =   23
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.CheckBox chkSelData 
         Caption         =   "ѡ��װ����(&S)"
         Height          =   210
         Left            =   765
         TabIndex        =   40
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.Frame fraSelData 
         Height          =   3090
         Left            =   585
         TabIndex        =   37
         Top             =   1140
         Width           =   4845
         Begin VB.PictureBox picXp 
            BorderStyle     =   0  'None
            Height          =   2760
            Left            =   75
            ScaleHeight     =   2760
            ScaleWidth      =   1845
            TabIndex        =   51
            Top             =   255
            Width           =   1845
            Begin VB.OptionButton optData 
               Caption         =   "��ѡ���ݷ���0"
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   52
               Top             =   15
               Width           =   1665
            End
            Begin VB.Label lblNoData 
               Caption         =   "�������鲻��ϸ��Ϊ��ѡ��������"
               Height          =   2730
               Left            =   60
               TabIndex        =   53
               Top             =   0
               Visible         =   0   'False
               Width           =   4410
            End
         End
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "ȫ��"
            Height          =   315
            Left            =   3090
            TabIndex        =   39
            Top             =   210
            Width           =   1100
         End
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "ȫѡ"
            Height          =   315
            Left            =   1965
            TabIndex        =   38
            Top             =   210
            Width           =   1100
         End
         Begin VB.ListBox lstData 
            Height          =   1950
            Index           =   0
            ItemData        =   "frmAppCreate.frx":6402
            Left            =   1950
            List            =   "frmAppCreate.frx":6404
            Style           =   1  'Checkbox
            TabIndex        =   41
            Top             =   525
            Visible         =   0   'False
            Width           =   2670
         End
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblNote 
         Caption         =   "    Ϊ�ܸ���ʹ�ã�ϵͳ׼���˲���Ӧ�����ݣ����ݲ�ͬ��ʹ�����������ѡ��װ��ͬ�������顣"
         Height          =   405
         Index           =   3
         Left            =   225
         TabIndex        =   26
         Top             =   720
         Width           =   5250
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "���Ĳ� ��װ����ѡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   25
         Top             =   225
         Width           =   2145
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   4230
      Index           =   4
      Left            =   1305
      TabIndex        =   18
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblNextDo 
         AutoSize        =   -1  'True
         Caption         =   "    ���""���""��ʼ�Զ�װ��ϵͳ������""ȡ��""��ֹϵͳװ�أ���""��һ��""���µ���Ӧ��ϵͳװ�����á�"
         Height          =   360
         Left            =   225
         TabIndex        =   45
         Top             =   2025
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRegAudit 
         AutoSize        =   -1  'True
         Caption         =   "    ���ڻ����߱���ϵͳӦ����Ȩ����Ȼ���Լ���װ�أ����޷�����ʹ�á�"
         Height          =   360
         Left            =   225
         TabIndex        =   22
         Top             =   1335
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "���岽 ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   165
         TabIndex        =   21
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "    �Ѿ�����˶Ը�ϵͳװ�ص�ȫ�����á�"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   20
         Top             =   720
         Width           =   3420
      End
   End
End
Attribute VB_Name = "frmAppCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum ���ݶ�
    sec���� = 0             '���д�����ֶε����ݱ�
    sec�ֶ��� = 1           '������ֶ�
    sec�ֶ����� = 2         'Long �� Raw
    sec���� = 3             '�����������ֶ������ֶ�ֵ����|�ָ�������������ɶ���ֶ���ɣ��������г�
    sec�������� = 4         'Insert �� Update
    sec�ļ��� = 5           '���д�������ݵ��ļ���������·��
End Enum
Private mstrIniPath      As String                 '��װ�����ļ�Ŀ¼
Private intDefSysCode   As String                 'ϵͳ���
Private strDefSysName   As String                 'ϵͳ����
Private strDefVersion   As String                 '�汾��
Private strDefSpace   As String                   '��ռ䶨�崮
Private strDefUser      As String                 '�µ�ȱʡ�û���
Private strDefData      As String                 '�û���ѡ������

Private mstrExtSysCode  As String                  'Ҫ������չ����ϵͳ�ı��
Private mstrExtVersion  As String                  'Ҫ������չ����ϵͳ�İ汾
Private mstrTbsPath As String                        'ȱʡ��ռ�·�����ƣ�������ʷ��ռ����

Private objText As TextStream
Private mstrLogFile As String
Private mclsRunScript As clsRunScript  '�ű�����ִ����
Private mfrmUpSys As frmAppUpgradeNew
Private intStep As Integer

Private mbln���� As Boolean    '���ΰ�װ�Ƿ����������װ�װ
Private mlng���� As Long       '���׺�
Private mlst��׼ As ListItem   '�����Ҫ��װ�����ף������ṩ��׼�������ݵ�ϵͳ

Private mcnOwner As New ADODB.Connection
Private intCount As Integer, intItems As Integer
        
Private aryRow() As String
Private aryVal() As String


Private Sub cboSpaceExtentType_Click(Index As Integer)
    txtSpaceExtentSize(Index).Enabled = (cboSpaceExtentType(Index).ListIndex = 1)
    If txtSpaceExtentSize(Index).Enabled Then
        If MsgBox("������������á��Զ��������ߴ硱ѡ��Ƿ�ԭ��ΪĬ��ֵ��", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            cboSpaceExtentType(Index).ListIndex = 0
            txtSpaceExtentSize(Index).Enabled = (cboSpaceExtentType(Index).ListIndex = 1)
        End If
    End If
End Sub

Private Sub chkEnjoy_Click()
'�Ƿ���װ
    Dim blnEnjoy As Boolean
    blnEnjoy = chkEnjoy.value = 1
    cmbEnjoy.Enabled = blnEnjoy
    If blnEnjoy Then
        fraOwner.Caption = "������"
        txtOwnerUsr.Text = cmbEnjoy.Tag
        txtOwnerUsr.Enabled = False
        txtOwnerLab.Enabled = False
        chkDBA.Enabled = False
    Else
        fraOwner.Caption = "�½�������"
        txtOwnerUsr.Text = strDefUser
        txtOwnerUsr.Enabled = True
        txtOwnerLab.Enabled = True
        chkDBA.Enabled = True
        
        If fraSetup(1).Visible = False Then Exit Sub
        txtOwnerUsr.SetFocus
    End If
    '���ÿؼ�λ���Լ�״̬
    lblNewLab.Visible = Not blnEnjoy
    txtOwnerLab.Visible = Not blnEnjoy
    chkDBA.Visible = Not blnEnjoy
    lblDBA.Visible = Not blnEnjoy
    '���ÿؼ�λ��
    txtOwnerUsr.Left = IIf(blnEnjoy, 1200, 810)
    txtOwnerPwd.Left = txtOwnerUsr.Left
    lblNewUser.Left = txtOwnerUsr.Left - lblNewUser.Width - 60
    lblNewPwd.Left = txtOwnerUsr.Left - lblNewPwd.Width - 60
    txtOwnerUsr.Top = IIf(blnEnjoy, 540, 360)
    lblNewUser.Top = txtOwnerUsr.Top + (txtOwnerUsr.Height - lblNewUser.Height) / 2
    txtOwnerPwd.Top = txtOwnerUsr.Top + txtOwnerUsr.Height + IIf(blnEnjoy, 240, 120)
    lblNewPwd.Top = txtOwnerPwd.Top + (txtOwnerPwd.Height - lblNewPwd.Height) / 2
    If mstrExtSysCode = "" Then txtOwnerPwd.SetFocus
End Sub

Private Sub chkSelData_Click()
'�Ƿ�װ��ѡ������
    Dim i As Integer
    Dim blnEnable As Boolean
    If chkSelData.value = 0 Then
        blnEnable = False
    Else
        blnEnable = True
    End If
    
    fraSelData.Enabled = blnEnable
    For i = optData.LBound To optData.UBound
        optData(i).Enabled = blnEnable
    Next
    For i = lstData.LBound To lstData.UBound
        lstData(i).Enabled = blnEnable
    Next
    
    cmdSelectAll.Enabled = blnEnable
    cmdClearAll.Enabled = blnEnable
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("��װδ��ɣ����ȡ����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    '���������װʱ�Ѵ����û�(���˵ڶ���)����ɾ���û�
    If chkEnjoy.value = 0 And txtOwnerUsr.Tag = "�Ѵ����û�" Then
    
        On Error Resume Next
        gstrSQL = "drop user " & txtOwnerUsr.Text
        gcnOracle.Execute gstrSQL
    End If
    
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdSelectAll_Click()
    Dim lngIndex As Long, lngCount As Long
    
    For lngIndex = lstData.LBound To lstData.UBound
        With lstData(lngIndex)
            If .Visible = True Then
                For lngCount = 0 To .ListCount - 1
                    .Selected(lngCount) = True
                Next
                
                .Refresh
            End If
        End With
    Next
End Sub

Private Sub cmdClearAll_Click()
    Dim lngIndex As Long, lngCount As Long
    
    For lngIndex = lstData.LBound To lstData.UBound
        With lstData(lngIndex)
            If .Visible = True Then
                For lngCount = 0 To .ListCount - 1
                    .Selected(lngCount) = False
                Next
                
                .Refresh
            End If
        End With
    Next
End Sub

Private Sub cmdSetupFile_Click()
    With frmMDIMain.dlgMain
        .InitDir = App.Path
        .FileName = lblSetupFile.Caption
        .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
        .Filter = "(Ӧ�ð�װ�����ļ�)|zlSetup.ini"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblSetupFile.Caption = .FileName
        End If
    End With
    If ChkSetupFile(True) = False Then
        mbln���� = False
        lblSetupFile.Caption = ""
        cmdSetupFile.SetFocus
    End If

End Sub

Private Sub cmdNext_Click()
    Dim objfrmUpSys As frmAppUpgradeNew
    Dim strError As String
    
    SetPromptText ""
    If fraSetup(0).Visible Then
        '------------------------------------------------------------
        '��һ����
        '------------------------------------------------------------
        If Trim(lblSetupFile.Caption) = "" Then
            MsgBox "δ��ȷѡ���������װ�����ļ������ܼ�����", vbExclamation, gstrSysName
            cmdSetupFile.SetFocus
            Exit Sub
        End If
        
        '------------------------------
        fraSetup(0).Visible = False
        fraSetup(1).Visible = True
        cmdPrevious.Enabled = True
        If cmbEnjoy.ListCount > 0 Then cmbEnjoy.ListIndex = 0
    
    ElseIf fraSetup(1).Visible Then
        '------------------------------------------------------------
        '�ڶ�����
        '------------------------------------------------------------
        If chkEnjoy.value = 1 Then       '�����������ȷ����������
            Set mcnOwner = gobjRegister.GetConnection(gstrServer, Trim(txtOwnerUsr.Text), Trim(txtOwnerPwd.Text), True, MSODBC, "", False)
            If mcnOwner.State = adStateClosed Then
                MsgBox "������������󣬲��ܼ�����", vbExclamation, gstrSysName
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
            
            Call SetSQLTrace(gstrServer, Trim(txtOwnerUsr.Text), mcnOwner)
        Else
            '��������ϵͳ�������뽨���û�
            If Len(Trim(txtOwnerUsr.Text)) = 0 Then
                MsgBox "����ȷָ�����û�����", vbExclamation, gstrSysName
                txtOwnerUsr.SetFocus
                Exit Sub
            End If
            If Len(Trim(txtOwnerPwd.Text)) = 0 Then
                MsgBox "��ϵͳ�涨������ָ�����û����롣", vbExclamation, gstrSysName
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
            If txtOwnerPwd.Text <> txtOwnerLab.Text Then
                MsgBox "���뼰����֤�Ĳ����ϡ�", vbExclamation, gstrSysName
                txtOwnerPwd.Text = ""
                txtOwnerLab.Text = ""
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
            
            Call gobjRegister.CreateUser(gcnOracle, txtOwnerUsr.Text, Trim(txtOwnerPwd.Text), strError)
            If strError <> "" Then
                MsgBox "�û��������벻�������ݿ�Ҫ�������¶��塣" & vbCrLf & strError, vbExclamation, gstrSysName
                txtOwnerUsr.SetFocus
                Exit Sub
            End If
            txtOwnerUsr.Tag = "�Ѵ����û�"
            
        End If
        
        '------------------------------
        fraSetup(1).Visible = False
        
        If mbln���� = False Then
            fraSetup(2).Visible = True
        Else
            '����ǰ�װ���ף���������ռ������
            fraSetup(3).Visible = True
        End If
        
    ElseIf fraSetup(2).Visible Then
        '------------------------------------------------------------
        '��������
        '------------------------------------------------------------
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If Len(Trim(txtSpaceFile(intCount).Text)) = 0 Then
                MsgBox "�붨��" & txtSpaceName(intCount).Caption & "��ռ�������ļ���", vbExclamation, gstrSysName
                Exit Sub
            End If
            If Val(txtSpaceSize(intCount).Text) < Val(txtSpaceSize(intCount).Tag) Then
                MsgBox "��ռ�" & txtSpaceName(intCount).Caption & "�������" & txtSpaceSize(intCount).Tag & "M��", vbExclamation, gstrSysName
                txtSpaceSize(intCount).Text = txtSpaceSize(intCount).Tag
                Exit Sub
            End If
            
            If Val(txtSpaceSize(intCount).Text) > 10000 Then
                MsgBox "��ռ�" & txtSpaceName(intCount).Caption & "����10G�ˡ�", vbExclamation, gstrSysName
                Exit Sub
            End If
        Next
        Call tbsSpace_Click
        
        fraSetup(2).Visible = False
        fraSetup(3).Visible = True
        If optData(0).Visible = False Then
            fraSetup(3).Visible = False
            fraSetup(4).Visible = True
            'cmdNext.Caption = "���(&F)"
            lblStep(4).Caption = "���Ĳ� ��Ʒ��Ȩ��֤"
        End If
    
    ElseIf fraSetup(3).Visible Then
        '------------------------------------------------------------
        '���Ĳ���
        '------------------------------------------------------------
        fraSetup(3).Visible = False
        fraSetup(4).Visible = True
        cmdNext.Caption = "���(&F)"
        lblStep(4).Caption = "���岽 ���"
    
    ElseIf fraSetup(4).Visible Then
        '------------------------------------------------------------
        '���岽��
        '------------------------------------------------------------
        If chkEnjoy.value = 0 Then
            Set gcnTools = GetConnection("ZLTOOLS")
            If gcnTools Is Nothing Then Exit Sub
        End If
        
        gstrSQL = "    �Ѿ���������еİ�װ���ã�ϵͳ�������Զ���װ���̡�" & vbCr & vbCr _
                & "    ��װ���̿������нϳ�ʱ�䣬�벻Ҫ����ǿ���жϣ�����" & vbCr _
                & "�����ܲ�������������Ӱ��ϵͳ���С�" & vbCr & vbCr _
                & "   ������װ��"
        If MsgBox(gstrSQL, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
        cmdCancel.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        fraSetup(4).Enabled = False
        
        '����������
        mstrLogFile = GetLogPath(LT_��װ, intDefSysCode * 100 + mlng����)
        Set objfrmUpSys = New frmAppUpgradeNew '�������ģ�����
        If Not objfrmUpSys.ToolsInstallUp(Me, stbThis.Panels(2), intDefSysCode * 100 + mlng����, lblSetupFile.Caption, mstrLogFile) Then
            cmdNext.Enabled = True
            Unload Me
            Exit Sub
        End If
        
        If SysInstall() Then
            MsgBox "��װ�ɹ������������Ӧ�ó���װ������ʹ�ø�ϵͳ��", vbInformation, gstrSysName
            On Error Resume Next
            Shell "notepad " & mstrLogFile
            err.Clear: On Error GoTo 0
        Else
            MsgBox "��װʧ�ܣ�ϵͳ���Զ�����Ѿ���װ�����ݡ�", vbInformation, gstrSysName
            lblStep(4).Caption = "���ڳ�ж�Ѿ���װ�����ݡ�"
            DoEvents
            Call UnInstall
        End If
        cmdNext.Enabled = True
        Unload Me
    End If

End Sub

Private Sub cmdPrevious_Click()
    If fraSetup(4).Visible Then
        cmdNext.Caption = "��һ��(&N)"
        fraSetup(4).Visible = False
        fraSetup(3).Visible = True
        If lstData.Count = 1 Then
            fraSetup(3).Visible = False
            fraSetup(2).Visible = True
        End If
    ElseIf fraSetup(3).Visible Then
        fraSetup(3).Visible = False
        
        If mbln���� = False Then
            fraSetup(2).Visible = True
        Else
            '���װ�װʱ������ռ������
            fraSetup(1).Visible = True
        End If
    ElseIf fraSetup(2).Visible Then
        fraSetup(2).Visible = False
        fraSetup(1).Visible = True
    ElseIf fraSetup(1).Visible Then
        fraSetup(1).Visible = False
        fraSetup(0).Visible = True
        cmdPrevious.Enabled = False
    End If

End Sub

Private Sub Form_Load()
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    mbln���� = False 'ȱʡ��Ϊ����
    Call ApplyOEM(stbThis)
    With imgSetup
        .Top = PicSetup.ScaleTop
        .Left = PicSetup.ScaleLeft
        .Height = PicSetup.ScaleHeight
        .Width = PicSetup.ScaleWidth
    End With
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    
    '���ݵ�ǰϵͳ�������ļ�ȷ��ȱʡ�ı�ռ��ļ�·��
    With rsTemp
        gstrSQL = "select NAME from V$DATAFILE where ROWNUM<2 order by CREATION_TIME"
        .Open gstrSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then
            mstrTbsPath = "C:\"
        Else
            If InStr(1, StrReverse(!name), "\") > 0 Then
                mstrTbsPath = Mid(!name, 1, Len(!name) - InStr(1, StrReverse(!name), "\") + 1)
            ElseIf InStr(1, StrReverse(!name), "/") > 0 Then
                mstrTbsPath = Mid(!name, 1, Len(!name) - InStr(1, StrReverse(!name), "/") + 1)
            Else
                mstrTbsPath = "C:\"
            End If
        End If
    End With
    
    '������ֵ�ǰĿ¼���ڰ�װ��ֲ�ļ�����ֱ����д
    mstrIniPath = GetSetupPath(App.Path)
    If Dir(mstrIniPath & "\zlSetup.ini") <> "" Then
        lblSetupFile.Caption = mstrIniPath & "\zlSetup.ini"
        If ChkSetupFile() = False Then
            mbln���� = False
            mstrIniPath = ""
            lblSetupFile.Caption = ""
        End If
    End If
    
End Sub

Private Function GetSetupPath(ByVal strAppPath As String) As String
'�õ�ȱʡ�İ�װ·��
    Dim strPath() As String
    Dim strTemp As String
    
    ReDim strPath(0 To 0) As String
    
    strTemp = Dir(strAppPath & "\", vbDirectory)
    Do While strTemp <> ""
        strTemp = UCase(strTemp)
        If InStr(strTemp, ".") = 0 Then
            If strTemp <> "APPLY" And strTemp <> "TOOLS" And strTemp <> "�����ļ�" Then
                ReDim Preserve strPath(0 To UBound(strPath) + 1)
                strPath(UBound(strPath)) = strTemp
            End If
        End If
        strTemp = Dir(, vbDirectory)
    Loop
    If UBound(strPath) = 1 Then
        GetSetupPath = strAppPath & "\" & strPath(1) & "\Ӧ�ýű�"
    Else
        GetSetupPath = strAppPath
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdNext.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
    Set mclsRunScript = Nothing

    Set objText = Nothing
    Set mcnOwner = Nothing
End Sub

Private Sub cmbEnjoy_Click()
    Dim rsTemp As New ADODB.Recordset
    
    With rsTemp
        gstrSQL = "select ������ from zlSystems where ���=" & cmbEnjoy.ItemData(cmbEnjoy.ListIndex)
        .Open gstrSQL, gcnOracle, adOpenKeyset
        cmbEnjoy.Tag = !������
        If txtOwnerUsr.Enabled = False Then
            txtOwnerUsr.Text = cmbEnjoy.Tag
        End If
    End With
End Sub


Private Sub lstData_Click(Index As Integer)
    SetPromptText Split(lstData(Index).Tag, "=")(lstData(Index).ListIndex + 1)
End Sub

Private Sub optData_Click(Index As Integer)
    SetPromptText optData(Index).ToolTipText
    For intCount = 0 To optData.UBound
        If intCount = Index And lstData(Index).ListCount > 0 Then
            lstData(intCount).Visible = True
            lblNoData.Visible = False
        Else
            lstData(intCount).Visible = False
            lblNoData.Visible = True
        End If
    Next
    
    cmdSelectAll.Visible = lstData(Index).ListCount > 0
    cmdClearAll.Visible = lstData(Index).ListCount > 0
    
    With lblNoData
        .Left = lstData(0).Left
        .Width = lstData(0).Width
        .Top = lstData(0).Top
        .Height = lstData(0).Height
        .Caption = vbCrLf & "     " & optData(Index).Caption & "�����鲻����ϸ�ֵĿ�ѡ�����"
    End With
End Sub


Private Sub tbsSpace_Click()
    For intCount = 0 To tbsSpace.Tabs.Count - 1
        txtSpaceName(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceFile(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceSize(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        chkSpaceExtd(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        cboSpaceExtentType(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceExtentSize(intCount).Visible = tbsSpace.Tabs(intCount + 1).Selected
        txtSpaceExtentSize(intCount).Enabled = (cboSpaceExtentType(intCount).ListIndex = 1)
        '������ռ����־����
        If tbsSpace.Tabs(intCount + 1).Selected Then
            chkLogin(intCount).Visible = UCase(txtSpaceName(intCount).Caption) Like "ZL9INDEX*"
        Else
            chkLogin(intCount).Visible = False
        End If
       
        If tbsSpace.Tabs(intCount + 1).Selected Then
            txtSpaceFile(intCount).SetFocus
        End If
    Next
End Sub

Private Sub txtOwnerUsr_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSpaceExtentSize_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub SetProgressVisible(ByVal blnVisible As Boolean)
    If blnVisible = True Then
        If stbThis.Panels.Count = 3 Then
            '����һ������
            stbThis.Panels.Add 3
            stbThis.Panels(3).AutoSize = sbrSpring
            stbThis.Panels(2).AutoSize = sbrNoAutoSize
            stbThis.Panels(2).MinWidth = 1440
        End If
        pgbState.Left = stbThis.Panels(3).Left + 30
        pgbState.Width = stbThis.Panels(4).Left - pgbState.Left - 150
        pgbState.Top = stbThis.Top + stbThis.Height / 3
        pgbState.Visible = True
    Else
        If stbThis.Panels.Count = 4 Then
            stbThis.Panels(2).AutoSize = sbrSpring
            stbThis.Panels.Remove 3
        End If
        pgbState.Visible = False
    End If
    
End Sub

Private Function ChkSetupFile(Optional blnMsg As Boolean) As Boolean
    Dim strTemp As String
    '-------------------------------------
    '�����Ͱ�װ�����ļ�����ȷ��
    '-------------------------------------
    mstrIniPath = Mid(lblSetupFile.Caption, 1, Len(lblSetupFile.Caption) - 11)
    '����ļ�ƥ���Լ��
    strTemp = ""
    If Dir(mstrIniPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & mstrIniPath & "zlSequence.sql"
    If Dir(mstrIniPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "���ݱ��ļ�" & mstrIniPath & "zlTable.sql"
    If Dir(mstrIniPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "Լ���ļ�" & mstrIniPath & "zlConstraint.sql"
    If Dir(mstrIniPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & mstrIniPath & "zlIndex.sql"
    If Dir(mstrIniPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "��ͼ�ļ�" & mstrIniPath & "zlView.sql"
    If Dir(mstrIniPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "���������ļ�" & mstrIniPath & "zlProgram.sql"
    
    '�����,��Ϊ9ϵͳû�д��ļ�
    'If Dir(mstrIniPath & "zlPackage.sql") = "" Then strTemp = strTemp & vbCr & "���ļ�" & mstrIniPath & "zlPackage.sql"
    
    If Dir(mstrIniPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "���������ļ�" & mstrIniPath & "zlManData.sql"
    If Dir(mstrIniPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "Ӧ�������ļ�" & mstrIniPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "���·�������װ������ļ���ʧ�����ܼ�����������" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '��װ�����ļ�����
    err = 0
    On Error Resume Next
    Set objText = gobjFile.OpenTextFile(lblSetupFile.Caption)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        intDefSysCode = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        strDefSysName = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�汾��]" Then
        strDefVersion = Trim(Mid(strTemp, 6))
    
        '�ж��Ƿ�Ӧ�ðѱ��ΰ�װ��Ϊ���װ�װ
        Dim lngTemp As Long
        Dim lngMax As Long        '�������׺�
        Dim blnHase  As Boolean   '�Ƿ���ͬϵͳ����
        Dim lngMin As Long
        Dim lstTemp As ListItem
        
        
        mbln���� = False
        mlng���� = 0
        lngMin = 99
        For Each lstTemp In frmAppStart.lvwSys.ListItems
            lngTemp = Mid(lstTemp.Key, 2)
            If lngTemp \ 100 = intDefSysCode Then
                'ϵͳ��ͬ
                blnHase = True
                If lngMax < lngTemp Mod 100 Then
                    lngMax = lngTemp Mod 100 '�����������׺�
                End If
                
                If strDefVersion = lstTemp.SubItems(1) Then
                    '�汾Ҳ��ͬ���ǾͿ�����
                    mbln���� = True
                    If lngMin > lngTemp Mod 100 Then
                        lngMin = lngTemp Mod 100 '������С�����׺�
                        Set mlst��׼ = lstTemp 'ȡ��С���׺���Ϊ��׼����
                    ElseIf lngMin = 99 Then
                        Set mlst��׼ = lstTemp '��ʼ�������ȡһ����Ϊ��׼����
                    End If
                End If
            End If
        Next
        If blnHase = True Then
            '��ͬϵͳ�İ�װ
            If mbln���� = False Then
                If blnMsg Then MsgBox "��ǰ���ݿ���Ҳ����ͬ���͵�ϵͳ���ڣ������ڰ汾����������������", vbInformation, gstrSysName
                Exit Function
            Else
                If blnMsg = False Then
                    Exit Function
                Else
                    If lngMax >= 99 Then
                        MsgBox "��ǰ���ݿ���Ҳ����ͬ���͵�ϵͳ���ڣ��������㹻�࣬����������", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If MsgBox("��ǰ���ݿ�������" & strDefSysName & "ϵͳ���ڣ����Ƿ�Ҫ������һ����", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                    mlng���� = lngMax + 1
                End If
            End If
        End If
    Else
        err.Raise 10
    End If
    Caption = "Ӧ��ϵͳ��װ" & " - " & strDefSysName & " V" & strDefVersion
    lblSysName.Caption = "ϵͳ����" & strDefSysName
    lblVersion.Caption = "�汾�ţ�" & strDefVersion
        
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[��ռ�]" Then
        strDefSpace = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�û���]" Then
        strDefUser = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[������]" Then
        strDefData = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    mstrExtSysCode = ""
    mstrExtVersion = ""
    If Not objText.AtEndOfStream Then
        '������չϵͳ������
        strTemp = Trim(objText.ReadLine)
        If Left(strTemp, 5) = "[��ϵͳ]" Then
            mstrExtSysCode = Trim(Mid(strTemp, 6))
            
            strTemp = Trim(objText.ReadLine)
            If Left(strTemp, 5) = "[���汾]" Then
                mstrExtVersion = Trim(Mid(strTemp, 6))
            Else
                mstrExtSysCode = ""
            End If
        End If
    End If
    Call FillShare  '�õ������嵥
    If mstrExtSysCode <> "" And cmbEnjoy.ListCount = 0 Then
        If blnMsg Then MsgBox "����չϵͳû�ҵ�����������ϵͳ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If err <> 0 Then
        If blnMsg Then MsgBox "��װ�����ļ���ʧ����ȷ��", vbExclamation, gstrSysName
        Exit Function
    End If
    objText.Close
    
    '��ռ���ȷ�Լ��
    intItems = tbsSpace.Tabs.Count
    For intCount = 0 To intItems - 2
        tbsSpace.Tabs.Remove 1
    Next
    
    err = 0
    On Error Resume Next
    aryRow = Split(strDefSpace, "||")
    For intCount = 0 To UBound(aryRow)
        aryVal = Split(aryRow(intCount), "|")
        If intCount = 0 Then
            tbsSpace.Tabs(1).Caption = aryVal(0)
            tbsSpace.Tabs(1).Key = aryVal(1)
        Else
            tbsSpace.Tabs.Add , aryVal(1), aryVal(0)
        End If
        If intCount > txtSpaceName.Count - 1 Then Load txtSpaceName(intCount)
        If intCount > txtSpaceFile.Count - 1 Then Load txtSpaceFile(intCount)
        If intCount > txtSpaceSize.Count - 1 Then Load txtSpaceSize(intCount)
        If intCount > chkSpaceExtd.Count - 1 Then Load chkSpaceExtd(intCount)
        If intCount > cboSpaceExtentType.Count - 1 Then Load cboSpaceExtentType(intCount)
        If intCount > txtSpaceExtentSize.Count - 1 Then Load txtSpaceExtentSize(intCount)
        '������ռ����־����
        If intCount > chkLogin.Count - 1 Then Load chkLogin(intCount)
        
        txtSpaceName(intCount).Caption = aryVal(1)
        If UCase(aryVal(1)) Like "ZL9INDEX*" Then
            chkLogin(intCount).value = 0
        Else
            chkLogin(intCount).value = 1
        End If
        chkLogin(intCount).Visible = False
        
        txtSpaceFile(intCount).Tag = aryVal(1)
        txtSpaceFile(intCount).Text = mstrTbsPath & txtSpaceFile(intCount).Tag & ".DBF"
        txtSpaceSize(intCount).Text = aryVal(2)
        txtSpaceSize(intCount).Tag = aryVal(3)
        
        If aryVal(4) = "T" Then
            chkSpaceExtd(intCount).value = 1
        Else
            chkSpaceExtd(intCount).value = 0
        End If
        
        
        '��ռ�����������
        cboSpaceExtentType(intCount).Clear
        cboSpaceExtentType(intCount).AddItem "�Զ��������ߴ�"
        cboSpaceExtentType(intCount).AddItem "ͳһ�������ߴ�"
        cboSpaceExtentType(intCount).ListIndex = 0
        txtSpaceExtentSize(intCount).Text = 1
        txtSpaceExtentSize(intCount).Enabled = (cboSpaceExtentType(intCount).ListIndex = 1)
    Next
    If err <> 0 Then
        If blnMsg Then MsgBox "��װ�����ļ���ռ����ô��󣬲��ܼ�����װ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If mstrExtSysCode = "" Then
        '����չϵͳ��������
        If mbln���� = False Then
            'û�ж����ף����ܹ���
            If cmbEnjoy.ListCount = 0 Then
                chkEnjoy.value = 0
                chkEnjoy.Enabled = False
            Else
                chkEnjoy.Enabled = True
            End If
            If chkEnjoy.value <> 1 Then
                txtOwnerUsr.Text = strDefUser
            End If
            
        Else
            chkEnjoy.Enabled = False '������ѡ����ֻ������
            chkEnjoy.value = 0
            txtOwnerUsr.Text = strDefUser & mlng����
            'Ҳ�����ڼ���ռ�����ã���Ϊ����ǰ��
        End If
    Else
        '���ݺϲ�����жϹ������
        chkEnjoy.Enabled = False '�����ϵͳ������չϵͳ����ô����ѡ������
        chkEnjoy.value = 1
    End If
    
    '���ݷ����ѡ�ļ�ƥ���Լ��
    Dim intOptions As Integer       'ѡ������Ŀ
    Dim lngHeight As Long           '�ؼ����и߶�
    
    For intCount = 0 To optData.UBound
        optData(intCount).Visible = False
    Next
    For intCount = 0 To lstData.UBound
        lstData(intCount).Visible = False
    Next
    
    intOptions = 0
    err = 0
    aryRow = Split(strDefData, "||")
    For intCount = 0 To UBound(aryRow)
        If Dir(mstrIniPath & "zlSelData" & intCount & ".sql") <> "" Then
            If intOptions > optData.Count - 1 Then Load optData(intOptions)
            optData(intOptions).Tag = intCount
            optData(intOptions).Left = optData(0).Left
            optData(intOptions).ToolTipText = ""
            optData(intOptions).Visible = True
            
            If intOptions > lstData.Count - 1 Then Load lstData(intOptions)
            lstData(intOptions).Left = fraSelData.Width / 2 - 300
            lstData(intOptions).Width = fraSelData.Width / 2 - optData(0).Left + 300
            lstData(intOptions).Top = lstData(0).Top
            lstData(intOptions).Tag = ""
            lstData(intOptions).Clear
            intItems = InStr(1, aryRow(intCount), ">")
            If intItems = 0 Then
                If InStr(1, aryRow(intCount), "=") = 0 Then
                    optData(intOptions).Caption = Trim(aryRow(intCount))
                Else
                    optData(intOptions).Caption = Trim(Left(aryRow(intCount), InStr(1, aryRow(intCount), "=") - 1))
                    optData(intOptions).ToolTipText = Trim(Mid(aryRow(intCount), InStr(1, aryRow(intCount), "=") + 1))
                End If
            Else
                optData(intOptions).Caption = Trim(Mid(aryRow(intCount), 1, intItems - 1))
                If InStr(1, optData(intOptions).Caption, "=") > 0 Then
                    optData(intOptions).ToolTipText = Trim(Mid(optData(intOptions).Caption, InStr(1, optData(intOptions).Caption, "=") + 1))
                    optData(intOptions).Caption = Trim(Left(optData(intOptions).Caption, InStr(1, optData(intOptions).Caption, "=") - 1))
                End If
                strTemp = Mid(aryRow(intCount), intItems + 1)
                aryVal = Split(strTemp, "|")
                For intItems = 0 To UBound(aryVal)
                    If Dir(mstrIniPath & "zlSelData" & intCount & intItems & ".sql") <> "" Then
                        If InStr(1, aryVal(intItems), "=") = 0 Then
                            lstData(intOptions).AddItem Trim(aryVal(intItems))
                            lstData(intOptions).Tag = lstData(intOptions).Tag & "="
                        Else
                            lstData(intOptions).AddItem Trim(Left(aryVal(intItems), InStr(1, aryVal(intItems), "=") - 1))
                            lstData(intOptions).Tag = lstData(intOptions).Tag & Mid(aryVal(intItems), InStr(1, aryVal(intItems), "="))
                        End If
                        lstData(intOptions).ItemData(lstData(intOptions).NewIndex) = Val(intItems)
                    End If
                Next
            End If
            intOptions = intOptions + 1
        End If
    Next
    cmdSelectAll.Left = lstData(0).Left
    cmdClearAll.Left = lstData(0).Left + lstData(0).Width - cmdClearAll.Width
    
    
    If err <> 0 Then
        If blnMsg Then MsgBox "��װ�����ļ����ݷ������ô��󣬲��ܼ�����װ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If intOptions = 1 Then
        optData(0).Top = lstData(0).Top
        optData(0).Height = lstData(0).Height
        lstData(0).Left = optData(0).Left
        lstData(0).Width = fraSelData.Width - optData(0).Left * 2
        lstData(0).ZOrder
    ElseIf intOptions <> 0 Then
        lngHeight = lstData(0).Height / intOptions
        For intCount = 0 To intOptions - 1
            optData(intCount).Top = lstData(0).Top + intCount * lngHeight
            optData(intCount).Height = lngHeight
        Next
    Else
        lblNoData.Visible = True
    End If
    If intOptions <> 0 Then
        optData(0).value = True
        Call optData_Click(0)
    End If

    '˳���ע���ļ�Ҳһ�������
    Call ChkRegFile
    SetPromptText ""
    
    ChkSetupFile = True
End Function

Private Sub ChkRegFile()
    '�ж�ϵͳ��Ȩ
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo errHand
    gstrSQL = "Select Count(*) From zltools.Zlregfunc f, zltools.Zlreginfo r, zltools.zlRegAudit t Where r.��Ŀ = '��Ȩ֤��' And f.ϵͳ = " & intDefSysCode
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.Fields(0).value > 0 Then
        Me.lblRegAudit.Caption = "    �Ѿ��߱���ϵͳӦ����Ȩ��������װ�غ�������Ȩʹ�á�"
        Exit Sub
    End If
errHand:
    Me.lblRegAudit.Caption = "    ���ڻ����߱���ϵͳӦ����Ȩ����Ȼ���Լ���װ�أ����޷�������Ȩʹ�ã�"
End Sub

Private Sub FillShare()
'�������õĹ����嵥
    Dim rsTemp As New ADODB.Recordset
    Dim varVersion As Variant, varExtVersin As Variant
    Dim i As Long, bln���� As Boolean
    
    cmbEnjoy.Clear
    If mstrExtSysCode = "" Then
        '��ϵͳ������չϵͳ���ɹ�������ϵͳ
        gstrSQL = "select ���,���� from zlsystems order by ���"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            cmbEnjoy.AddItem rsTemp("����") & "(" & rsTemp("���") & ")"
            cmbEnjoy.ItemData(cmbEnjoy.NewIndex) = rsTemp("���")
            rsTemp.MoveNext
        Loop
    Else
        '����չϵͳ���Ǳ���Ҫ�����������ж�
        '1)ϵͳ�����
        '2)û����������ͬϵͳ��չ
        '3)�汾���ܵ���Ҫ��
        gstrSQL = "select A.���,A.����,A.�汾�� from zlsystems A " & _
                  "  Where floor(A.��� / 100) = " & mstrExtSysCode & _
                  "        and not exists (select B.��� from zlsystems B where B.�����=A.��� and floor(B.���/100)=" & intDefSysCode & ")"
        
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        
        varExtVersin = Split(mstrExtVersion, ".")
        Do Until rsTemp.EOF
            '�жϰ汾
            bln���� = True
            varVersion = Split(rsTemp("�汾��"), ".")
            
            For i = LBound(varExtVersin) To UBound(varExtVersin)
                If Val(varExtVersin(i)) > Val(varVersion(i)) Then
                    '�ű��еİ汾�Ÿ���ʵ�����ݿ�ģ���������
                    bln���� = False
                    Exit For
                ElseIf Val(varExtVersin(i)) < Val(varVersion(i)) Then
                    '�Ѿ����㣬��Ҫ�ٱȽ���һλ
                    Exit For
                End If
            Next
            
            If bln���� = True Then
                '��������
                cmbEnjoy.AddItem rsTemp("����") & "(" & rsTemp("���") & ")"
                cmbEnjoy.ItemData(cmbEnjoy.NewIndex) = rsTemp("���")
            End If
            rsTemp.MoveNext
        Loop
        
    End If
End Sub
Private Function SysInstall() As Boolean
    '----------------------------------
    '���ܣ����ϵͳ�İ�װ����
    '---------��װ�㷨-----------------
    '    ������ϵͳ���ݱ�ռ�
    '    If not �����Ѿ���װ��ϵͳ Then
    '        ������ϵͳ������
    '        �ɹ��������������Ҫ�Ĺ������ݶ���Ȩ��
    '    End If
    '    ������ϵͳ���ݶ���
    '    �������ݼ���ѡ���ݰ�װ
    '----------------------------------
    Dim rsTemp As New ADODB.Recordset, cnCtxsys As New ADODB.Connection
    Dim strTmpSpace As String
    Dim strTemp As String, strError As String
    Dim intVer As Integer
    Dim blnIgnoreErr As Boolean     '���Դ���
    Dim strPassword As String, strUserName As String, lngAbort As Long, cllRoles As Collection
    
    strUserName = UCase(Trim(txtOwnerUsr.Text))
    On Error GoTo errHand
    intVer = GetOracleVersion
    gcnOracle.Execute "Grant Select on sys.v_$session to Public"
    gcnOracle.Execute "Grant Select on sys.v_$parameter to Public"
        
    With rsTemp
        gstrSQL = "SELECT TEMPORARY_TABLESPACE FROM DBA_USERS WHERE USERNAME='ZLTOOLS'"
        .Open gstrSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then SysInstall = False: Exit Function
        strTmpSpace = .Fields(0).value
    End With
    
    If mbln���� = False Then
        '������ϵͳ���ݱ�ռ�
        SetPromptText "������ռ�"
        pgbState.value = 0
        
        '
        SetProgressVisible True
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If CreateTbs(txtSpaceName(intCount).Caption, _
                            txtSpaceFile(intCount).Text, _
                            txtSpaceSize(intCount).Text, _
                            chkSpaceExtd(intCount).value, _
                            False, _
                            cboSpaceExtentType(intCount).ListIndex = 0, _
                            Val(txtSpaceExtentSize(intCount).Text), _
                            chkLogin(intCount).value = 0 _
                            ) <> 1 Then
                GoTo errHand
            End If
            pgbState.value = (intCount + 1) / tbsSpace.Tabs.Count * 100
            DoEvents
        Next
        pgbState.value = 0
        SetProgressVisible False
    End If
    
    '����������Ѿ���װ��ϵͳ
    If chkEnjoy.value = 0 Then
        '������ϵͳ������
        SetPromptText "��Ȩ��������" & strUserName
        pgbState.value = 0
        SetProgressVisible True
                
        '�ڵڶ���ʱ�Ѵ���
        
        gstrSQL = "alter user " & strUserName & _
                " DEFAULT TABLESPACE " & txtSpaceName(0).Caption & _
                " TEMPORARY TABLESPACE " & strTmpSpace
        gcnOracle.Execute gstrSQL
        
        '12c��resource��ɫȱʡû��UNLIMITED TABLESPACEȨ��
        '����CREATE TRIGGERȨ�ޣ�������ʷ����ת���Ĵ洢������Ϊ��������ı���ʱ������������ת����ɾ����
        '�洢������execute immediateִ�ж�̬SQLʱ������ʾ��Ȩ����ʹ������ɫ��Ȩ�ޣ����磺��ʹ��DBA����Ȼ��Ҫ��Ȩ��
        gstrSQL = "Grant Connect,Resource," & IIf(chkDBA.value = 1, "DBA,", "") & _
                " UNLIMITED TABLESPACE,Create Table,Create Sequence,Create Role,Create User,Drop User,Alter User,Create Public Synonym,Drop Public Synonym," & _
                " Alter Session,Create Session,Create Synonym,Create View,Create Database Link,Create Cluster," & _
                " Create Materialized View, Alter Any Materialized View, Drop Any Materialized View,CREATE TRIGGER" & _
                " to " & strUserName & " With Admin Option"
        gcnOracle.Execute gstrSQL
        gstrSQL = "Grant Select on sys.dba_role_privs to " & strUserName & " With Grant Option"
        gcnOracle.Execute gstrSQL
        gstrSQL = "Grant Select on sys.dba_roles to " & strUserName
        gcnOracle.Execute gstrSQL
        gstrSQL = "Grant Execute on sys.dbms_sql to " & strUserName & " With Grant Option"
        gcnOracle.Execute gstrSQL
        
        gstrSQL = "Grant Select on sys.gv_$session to " & strUserName & " With Grant Option"
        gcnOracle.Execute gstrSQL
     
        On Error Resume Next '����ȫ�ļ����Ĳ������п���û�и��û������԰Ѵ�������
        gstrSQL = "Grant CTXAPP to " & strUserName & " With Admin Option"
        gcnOracle.Execute gstrSQL
        gcnOracle.Execute "alter user ctxsys identified by ctxsys"
        gcnOracle.Execute "alter user  ctxsys account Unlock"
        cnCtxsys.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServer, "ctxsys", "ctxsys"
        cnCtxsys.Execute "Grant Execute on ctx_ddl to " & strUserName & " With Grant Option" 'Ϊ���ڹ�����ִ�а�����
        
        err = 0: On Error GoTo errHand

        '�ɹ��������������Ҫ�Ĺ������ݶ���Ȩ��
        SetPromptText "�����߶���Ȩ������" & strUserName
        SetProgressVisible False
        Call ReGrantForTools(gcnTools, strUserName)
    End If
    
    '��д��װϵͳ�嵥
    gstrSQL = "insert into zlSystems(���,�����,����,������,��װ����,������װ,�汾��)" & _
            " values(" & intDefSysCode * 100 + mlng���� '��ŵ�ǰ��λ��ϵͳ�ţ�����λ�����׺�
    If chkEnjoy.value = 1 Then
        gstrSQL = gstrSQL & "," & cmbEnjoy.ItemData(cmbEnjoy.ListIndex)
    Else
        gstrSQL = gstrSQL & ",null"
    End If
    gstrSQL = gstrSQL & ",'" & strDefSysName & "'"
    gstrSQL = gstrSQL & ",'" & strUserName & "'"
    gstrSQL = gstrSQL & ",sysdate,0,'" & strDefVersion & "')"
    gcnOracle.Execute gstrSQL
    
    
    '������ϵͳ���ݶ���
    Set mcnOwner = gobjRegister.GetConnection(gstrServer, strUserName, Trim(txtOwnerPwd.Text), True, MSODBC, "", False)
    strPassword = gobjRegister.GetPassword
    Call SetSQLTrace(gstrServer, strUserName, mcnOwner)
    
    Set cllRoles = New Collection
    
    '�Ƿ���Դ���
    blnIgnoreErr = chkEnjoy.value <> 0
    If gblnInIDE Then blnIgnoreErr = False
    Set mclsRunScript = New clsRunScript
    Call mclsRunScript.InitGlobalPara(Me, intDefSysCode * 100 + mlng����, blnIgnoreErr, mstrLogFile)
    Call mclsRunScript.InitUserList(strUserName, strPassword)
    Set mclsRunScript.Connection = mcnOwner: mclsRunScript.ConnectType = 0
    mclsRunScript.IsRoleCollect = True '�ռ���ɫ
    
    SetProgressVisible True
    SetPromptText "��������"
    If RunSQLScript(mstrIniPath & "zlSequence.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If

    SetPromptText "�������ݱ�"
    If RunSQLScript(mstrIniPath & "zlTable.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If

    SetPromptText "����Լ��"
    If RunSQLScript(mstrIniPath & "zlConstraint.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    
    SetPromptText "��������"
    If RunSQLScript(mstrIniPath & "zlIndex.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    SetPromptText "������ͼ"
    If RunSQLScript(mstrIniPath & "zlView.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If

    SetPromptText "���������"
    If RunSQLScript(mstrIniPath & "zlProgram.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    
    If Dir(mstrIniPath & "zlPackage.sql") <> "" Then
        SetPromptText "������"
        If RunSQLScript(mstrIniPath & "zlPackage.sql") = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    End If
    Set cllRoles = mclsRunScript.Roles
    
    If cllRoles.Count <> 0 Then
        '��Ҫ����صĽ�ɫ������Ȩ
        SetPromptText "��ɫ��Ȩ����"
        If GrantToRole(mcnOwner, cllRoles, strUserName) = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    End If
    
    
    If chkEnjoy.value <> 0 Then
        '����װʱ����Ҫ���±��루��Ϊ������ͬ���Ĵ洢���̱����´����ˣ�
        SetPromptText "�������"
        Call ReCompileProcedure(mcnOwner)
    End If
    
   
    '��������
    SetPromptText "�������ݰ�װ"
    If mbln���� = False Then
        If RunSQLScript(mstrIniPath & "zlManData.sql") = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    Else
        'ͨ�����ݿ��п����õ�
        If CopyManageData(mcnOwner) = False Then GoTo errHand
    End If
    SetPromptText "Ӧ�����ݰ�װ"
    If RunSQLScript(mstrIniPath & "zlAppData.sql") = False Then
        SetProgressVisible False: GoTo errHand:
    End If
    
    '��װ����
    SetPromptText "�̶�����װ"
    If mbln���� = False Then
        If RunSQLScript(mstrIniPath & "zlReport.sql") = False Then
            SetProgressVisible False: GoTo errHand:
        End If
    Else
        'ͨ�����ݿ��п����õ�
        If CopyReport(mcnOwner, Mid(mlst��׼.Key, 2), intDefSysCode * 100 + mlng����) = False Then GoTo errHand
    End If
    
    '��ѡ���ݰ�װ
    If chkSelData.value = 1 Then
        For intCount = 0 To optData.UBound
            If optData(intCount).value = True Then
                SetPromptText optData(intCount).Caption
                If RunSQLScript(mstrIniPath & "zlSelData" & optData(intCount).Tag & ".sql") = False Then
                    SetProgressVisible False: GoTo errHand:
                End If
                
                For intItems = 0 To lstData(intCount).ListCount - 1
                    If lstData(intCount).Selected(intItems) = True Then
                        SetPromptText lstData(intCount).List(intItems)
                        If RunSQLScript(mstrIniPath & "zlSelData" & optData(intCount).Tag & lstData(intCount).ItemData(intItems) & ".sql") = False Then
                            SetProgressVisible False: GoTo errHand:
                        End If
                    End If
                Next
            End If
        Next
    End If
    
    '������װ���µ�������ʵ����ֵ��ƥ��
    SetPromptText "���м��"
    DoEvents
    Call ChkSequence
    
    '��д��װ��¼Ϊ������װ
    gstrSQL = "update zlSystems set ������װ=1 where ���=" & intDefSysCode * 100 + mlng����
    gcnOracle.Execute gstrSQL
    gstrSQL = "insert into zlSysFiles(ϵͳ,����,�ļ���,����,������)" & _
            " values (" & intDefSysCode * 100 + mlng���� & ",1,'" & lblSetupFile.Caption & "',sysdate,user)"
    gcnOracle.Execute gstrSQL
    
    If CheckHavHistory(intDefSysCode * 100 + mlng����) Then
        '���˺飺���봴����ʷ���ݿռ�
        If frmHistorySpaceSet.ShowInstall(Me, mcnOwner, strUserName, _
            strPassword, intDefSysCode * 100 + mlng����, 0, 0) = False Then
            If mcnOwner.State = adStateOpen Then mcnOwner.Close
            Set mcnOwner = Nothing
            Exit Function
        End If
    End If
    
    '������ǰ�����ߵ�ȫ������Ĺ���ͬ���('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
    mcnOwner.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
    
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    Set mcnOwner = Nothing
    Set mclsRunScript = Nothing
    SysInstall = True
    Exit Function

errHand:
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    
    Set mcnOwner = Nothing
    SetProgressVisible False
    SysInstall = False
End Function

Private Sub SetPromptText(ByVal strText As String)
    stbThis.Panels(2).Text = strText
    stbThis.Panels(2).ToolTipText = strText
End Sub

Private Function UnInstall() As Boolean
    '----------------------------------
    '���ܣ�ɾ���Ѿ��İ�װ����
    '----------------------------------
    Dim rsTemp As New ADODB.Recordset, rsSys As New ADODB.Recordset, blnCanRemoveMSGData As Boolean
    Dim strSpaces As String, strFiles As String, aryFile() As String, strErrInfo As String
    Dim lngRowH As Long
    
    
    If mbln���� = False Then
        '���������ļ�
        strSpaces = ""
        If intDefSysCode = 1 Or intDefSysCode = 25 Then
            'ZLMSGDATA��ռ��׼��Ҳ���ڣ�LISҲ���ڣ���ֻ����һ��ϵͳʱ����ֱ��ж��
            gstrSQL = "Select Count(1) ����" & vbNewLine & _
                        "From Zlsystems" & vbNewLine & _
                        "Where Floor(��� / 100) In (1, 25)"
            rsSys.Open gstrSQL, gcnOracle
            blnCanRemoveMSGData = rsSys!���� = 1
        Else
            blnCanRemoveMSGData = True 'û��ZLMSGDATA��ռ䣬Ϊ�˼��߼�
        End If
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If blnCanRemoveMSGData Then
                strSpaces = strSpaces & ",'" & UCase(Trim(txtSpaceName(intCount).Caption)) & "'"
            ElseIf UCase(Trim(txtSpaceName(intCount).Caption)) <> "ZLMSGDATA" Then
                strSpaces = strSpaces & ",'" & UCase(Trim(txtSpaceName(intCount).Caption)) & "'"
            End If
            DoEvents
        Next
        strFiles = ""
        With rsTemp
            gstrSQL = "select F.NAME from V$TABLESPACE T,V$DATAFILE F where T.TS#=F.TS#  and T.NAME in(" & Mid(strSpaces, 2) & ")"
            .Open gstrSQL, gcnOracle
            Do While Not .EOF
                strFiles = strFiles & ";" & .Fields(0).value
                DoEvents
                .MoveNext
            Loop
        End With
    End If
    strErrInfo = ""
    err = 0
    On Error Resume Next
    
    SetPromptText "��������Ѱ�װ�����ݡ�"
    'ɾ����װ��¼
    gstrSQL = "delete from zlSystems where ���=" & intDefSysCode * 100 + mlng����
    gcnOracle.Execute gstrSQL
    
    '������Ч�˵�
    With rsTemp
        Do
            If .State = adStateOpen Then .Close
            gstrSQL = "select 1 from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            .Open gstrSQL, gcnOracle
            If .EOF Then Exit Do
            gstrSQL = "delete from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            gcnOracle.Execute gstrSQL
        Loop
    End With
    
    'ɾ����ϵͳ������
    If chkEnjoy.value = 0 Then
        SetPromptText "����ɾ���Ѵ������û���"
        intCount = 0
        Do
            gcnOracle.Execute "drop user " & txtOwnerUsr.Text & " cascade"
            With rsTemp
                If .State = adStateOpen Then .Close
                .Open "select * from all_users where username='" & UCase(txtOwnerUsr.Text) & "'", gcnOracle
                If .EOF Then Exit Do
            End With
            intCount = intCount + 1
            DoEvents
            If intCount > 10000 Then
                strErrInfo = strErrInfo & vbCr & "�û�:" & txtOwnerUsr
                Exit Do
            End If
        Loop
    End If
    
    If mbln���� = False Then
        'ɾ����ϵͳ���ݱ�ռ�
        SetPromptText "����ɾ���Ѵ����ı�ռ�������ļ���"
        For intCount = 0 To tbsSpace.Tabs.Count - 1
            If CheckSpaceIsUse("��ռ�", txtSpaceName(intCount).Caption, txtOwnerUsr.Text) = False Then
                'û�������û�ʹ�ã�����ɾ��
                gcnOracle.Execute "alter tablespace " & txtSpaceName(intCount).Caption & " offline"
                
                DoEvents
                gcnOracle.Execute "drop tablespace " & txtSpaceName(intCount).Caption & " including contents and datafiles cascade constraints"
            End If
        Next
        
        'ȡ��ֱ��ɾ���ļ�����䣬��Ϊ�������ļ���һ�������ݿ�������ļ�
    End If
    
    SetPromptText ""
    If strErrInfo <> "" Then
        MsgBox "��������Oracle��,�ֹ�ɾ���������ݣ�" & strErrInfo, vbExclamation, gstrSysName
    Else
        MsgBox "����Ӳ�̿ռ�����ݿ�ϵͳ��ȷ����������°�װ��", vbExclamation, gstrSysName
    End If
End Function


Private Function CreateTbs(TbsName As String, TbsFile As String, TbsSize As Integer, Optional AutoExtend As Boolean, _
     Optional Temp As Boolean, Optional AutoAllocate As Boolean, Optional ExtentSize As Integer, Optional Nologging As Boolean) As Byte
    '----------------------------------------------
    '���ܣ�ϵͳ�û�,���ݲ���������ռ�,�̶�Ϊ���ع�������(8i��ǰ��֧��,��ʱֻ�ܴ����ֵ��������)
    '       ������漰LOB�ֶε�ԭ��,������ASSM��ռ�(��9i����֧��,SEGMENT SPACE MANAGEMENT AUTO)
    '������
    '   TbsName:��ռ�����
    '   TbsFile:��ռ��ļ�
    '   TbsSize:��ռ��С(MΪ��λ)
    '   Extend:�Ƿ��Զ�������,����ͳһ��Χ�ߴ�
    '   ExtentSize:ͳһ���ߴ�,��ʱ��ռ����ָ���ߴ�(OracleȱʡΪ1M)
    '   Temp:�Ƿ�Ϊ��ʱ��ռ�
    '���أ�1-�����ɹ���2-��ռ��Ѿ����ڣ�3-����ʧ��
    '----------------------------------------------
    DoEvents
    If Temp Then
        gstrSQL = "CREATE TEMPORARY TABLESPACE " & TbsName & " TEMPFILE '" & TbsFile & "'"
    Else
        gstrSQL = "CREATE TABLESPACE " & TbsName & " DATAFILE '" & TbsFile & "'"
    End If
    gstrSQL = gstrSQL & _
            " SIZE " & TbsSize & "M REUSE " & _
             IIf(AutoExtend, "AUTOEXTEND ON NEXT " & IIf(TbsSize \ 10 = 0, 1, TbsSize \ 10) & "M", "") & _
            " EXTENT MANAGEMENT LOCAL " & _
                IIf(AutoAllocate And Not Temp, " AUTOALLOCATE", " UNIFORM SIZE " & IIf(ExtentSize = 0, "1", ExtentSize) & "M") & _
                IIf(Nologging And Not Temp, " Nologging", "")
            
    err = 0
    On Error Resume Next
    gcnOracle.Execute gstrSQL
    DoEvents
    If err = 0 Then
        CreateTbs = 1
    ElseIf gcnOracle.Errors.Count > 0 Then
        'ORA-01543: ��ռ�'XXX'�Ѿ�����
        If UCase(gcnOracle.Errors(0).Description) Like "ORA-01543: *'ZLMSGDATA'*" Then
            err.Clear
            CreateTbs = 1
        Else
            If MsgBox("�������������Ƿ�����������" & vbCrLf & vbTab & gcnOracle.Errors(0).Description, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                CreateTbs = 2
            Else
                CreateTbs = 1
            End If
        End If
    Else
        MsgBox "��ռ�" & TbsName & "�޷�������������̴�С�ȡ�", vbExclamation, gstrSysName
        CreateTbs = 2
    End If

End Function

Private Function GrantToRole(ByVal cnThis As ADODB.Connection, ByVal cllRoles As Collection, ByVal strOwnerName As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�����¶���صĽ�ɫ������Ȩ
    '���:cllRoles-��ɫ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-09 16:42:54
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, strRoleName As String
    Dim lngCount As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strOwner() As String
    ReDim strOwner(0)
    strOwner(0) = strOwnerName
    '���ļ�������֮���ٴ����ɫ����Ȩ
    lngCount = cllRoles.Count
    pgbState.value = 0
    If lngCount = 0 Then Exit Function
    SetProgressVisible True
    '����ϵͳ�Ų�ͬ�����ݿ���û����
    Dim lngNewSystem As Long, lngOldSystem  As Long
    lngNewSystem = intDefSysCode * 100 + mlng����
    lngOldSystem = Mid(mlst��׼.Key, 2)
            
    For i = 0 To cllRoles.Count
        strRoleName = cllRoles(i)
        If mbln���� = True Then
            gstrSQL = "insert into zlRoleGrant(ϵͳ,��ɫ,���,����) " & _
                   " select " & lngNewSystem & ",��ɫ,���,���� from zlRoleGrant where ��ɫ='" & strRoleName & "' and ϵͳ=" & lngOldSystem
            cnThis.Execute gstrSQL
        End If
        gstrSQL = "select B.����,B.Ȩ�� from zlrolegrant A,zlprogprivs B " & _
                    " where A.��ɫ='" & strRoleName & "' and B.������='" & UCase(Trim(txtOwnerUsr.Text)) & "' and A.ϵͳ=B.ϵͳ and A.���=B.��� and A.����=B.���� "
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, cnThis, adOpenStatic, adLockReadOnly
        
        Do Until rsTemp.EOF
            gstrSQL = "GRANT " & rsTemp("Ȩ��") & " ON " & rsTemp("����") & " TO " & strRoleName
            cnThis.Execute gstrSQL
            rsTemp.MoveNext
        Loop
        
        Call GrantSpecialToRole(cnThis, strRoleName, False, strOwner, True)
        pgbState.value = Int(pgbState.value / lngCount * 100)
        
    Next
    SetProgressVisible False
    GrantToRole = True
End Function

Private Function CopyManageData(ByVal cnExecuter As ADODB.Connection) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngNewSystem As Long
    Dim lngOldSystem As Long
    Dim strOldOwner As String
    
    pgbState.value = 0
    SetProgressVisible True
    
    lngNewSystem = intDefSysCode * 100 + mlng����
    lngOldSystem = Mid(mlst��׼.Key, 2)
    
    strOldOwner = GetOwnerName(lngOldSystem, gcnOracle)
    On Error GoTo errHandle
    'zlComponent����
    gstrSQL = "insert into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ) " & _
                "select ����,����,���汾,�ΰ汾,���汾," & lngNewSystem & " from zlComponent where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 5
    
    'zlPrograms����
    gstrSQL = "insert into zlPrograms(���,����,˵��,����,ϵͳ) " & _
                "select ���,����,˵��,����," & lngNewSystem & " from zlPrograms where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 20
    
    'zlProgFuncs����
    gstrSQL = "insert into zlProgFuncs(���,����,ϵͳ) " & _
                "select ���,����," & lngNewSystem & " from zlProgFuncs where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 35
    
    'zlProgPrivs����
    gstrSQL = "insert into zlProgPrivs(���,����,������,����,Ȩ��,ϵͳ) " & _
                "select ���,����,decode(������,'" & strOldOwner & "',user,������),����,Ȩ��," & lngNewSystem & " from zlProgPrivs where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 65
    
    'zlMenus����
    '������Ч�˵�
    With rsTemp
        Do
            If .State = adStateOpen Then .Close
            gstrSQL = "select 1 from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            .Open gstrSQL, cnExecuter
            If .EOF Then Exit Do
            gstrSQL = "delete from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            cnExecuter.Execute gstrSQL
        Loop
    End With
    CopyMenu gcnOracle, lngOldSystem, lngNewSystem
    pgbState.value = 85
    
    'zlBaseCode����
    gstrSQL = "insert into zlBaseCode(����,�̶�,˵��,����,ϵͳ) " & _
                "select ����,�̶�,˵��,����," & lngNewSystem & " from zlBaseCode where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 90
    
    'zlbaktables����
    gstrSQL = "Insert Into zltools.zlbaktables (ϵͳ, ����, ���, ���, ֱ��ת��) select " & lngNewSystem & ", ����, ���, ���, ֱ��ת�� from zltools.zlbaktables where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    
    'zlDataMove����
    gstrSQL = "insert into zlDataMove(���,����,˵��,�����ֶ�,ת������,�ϴ�����,ϵͳ,״̬) " & _
                "select ���,����,˵��,�����ֶ�,ת������,�ϴ�����," & lngNewSystem & ",״̬ from zlDataMove where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 95
    
    'zlAutoJobs����
    gstrSQL = "insert into zlAutoJobs(����,���,����,˵��,����,����,ִ��ʱ��,���ʱ��,ϵͳ) " & _
                "select ����,���,����,˵��,����,����,ִ��ʱ��,���ʱ��," & lngNewSystem & " from zlAutoJobs where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 97
    
    'zlParameters����
    gstrSQL = "Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵��) " & _
            " Select zlParameters_ID.Nextval," & lngNewSystem & ",ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 99
    
    pgbState.value = 0
    pgbState.Visible = True
    CopyManageData = True
    Exit Function
errHandle:
    If MsgBox("�������д����Ƿ������" & vbCrLf & "    " & err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
    pgbState.value = 0
    pgbState.Visible = True
    
End Function

Private Sub ChkSequence()
    '----------------------------------------------
    '���ܣ��������еĵ�ǰ����
    '----------------------------------------------
    Dim rsLst As ADODB.Recordset
    
    pgbState.value = 0
    SetProgressVisible True
    
    Set rsLst = GetSequence("", mcnOwner)
    With rsLst
        Do Until .EOF
            DoEvents
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            Call AdjustNameSequece(!Owner & "." & !Table_Name, mcnOwner, !Column_Name)
            .MoveNext
        Loop
    End With
    Call Adjust����ID(mcnOwner)
    
    pgbState.value = 0
    SetProgressVisible False
End Sub

Private Function RunSQLScript(ByVal strFile As String) As Boolean
'���ܣ�ִ��SQL�ű�
'      strFile=SQL�ű���
'���أ�RunSQLScript=�ļ��Ƿ�ִ�гɹ�
    Dim strTmp As String
    Dim strTmpPath As String
    Dim strCaprion As String
    
    With mclsRunScript
        .ProcMode = 0
        pgbState.value = 0
        If .OpenFile(strFile) Then
            Do While Not .EOF
                pgbState.value = .ProcessValue
                Call .CollectRoles
                If .ExecuteSQL = False Then Exit Function
                Call .ReadNextSQL
            Loop
            RunSQLScript = True
        Else
            RunSQLScript = False
        End If
    End With
End Function

