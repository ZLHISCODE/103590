VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillings 
   AutoRedraw      =   -1  'True
   Caption         =   "סԺ���ʱ�"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "frmBillings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11760
   Begin VB.Timer tmrStatuPati 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picStatuPancl 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6690
      ScaleHeight     =   300
      ScaleWidth      =   2340
      TabIndex        =   34
      Top             =   7065
      Width           =   2340
      Begin VB.Label lblStatuPati 
         Caption         =   "����Ƿ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   0
         TabIndex        =   35
         Top             =   45
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6750
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillings.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillings.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   7005
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillings.frx":1A7E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13705
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   88
            Key             =   "�������"
            Object.ToolTipText     =   "�������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "ҽ������"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmBillings.frx":2312
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmBillings.frx":294C
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ListView lvwPati 
      Height          =   2550
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1875
      Visible         =   0   'False
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   4498
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "סԺ��"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "�Ա�"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "��Ժ"
         Object.Width           =   970
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   4290
      Left            =   15
      TabIndex        =   2
      Top             =   1080
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   7567
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Frame fraAppend 
      Height          =   570
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "���:F6"
      Top             =   5280
      Width           =   11895
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   165
         Width           =   1800
      End
      Begin VB.ComboBox cbo������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7110
         TabIndex        =   6
         Top             =   180
         Width           =   1890
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   9840
         TabIndex        =   7
         Top             =   180
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "YYYY-MM-DD HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chk�Ӱ� 
         Caption         =   "�Ӱ�ִ��(&A)"
         Height          =   270
         Left            =   240
         TabIndex        =   3
         Top             =   195
         Width           =   1395
      End
      Begin VB.Label lblBaby 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ӥ����(&B)"
         Height          =   240
         Left            =   1560
         TabIndex        =   4
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   9075
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   6480
         TabIndex        =   21
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11760
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6360
      Width           =   11760
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "5685"
         Text            =   "0.00"
         Top             =   90
         Width           =   1815
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "2760"
         Text            =   "0.00"
         Top             =   90
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10035
         TabIndex        =   13
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   8745
         TabIndex        =   12
         ToolTipText     =   "�ȼ���F2"
         Top             =   165
         Width           =   1100
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "������ʱ�:F3"
         Top             =   105
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   690
         MaxLength       =   8
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   7260
         TabIndex        =   15
         Top             =   165
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   6150
         TabIndex        =   14
         Top             =   165
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2205
         TabIndex        =   30
         Top             =   195
         Width           =   510
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5130
         TabIndex        =   29
         Top             =   195
         Width           =   510
      End
   End
   Begin VB.Frame fraTitle 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   22
      Top             =   420
      Width           =   11910
      Begin VB.PictureBox picUnit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         ScaleHeight     =   375
         ScaleWidth      =   2730
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   150
         Width           =   2730
         Begin VB.ComboBox cbo�������� 
            Height          =   300
            Left            =   840
            TabIndex        =   0
            Text            =   "cbo��������"
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Left            =   60
            TabIndex        =   32
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   1305
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11310
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   165
         Width           =   525
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   11400
         TabIndex        =   27
         Top             =   195
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   9360
         TabIndex        =   26
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Frame fraDrawDept 
      Height          =   645
      Left            =   0
      TabIndex        =   33
      Top             =   5730
      Visible         =   0   'False
      Width           =   13575
      Begin VB.ComboBox cboDrawDept 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   225
         Width           =   3315
      End
      Begin VB.Label lblDrawDrugDept 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ����"
         Height          =   180
         Left            =   495
         TabIndex        =   8
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "סԺ���ʱ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   195
      TabIndex        =   25
      ToolTipText     =   "���:F6"
      Top             =   60
      Width           =   1875
   End
End
Attribute VB_Name = "frmBillings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'����������������������������������������������������������������������������������������������������������������������������������������
'��ڲ�����
'����ʼ״̬����:
Public mbytInState As Byte '0-ִ��,1-���,2-����,3-����
Public mstrInNO As String '��mbytInState=1,2,3ʱ��Ч,���ڵ��ݺ�
Public mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���

Public mstr����IDs As String '����ʱ,����Ĳ���ID��,��Ϊ����
Public mstrTime As String '�����������ݵĵǼ�ʱ��
Public mblnDelete As Boolean '�Ƿ�����˷ѵ���
Public mlngDelRow As Long '���ⲿ��������ʱ��ȱʡ���ʵķ��ü�¼

Public mlngUnitID As Long '��ǰ���ʲ���,Ϊ0ʱ��ʾ���в���
Public mlngDeptID As Long '��ǰ���ʿ���,Ϊ0ʱ��ʾ���п���
Public mbytUseType As Byte '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����
Public mlng����ID As Long '���ҷ�ɢ������
Public mstrPrivs As String
Public mlngModule As Long

Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����

'����������������������������������������������������������������������������������������������������������������������������������������
'���ݶ���
Private mrsWork As New ADODB.Recordset '�����ϰ��ҩ��
Private mblnWork As Boolean '��ǰ�Ƿ��������ϰ��ҩ��
Private mlngҩƷ���ID As Long '��ǰ���ݲ�����ҩƷ������ID
Private mlng�������ID As Long '��ǰ���ݲ���������������ID
''''''''''''''''
Private mrsClass As ADODB.Recordset '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsMedPayMode As ADODB.Recordset '���п��õ�ҽ�Ƹ��ʽ
Private mrsLevel As ADODB.Recordset '��ѡ���˷ѱ�
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsInfo As New ADODB.Recordset '������Ϣ(�����˱�ʶ�����￨�š���ҳID��סԺ�š����š��������Ա����䡢�ѱ�)
Private mrs�������� As ADODB.Recordset  '��ѡ�Ŀ�������
Private mrs������ As ADODB.Recordset    '��ѡҽ���ͻ�ʿ
Private mrs��ҩ���� As ADODB.Recordset

Private mstrUseMoney As String  '��ǰ����ʣ���
Private mstrUnitIDs As String   '��ǰ����Ա�����в���ID
'�������
Private mobjBill As ExpenseBill '������õ��ݶ������
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail '�������շ�ϸĿ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mcolMoneys As BillInComes  '���������Ŀ���ܼ���(��ʾ����ӡʱʹ��)���

Private Enum BillColType       '���ݿؼ���������
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum

Private Enum BillCol
    ���� = 0
    �Ա� = 1
    ���� = 2
    ���� = 3
    �ѱ� = 4
    ��� = 5
    ��Ŀ = 6
    ��Ʒ�� = 7
    ��� = 8
    ��λ = 9
    ���� = 10
    ���� = 11
    ���� = 12
    Ӧ�ս�� = 13
    ʵ�ս�� = 14
    ִ�п��� = 15
    ��־ = 16
    ���� = 17
End Enum

Private Enum Pan
    C2��ʾ��Ϣ = 2
End Enum


'�������
Private mblnSendMateria As Boolean  '���ʺ��Զ���ҩ
Private mstrWarn As String '���˱�����ѡ����������(eg:";����:DEF5;����:�ѱ���;����:567DF;����:G")
Private mrsWarn As ADODB.Recordset  '���в�����������
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀ�ĳ����鷽ʽ

Private mlngPreRow As Long '��ǰ�к�,�����иı�ʱ�ж�
Private mblnEnterCell As Boolean '�����Ƿ�ִ��EnterCell�¼�
Private mcurModiMoney As Currency '�޸ĵĵ�ǰ���ݵ�ǰ���˵Ľ��,��bill_entercell��ȡֵ


Private mblnDrop As Boolean '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnPrint As Boolean '��ȡ��˵�ʱ�Ƿ����Ҫ��ӡ���շ����
Private marrColData() As Integer '��ǰ���ݱ༭����ӳ��
Private marrSerial() As Integer '��¼���ʵ��ݷ����е����

Private mcolPatiInfo As Collection '��¼���ʵ��ݵĲ���ID,��ҳID,Ӥ����

Private mblnOne As Boolean      '�Ƿ�ֻ��һ�������շ����
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����
Private mlngPreUnit As Long '���ü�
Private Const STR_HEAD = "����,750,1;����,300,1;����,300,1;����,450,1;�ѱ�,500,1;���,650,1;��Ŀ,1700,1;��Ʒ��,1800,1;���,950,1;" & "��λ,550,4;��,300,1;����,450,1;����,850,7;Ӧ�ս��,850,7;ʵ�ս��,850,7;ִ�п���,1200,1;��־,450,5;����,450,1"
Private mstrҩƷ�۸�ȼ� As String, mstr���ļ۸�ȼ� As String, mstr��ͨ�۸�ȼ� As String

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    '���˺� ����:27378 ����:2010-01-27 13:35:37
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "ִ�п���"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "��ҩҩ��"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If InStr(",4,5,6,7,", .�շ����) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, j As Long, bytsubs As Integer
    Dim bln��������ۿ� As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Or chkCancel.Value = 1 Then Cancel = True: Exit Sub
    
    If mobjBill.Details.Count >= Row Then
        '��������Ŀ����ɾ��ȷ��
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("����Ŀ���� " & bytsubs & " ��������Ŀ,ɾ������ĿҲ��ɾ�����Ĵ�����Ŀ,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).�������� <> 0 Then '������Ŀɾ��ȷ��
            If MsgBox("����Ŀ��[" & mobjBill.Details(mobjBill.Details(Row).��������).Detail.���� & "]�Ĵ�����Ŀ,ȷ��Ҫɾ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln��������ۿ� = gbln��������ۿ�
            End If
        ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        If bln��������ۿ� Then lngMainRow = mobjBill.Details(Bill.Row).�������� '����Ǵ���,ɾ��֮ǰ���´���Ĵ�������,���������,����ɾ��,��������
        
        'ɾ������
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).�������� = Row Then
                Call DeleteDetail(i) '��˳��ɾ���������
            End If
        Next
        Call DeleteDetail(Row) 'ɾ������
        
        'ɾ��ֻ�����˲��˵�������ʱ,�ò�����Ϣ�������
        For i = 1 To Bill.Rows - 1
            If mobjBill.Details.Count < i Then
                For j = 0 To Bill.Cols - 1
                    Bill.TextMatrix(i, j) = ""
                Next
            End If
        Next
        
        '���¼��㲢ˢ��
        If bln��������ۿ� Then
            If CheckItemHaveSub(lngMainRow) Then
                Call Calc��������ʵ��(lngMainRow)
            Else
                Call CalcMoney(lngMainRow, False) 'ֻ��һ��������,����ȫ����ɾ��ʱ,������ͨ���������
            End If
        End If
        
        Call ShowDetails
        
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '���ÿؼ�������ɾ��
        
        mlngPreRow = 0  '��ʾ�иı���
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        Call SetDrawDrugDeptEnabled
    End If
End Sub

Private Sub Bill_CellCheck(Row As Long, Col As Long)
'˵��������ȫ��Ϊ��Ҫ����,������ȫ��Ϊ��������
    Dim i As Long, strCheck As String, bytTime As Integer
    Dim blnReSet As Boolean
    
    If Bill.TextMatrix(Row, BillCol.�Ա�) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then Exit Sub
    
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "" '������δ��������Ч
        Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    '����:  If mobjBill.Details(i).����ID = mobjBill.Details(Row).����ID Then '���˺�:�ಡ��ʱ,û�м����صĲ���
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).����ID = mobjBill.Details(Row).����ID Then
            If mobjBill.Details(i).�շ���� = "F" And mobjBill.Details(i).���ӱ�־ = 0 And i <> Row Then bytTime = bytTime + 1
        End If
    Next
    
    blnReSet = bytTime > 0
    If blnReSet = False Then     '����ֻ���ڸ����������ָĳ���������,��Ҫ���¼ƴ���:25495
        blnReSet = (strCheck = "" And mobjBill.Details(Row).�շ���� = "F" And mobjBill.Details(Row).���ӱ�־ = 1)
    End If
    
    If blnReSet Then
        mobjBill.Details(Row).���ӱ�־ = IIf(strCheck = "", 0, 1)
        Call CalcMoneys(Row)
        Call ShowDetails(Row)
        CalcOneTotal (Bill.Row)
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "�����б�Ȼ��һ���������Ǹ���������", vbInformation, gstrSysName
    End If
End Sub

Private Sub Bill_CommandClick()
    Dim lng��Ŀid As Long, blnCancel As Boolean, bln��ʿ As Boolean, int���� As Integer
    Dim str��� As String, str��׼��Ŀ As String
    Dim int�������� As Integer, int������Դ As Integer
    Dim lng����ID As Long
    
    Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
    If gbln�շ���� Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
        End If
    Else
        str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
    End If
    
    'ҽ��������׼��Ŀ
    If mobjBill.Details.Count >= Bill.Row Then
        If Val(mobjBill.Details(Bill.Row).��ҩ����) > 0 Then '����
            int���� = Val(mobjBill.Details(Bill.Row).��ҩ����)
            '���˺�:24862
            If zl_Check��׼��Ŀ(gclsInsure, int����, mobjBill.Details(Bill.Row).����ID, False) Then str��׼��Ŀ = Get������׼��Ŀ(mobjBill.Details(Bill.Row).����ID, "A.ID")
        End If
        lng����ID = mobjBill.Details(Bill.Row).����ID
    ElseIf mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!����) Then
            int���� = mrsInfo!����
            '���˺�:24862
            If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
        End If
         lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    int�������� = -2
    If mobjBill.Details.Count >= Bill.Row Then
        int�������� = mobjBill.Details(Bill.Row).��������
    ElseIf mrsInfo.State = 1 Then
        int�������� = mrsInfo!��������
    End If
    If int�������� <> -2 Then
        If int�������� = 0 Or int�������� = 2 Then
            int������Դ = 2
        ElseIf int�������� = 1 Or int�������� = -1 Then
            int������Դ = 1
        End If
    Else
        int������Դ = 2
    End If
    
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, gblnסԺ��λ, str���, , , str��׼��Ŀ, _
        zl��ȡ��ҩ��̬(lng����ID, Bill.Row), , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If lng��Ŀid <> 0 Then
        Bill.Text = lng��Ŀid
        mblnSelect = True
        Call Bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub Bill_EditKeyPress(KeyAscii As Integer)
    '��һλ����������ĸ,����λ����,���ֵ�ascii�Ǹ���
     If Bill.TextMatrix(0, Bill.Col) = "����" Then
        If KeyAscii <> 13 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Not (Bill.Text = "" Or Bill.SelLength = Len(Bill.Text)) And _
                InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 And KeyAscii > 0 Then
                '����ô���,�򲻽�����صļ��
                '53113
                If Left(Bill.Text, 1) = "/" Then Exit Sub
                KeyAscii = 0: Beep: Exit Sub
            End If
        Else
            If Bill.Active And Bill.ColData(Bill.Col) <> BillColType.Text_UnModify Then
                
                 If cbo��������.ListIndex <> -1 And Bill.Text = "" Then
                    KeyAscii = 0
                    Call FillPatient(cbo��������.ItemData(cbo��������.ListIndex))
                    If Bill.Top + Bill.CellTop + lvwPati.Height > sta.Top Then
                        lvwPati.Top = Bill.Top + Bill.CellTop - lvwPati.Height - 30
                    Else
                        lvwPati.Top = Bill.Top + Bill.CellTop + Bill.RowHeight(1) - 15
                    End If
                        
                    lvwPati.Visible = True
                    lvwPati.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub ShowStock(strҩƷ As String, dbl��� As Double)
'���ܣ���ʾҩƷ�����ĵĿ��
    If InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0 Then
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]���ÿ��:" & dbl���
    Else
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]" & IIf(dbl��� > 0, "��", "��") & "���."
    End If
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'���ܣ�����������
    Dim lng����ID As Long, lng��ҳID As Long, dblStock As Double, i As Long, lngCur����ID As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim curTotal As Currency, blnCopy As Boolean, bln��ʿ As Boolean
    Dim dblPreTime As Double, dblPreMoney As Double, dblNum As Double, dblTemp As Double
    Dim lngDoUnit As Long, strScope As String, curModi As Currency
    Dim strҽ�Ƹ��� As String, blnSkip As Boolean, blnInput As Boolean
    Dim cur��� As Currency, cur���ն� As Currency, curItemMoney As Currency
    Dim rsҩƷ��Ϣ As ADODB.Recordset, int���� As Integer
    Dim lng���˿���ID As Long, lng���˲���ID As Long, int������Դ As Integer
    Dim str��׼��Ŀ As String, int�������� As Integer, lngOld���� As Long
    Dim lng��Ŀid As Long, str��� As String, strҩ��IDs As String, strժҪ As String
    Dim bln�������� As Boolean, strPriceGrade As String
    Dim colStock As Collection
    
    On Error GoTo errH
    
    If KeyCode = 13 Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.Cols - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "ִ�п���" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = BillColType.Text_UnModify Then Exit Sub
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "����" '��ȡ������Ϣ
                If Bill.Text <> "" Then
                                        
                    If mobjBill.Details.Count >= Bill.Row Then
                        If Bill.Text = mobjBill.Details(Bill.Row).���� Then Exit Sub '˫���Ѷ�ȡ�Ĳ���,δ�ı�ʱ���ض�
                    End If
                    Dim blnMsgbox As Boolean
                    If Not GetPatient(Bill.Text, IsNumeric(Bill.Text) And IsNumeric(Left(Bill.Text, 1)), blnMsgbox) Then
                        If Not blnMsgbox Then
                            MsgBox "����ı�ʶ���ܶ�ȡ������Ϣ�����������Ƿ���ȷ��", vbExclamation, gstrSysName
                        End If
                        If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                        Call Beep: Cancel = True: Exit Sub
                    Else
                        
                        '����ʣ���(���þ��￨����)
                        curModi = 0
                        If mstrInNO <> "" And gbytBilling = 0 Then
                            curModi = GetBillMoney(2, mstrInNO, mrsInfo!����ID)
                        End If
                        '���˵��շ��ö�
                        cur���ն� = mrsInfo!���ն� - curModi
                        
                        Set rsTmp = GetMoneyInfo(mrsInfo!����ID, CDbl(curModi), True, 2)
                    
                        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                        '���˺�:26952
                        cur��� = 0
                        If Not rsTmp Is Nothing Then
                            If rsTmp.State = 1 Then
                                If rsTmp.EOF = False Then
                                    cur��� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
                                End If
                            End If
                        End If
                        If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(1, Val(Nvl(mrsInfo!����ID))) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                        gbytWarn = BillingWarn(mstrPrivsOpt, Trim(Nvl(mrsInfo!����)) & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), GetPatiWarnRange(Val(Nvl(mrsInfo!����ID)), IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID)), _
                             mrsWarn, cur���, cur���ն�, 0, IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), "", "", mstrWarn, True, , 0, True)
                        '����:0;û�б���,����
                        '     1:������ʾ���û�ѡ�����
                        '     2:������ʾ���û�ѡ���ж�
                        '     3:������ʾ�����ж�
                        '     4:ǿ�Ƽ��ʱ���,����
                        '     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
                        If gbytWarn = 2 Or gbytWarn = 3 Then
                            If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                            Call Beep: Cancel = True: Exit Sub
                        End If
                        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
                        
                        '��һ�ж�λסԺҽʦ
                        If Bill.Row = 1 And cbo��������.ListIndex <> -1 Then Call cbo��������_Click
                        
                        Bill.Text = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        Bill.TextMatrix(Bill.Row, BillCol.����) = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        Bill.TextMatrix(Bill.Row, BillCol.�Ա�) = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
                        Bill.TextMatrix(Bill.Row, BillCol.����) = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        Bill.TextMatrix(Bill.Row, BillCol.����) = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        'ȱʡ�ѱ�
                        Bill.TextMatrix(Bill.Row, BillCol.�ѱ�) = IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�)
                        Bill.MsfObj.CellForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
                        Dim dblԤ����� As Double, dblFee As Double, dblʣ�� As Double
                        If Not rsTmp Is Nothing Then
                            curTotal = CalcOneTotal(Bill.Row)
                            'sta.Panels(3).Text = mrsInfo!���� & "Ԥ��:" & Format(rsTmp!Ԥ�����, "0.00")
                            'sta.Panels(3).Text = sta.Panels(3) & "/����:" & Format(rsTmp!�������, gstrDec)
                            'sta.Panels(3).Text = sta.Panels(3) & "/���:" & Format(rsTmp!Ԥ����� - rsTmp!�������, "0.00")
                            dblԤ����� = Val(Nvl(rsTmp!Ԥ�����)): dblFee = Val(Nvl(rsTmp!�������))
                            dblʣ�� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
                            mstrUseMoney = dblԤ����� & "," & dblFee & "," & dblʣ��
                        Else
                            'sta.Panels(3).Text = mrsInfo!���� & "Ԥ��:0.00/����:" & gstrDec & "/���:0.00"
                            mstrUseMoney = "0,0,0": dblԤ����� = 0: dblFee = 0: dblʣ�� = 0
                        End If
                        
                        strInfo = GetPatientDue(Val(mrsInfo!����ID))
                        'If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/Ӧ�տ�:" & Format(strInfo, "0.00")
                        Call SetStatuPatiInfor(Nvl(mrsInfo!����), dblԤ�����, dblFee, dblʣ��, Val(strInfo))
                                                
                        Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
                        
                        mstrUseMoney = mstrUseMoney & "," & cur���ն�
                        
                        
                        If mobjBill.Details.Count >= Bill.Row Then
                            mlngPreRow = 0  '�޸�������ʱ,�ָ���ֵ,�Ա���ʾ���
                            '�޸Ĳ�����Ϣ
                            With mobjBill.Details(Bill.Row)
                                .�������� = IIf(IsNull(mrsInfo!��������), 0, mrsInfo!��������)
                                .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                                .��ҳID = IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID)
                                .Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
                                
                                .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                                .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                                
                                .������ = IIf(IsNull(mrsInfo!������), 0, mrsInfo!������)
                                
                                .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                                .סԺ�� = IIf(IsNull(mrsInfo!סԺ��), 0, mrsInfo!סԺ��)
                                .���� = Bill.TextMatrix(Bill.Row, BillCol.����)
                                .�Ա� = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
                                .���� = Bill.TextMatrix(Bill.Row, BillCol.����)
                                
                                '�����ʱ���,����ʱ,��ҩ�������ڼ�¼�ò��˵�����
                                .��ҩ���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                                
                                .�ѱ� = zlStr.NeedName(Bill.TextMatrix(Bill.Row, BillCol.�ѱ�))
                                .ҽ�Ƹ��� = IIf(IsNull(mrsInfo!ҽ�Ƹ��ʽ), "", mrsInfo!ҽ�Ƹ��ʽ)
                                
                                .���￨�� = mstrUseMoney
                                
                                '����ж�Ӧ�Ĵ�����Ŀ,�������Ŀ�Ĳ�����ϢҲ����
                                For i = Bill.Row + 1 To mobjBill.Details.Count
                                    If mobjBill.Details(i).�������� = Bill.Row Then
                                        mobjBill.Details(i).�������� = .��������
                                        mobjBill.Details(i).����ID = .����ID
                                        mobjBill.Details(i).��ҳID = .��ҳID
                                        mobjBill.Details(i).����ID = .����ID
                                        mobjBill.Details(i).����ID = .����ID
                                        mobjBill.Details(i).������ = .������
                                        mobjBill.Details(i).���� = .����
                                        mobjBill.Details(i).סԺ�� = .סԺ��
                                        mobjBill.Details(i).���� = .����
                                        mobjBill.Details(i).�Ա� = .�Ա�
                                        mobjBill.Details(i).���� = .����
                                        mobjBill.Details(i).��ҩ���� = .��ҩ����
                                        mobjBill.Details(i).�ѱ� = .�ѱ�
                                        mobjBill.Details(i).ҽ�Ƹ��� = .ҽ�Ƹ���
                                        mobjBill.Details(i).���￨�� = .���￨��
                                        mobjBill.Details(i).Ӥ���� = .Ӥ����
                                    End If
                                Next
                            End With
                        End If
                        
                        If Not IsNull(mrsInfo!��Ժ����) Then
                            MsgBox "��������" & vbCrLf & vbCrLf & "�ò������� " & Format(mrsInfo!��Ժ����, "yyyy-MM-dd") & " ��Ժ�����ڶԸò���ǿ�ƽ��м��ʣ�", vbInformation, gstrSysName
                            If mrsInfo!��Ժ���� < CDate(txtDate.Text) Then
                                txtDate.Text = Format(mrsInfo!��Ժ����, "yyyy-MM-dd HH:mm:ss")
                            End If
                        Else
                            txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                        End If
                        If Not IsNull(mrsInfo!��Ժ����) Then
                            sta.Panels(2).Text = "��Ժ����:" & Format(mrsInfo!��Ժ����, "yyyy-MM-dd")
                            strInfo = GetInsureInfo(mrsInfo!����ID)
                            If strInfo <> "" Then sta.Panels(2).Text = sta.Panels(2).Text & "/�ʺ�:" & Split(strInfo, ";")(1)
                        End If
                        
                        If mobjBill.Details.Count = 0 Then txt����.Text = gstrDec
                        
                        '�����ĩ����δ������õĲ�����ɾ������(δȷ����)
                        If mobjBill.Details.Count = Bill.Rows - 2 And Bill.Row = Bill.Rows - 2 Then
                            Bill.RemoveMSFItem Bill.Rows - 1
                        End If
                        
                        '���˱��ˣ����°��ѱ���
                        If mobjBill.Details.Count >= Bill.Row Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            '������ĿҲ��Ҫ���¼����ˢ����ʾ
                            For i = Bill.Row + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).�������� = Bill.Row Then
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            Next
                        End If
                    End If
                End If
                
                '�Զ�������һ�еķ�����Ŀ(����ʱ���˲�ͬ)
                If Bill.Row > 1 And mobjBill.Details.Count < Bill.Row Then
                    If mrsInfo.State = 1 Then 'ҽ�����˲�����
                        If IsNull(mrsInfo!����) Then
                            '�����ʱ�ۻ����ҩƷ,���Զ�����(��ֹ�ظ�)
                            If mobjBill.Details(Bill.Row - 1).���� <> mrsInfo!���� And mobjBill.Details(Bill.Row - 1).�������� = 0 _
                                And Not (mobjBill.Details(Bill.Row - 1).Detail.��� _
                                    Or mobjBill.Details(Bill.Row - 1).Detail.����) And Not (mobjBill.Details(Bill.Row - 1).�շ���� = "F" And mobjBill.Details(Bill.Row - 1).���ӱ�־ = 1) Then
                                
                                '���˺�:���븽������������: Not (mobjBill.Details(Bill.Row - 1).�շ���� = "F" And mobjBill.Details(Bill.Row - 1).���ӱ�־ = 1)
                                '����:
                                blnCopy = True '��־Ҫ���Ʒ�����
                                
                                With mobjBill.Details(Bill.Row - 1)
                                    mobjBill.Details.Add .Detail, .�շ�ϸĿID, .��� + 1, .��������, IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID), _
                                    IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID), IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID), _
                                    IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID), IIf(IsNull(mrsInfo!����), "", mrsInfo!����), _
                                    IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�), IIf(IsNull(mrsInfo!����), "", mrsInfo!����), _
                                    IIf(IsNull(mrsInfo!סԺ��), 0, mrsInfo!סԺ��), IIf(IsNull(mrsInfo!����), "", mrsInfo!����), _
                                    IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�), IIf(IsNull(mrsInfo!��������), "", mrsInfo!��������), _
                                    .�շ����, .���㵥λ, "", .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, mstrUseMoney, , _
                                    IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), IIf(IsNull(mrsInfo!ҽ�Ƹ��ʽ), "", mrsInfo!ҽ�Ƹ��ʽ)
                                    
                                End With
                                Call CalcMoneys(Bill.Row)
                                
                                Call ShowDetails(Bill.Row)
                                Set mrsInfo = New ADODB.Recordset: mstrUseMoney = "" '��ǰ������Ϣ��Ч(���ⱻ����������)
                                Bill.Col = IIf(gbln�շ����, BillCol.���, BillCol.��Ŀ)   '��λ����Ŀ������
                                Bill.Text = "": Bill.SetFocus
                                Cancel = True
                            End If
                        End If
                    End If
                    
                    If Bill.ColData(BillCol.���) <> BillColType.UnFocus And Me.Visible And _
                        Bill.TextMatrix(Bill.Row - 1, BillCol.����) = Bill.TextMatrix(Bill.Row, BillCol.����) Then _
                        Call zlCommFun.PressKey(13)
                End If
            Case "���"
                If Bill.ListIndex <> -1 Then '���������ʱ���ᶨλ�������
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        'һ���ĸ��շ����,�����(����)ԭ�и���Ŀ����
                        For i = 5 To Bill.Cols - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Details.Count >= Bill.Row Then
                            Set mobjBill.Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Details(Bill.Row)
                                .�շ�ϸĿID = 0: .�շ���� = ""
                            End With
                            Call CalcMoneys
                        End If
                    End If
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '��ʱ��RowData��¼��ѡ����շ����
                End If
            Case "��Ŀ"
                If blnCopy Then Bill.Text = mobjBill.Details(mobjBill.Details.Count).Detail.ID
                
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    '��������������Ŀ�ϰ��س�,��ѡ����ѡ��
                    If mobjBill.Details.Count >= Bill.Row Then
                        'ͨ����ťѡ���Ƿ��ص�ID,�����������ı�,�����һ����,�򲻸ı�
                        If Bill.TextMatrix(Bill.Row, BillCol.��Ŀ) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                        
                    'ҽ��������׼��Ŀ
                    If mobjBill.Details.Count >= Bill.Row Then  '�޸��Ѽ���ķ���
                        If Val(mobjBill.Details(Bill.Row).��ҩ����) > 0 Then '����
                            int���� = Val(mobjBill.Details(Bill.Row).��ҩ����)
                            '���˺�:24862
                            If zl_Check��׼��Ŀ(gclsInsure, int����, mobjBill.Details(Bill.Row).����ID, False) Then str��׼��Ŀ = Get������׼��Ŀ(mobjBill.Details(Bill.Row).����ID, "A.ID")
                            
                        End If
                        lngCur����ID = mobjBill.Details(Bill.Row).����ID
                    ElseIf mrsInfo.State = 1 Then   '�ò��˵�һ������
                        If Not IsNull(mrsInfo!����) Then
                            int���� = mrsInfo!����
                            '���˺�:24862
                            If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
                        End If
                        lngCur����ID = Val(Nvl(mrsInfo!����ID))
                    ElseIf Bill.TextMatrix(Bill.Row, BillCol.����) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then  'ͬһ���˶�������
                        If Val(mobjBill.Details(Bill.Row - 1).��ҩ����) > 0 Then '����
                            int���� = Val(mobjBill.Details(Bill.Row - 1).��ҩ����)
                            '���˺�:24862
                            If zl_Check��׼��Ŀ(gclsInsure, int����, mobjBill.Details(Bill.Row - 1).����ID, False) Then str��׼��Ŀ = Get������׼��Ŀ(mobjBill.Details(Bill.Row - 1).����ID, "A.ID")
                        End If
                        lngCur����ID = mobjBill.Details(Bill.Row - 1).����ID
                    End If
                
                    sta.Panels(2).Text = ""
                    sta.Panels("MedicareType").Text = ""
                    blnInput = True
                    If mblnSelect Or blnCopy Then
                        mblnSelect = False: blnCopy = False '���������־
                        Set mobjDetail = GetInputDetail(Val(Bill.Text), int����)
                    Else
                        If gbln�շ���� Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(2) = "û��ȷ���������,�����������"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
                            str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
                        End If
                        
                        int�������� = -2
                        If mobjBill.Details.Count >= Bill.Row Then
                            int�������� = mobjBill.Details(Bill.Row).��������
                        ElseIf mrsInfo.State = 1 Then
                            int�������� = mrsInfo!��������
                        End If
                        If int�������� <> -2 Then
                            If int�������� = 0 Or int�������� = 2 Then
                                int������Դ = 2
                            ElseIf int�������� = 1 Or int�������� = -1 Then
                                int������Դ = 1
                            End If
                        Else
                            int������Դ = 2
                        End If
                        
                        lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, gblnסԺ��λ, str���, _
                            Bill.Text, Bill.TxtHwnd, str��׼��Ŀ, zl��ȡ��ҩ��̬(lngCur����ID, Bill.Row), _
                            , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                        If lng��Ŀid <> 0 Then
                            Set mobjDetail = GetInputDetail(lng��Ŀid, int����)
                            
                            If int���� <> 0 Then sta.Panels("MedicareType").Text = Getҽ������(lng��Ŀid, int����)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If mrsInfo.State = 0 And Bill.TextMatrix(Bill.Row, BillCol.����) = "" Then
                        sta.Panels(2) = "û��ȷ��������Ϣ,���в��ܼ������룡"
                        Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                    
                    Bill.TxtVisible = False '(���Ӳ���)
                    
                    '�շ��뷢ҩ����ʱ����������ʱ�ۼ�����ҩƷ
                    If InStr(",5,6,7,", mobjDetail.���) > 0 And gbln���뷢ҩ Then
                        If mobjDetail.��� Or mobjDetail.���� Then
                            MsgBox "��ҩ���봦��ʱ��������ʱ�ۻ����ҩƷ��", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '��鶾�����ͼ�ֵ����Ȩ��
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        Set rsҩƷ��Ϣ = ReadҩƷ��Ϣ(mobjDetail.ID)
                        If Not rsҩƷ��Ϣ Is Nothing Then
                            If IIf(IsNull(rsҩƷ��Ϣ!�������), "", rsҩƷ��Ϣ!�������) = "����ҩ" _
                                And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
                                MsgBox """" & mobjDetail.���� & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf IIf(IsNull(rsҩƷ��Ϣ!�������), "", rsҩƷ��Ϣ!�������) = "����ҩ" _
                                And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
                                MsgBox """" & mobjDetail.���� & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf (IIf(IsNull(rsҩƷ��Ϣ!��ֵ����), "", rsҩƷ��Ϣ!��ֵ����) = "����" _
                                Or IIf(IsNull(rsҩƷ��Ϣ!��ֵ����), "", rsҩƷ��Ϣ!��ֵ����) = "����") _
                                And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
                                MsgBox """" & mobjDetail.���� & """Ϊ���ػ򰺹�ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '����ID,����
                    lng����ID = 0: int���� = 0
                    If mrsInfo.State = 1 Then
                        lng����ID = Nvl(mrsInfo!����ID, 0)
                        lng��ҳID = Nvl(mrsInfo!��ҳID, 0)
                        int���� = Nvl(mrsInfo!����, 0)
                        lng���˿���ID = Nvl(mrsInfo!����ID, 0)
                        lng���˲���ID = Nvl(mrsInfo!����ID, 0)
                        strҽ�Ƹ��� = "" & mrsInfo!ҽ�Ƹ��ʽ
                    ElseIf Bill.TextMatrix(Bill.Row, 0) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then
                        lng����ID = mobjBill.Details(Bill.Row - 1).����ID
                        lng��ҳID = mobjBill.Details(Bill.Row - 1).��ҳID
                        int���� = Val(mobjBill.Details(Bill.Row - 1).��ҩ����)
                        lng���˿���ID = mobjBill.Details(Bill.Row - 1).����ID
                        lng���˲���ID = mobjBill.Details(Bill.Row - 1).����ID
                        strҽ�Ƹ��� = mobjBill.Details(Bill.Row - 1).ҽ�Ƹ���
                    Else
                        lng����ID = mobjBill.Details(Bill.Row).����ID
                        lng��ҳID = mobjBill.Details(Bill.Row).��ҳID
                        int���� = Val(mobjBill.Details(Bill.Row).��ҩ����)
                        lng���˿���ID = mobjBill.Details(Bill.Row).����ID
                        lng���˲���ID = mobjBill.Details(Bill.Row).����ID
                        strҽ�Ƹ��� = mobjBill.Details(Bill.Row).ҽ�Ƹ���
                    End If
                                        
                    '�������ò��˲�������
                    If InStr(",5,6,7,", mobjDetail.���) = 0 Then
                        If Not CheckFeeItemLimitDept(mobjDetail.ID, lng���˲���ID, lng���˿���ID) Then
                            MsgBox "���շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '���ҩƷ�����Ƿ��ظ�:������ʱ��ͬһҩ���������ظ�(����ֻ����)
                    If InStr(",5,6,7,", mobjDetail.���) > 0 _
                        Or (mobjDetail.��� = "4" And mobjDetail.��������) Then
                        If PhysicExist(mobjDetail, Bill.Row, lng����ID) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '��鴦��ְ��
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        mobjDetail.����ְ�� = Get����ְ��(mobjDetail.ID)
                        'ҽ���򹫷Ѳ��˼��
                        If strҽ�Ƹ��� <> "" Then
                            If CheckDuty(mobjDetail, False, strҽ�Ƹ���) > 0 Then
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                        '���в�����Ŀ���
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    'ҽ�����˷�������,ÿ�в��˿��ܲ�ͬ�����ҿ����޸����������У�����ÿ�ζ�ȡ����������Ŀ
                    If int���� > 0 And mobjDetail.Ҫ������ Then
                        Set rsTmp = GetAuditRecord(lng����ID, lng��ҳID, mobjDetail.ID)
                        If rsTmp.RecordCount = 0 Then
                            MsgBox "��ǰ����δ����׼ʹ��[" & mobjDetail.���� & "]��", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        ElseIf Not IsNull(rsTmp!��������) Then
                            If rsTmp!�������� <= 0 Then
                                MsgBox "��ǰ����ʹ��[" & mobjDetail.���� & "]�Ѵﵽ��׼��ʹ������" & FormatEx(rsTmp!ʹ������ / IIf(gblnסԺ��λ, mobjDetail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                                        
                    
                    '��ȡҩƷ�����Ϣ
                     '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
                    If mobjDetail.��� = "4" Then
                        lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, lng���˲���ID)
                        If lngDoUnit = 0 Then lngDoUnit = Get��������ID
                    End If
                    
                    '���˿���ID
                    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    
                    int������Դ = Get������Դ(Bill.Row)
                    lngDoUnit = Get�շ�ִ�п���ID(mobjDetail.���, mobjDetail.ID, _
                        mobjDetail.ִ�п���, lng���˿���ID, Get��������ID, int������Դ, lngDoUnit, lng���˲���ID)
                        
                    '��ȡҩƷ���
                    If ReadDrugAndStuffStock(lngDoUnit, mobjDetail) = False Then
                        Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
             
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        '��������
                        mobjDetail.�������� = Get��������(mobjDetail.ID)
                    End If
                    
                   '������Ŀ��Ӧ���
                    If int���� > 0 Then
                        If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                            strPriceGrade = mstrҩƷ�۸�ȼ�
                        ElseIf mobjDetail.��� = "4" Then
                            strPriceGrade = mstr���ļ۸�ȼ�
                        Else
                            strPriceGrade = mstr��ͨ�۸�ȼ�
                        End If
                        If Not CheckMediCareItem(mobjDetail.ID, int����, mobjDetail.����, mobjDetail.��� = False, , strPriceGrade) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '����ժҪ(ȡ���е����Ա��޸�)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            strժҪ = mobjBill.Details(Bill.Row).ժҪ
                        End If
                    End If
                    
                    If mrsInfo.State = 1 Then
                        '������޸ĸ��շ�ϸĿ��
                        Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                        '���ò�����Ϣ
                        With mobjBill.Details(Bill.Row)
                            .�������� = Nvl(mrsInfo!��������, 0)
                            .����ID = Nvl(mrsInfo!����ID, 0)
                            .��ҳID = Nvl(mrsInfo!��ҳID, 0)
                            
                            .����ID = Nvl(mrsInfo!����ID, 0)
                            .����ID = Nvl(mrsInfo!����ID, 0)
                            
                            .������ = Nvl(mrsInfo!������, 0)
                            
                            .���� = "" & mrsInfo!����
                            .סԺ�� = Nvl(mrsInfo!סԺ��, 0)
                            .���� = Bill.TextMatrix(Bill.Row, BillCol.����)
                            .�Ա� = Nvl(mrsInfo!�Ա�)
                            .���� = Bill.TextMatrix(Bill.Row, BillCol.����)
                            .�ѱ� = zlStr.NeedName(Bill.TextMatrix(Bill.Row, BillCol.�ѱ�))
                            .ҽ�Ƹ��� = Nvl(mrsInfo!ҽ�Ƹ��ʽ)
                            
                            '�����ʱ���,����ʱ,��ҩ�������ڼ�¼�ò��˵�����
                            .��ҩ���� = Nvl(mrsInfo!����)
                            
                            .���￨�� = mstrUseMoney
                        End With
                    ElseIf Bill.TextMatrix(Bill.Row, BillCol.����) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then
                        '������޸ĸ��շ�ϸĿ��
                        Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                        '���ò�����Ϣ
                        With mobjBill.Details(Bill.Row)
                            .�������� = mobjBill.Details(Bill.Row - 1).��������
                            .����ID = mobjBill.Details(Bill.Row - 1).����ID
                            .��ҳID = mobjBill.Details(Bill.Row - 1).��ҳID
                            
                            .����ID = mobjBill.Details(Bill.Row - 1).����ID
                            .����ID = mobjBill.Details(Bill.Row - 1).����ID
                            
                            .������ = mobjBill.Details(Bill.Row - 1).������
                            
                            .���� = mobjBill.Details(Bill.Row - 1).����
                            .סԺ�� = mobjBill.Details(Bill.Row - 1).סԺ��
                            .���� = mobjBill.Details(Bill.Row - 1).����
                            .�Ա� = mobjBill.Details(Bill.Row - 1).�Ա�
                            .���� = mobjBill.Details(Bill.Row - 1).����
                            .�ѱ� = Mid(Bill.TextMatrix(Bill.Row, BillCol.�ѱ�), InStr(Bill.TextMatrix(Bill.Row, BillCol.�ѱ�), "-") + 1)
                            .ҽ�Ƹ��� = mobjBill.Details(Bill.Row - 1).ҽ�Ƹ���
                            
                            '�����ʱ���,����ʱ,��ҩ�������ڼ�¼�ò��˵�����
                            .��ҩ���� = mobjBill.Details(Bill.Row - 1).��ҩ����
                            
                            .���￨�� = mobjBill.Details(Bill.Row - 1).���￨��
                        End With
                    Else
                        '������޸ĸ��շ�ϸĿ��,���޸ķ�����Ŀ
                        Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    End If
                    '59051
                    '����ժҪ(������������и���ժҪ)
                    If mobjBill.Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Details(Bill.Row).ժҪ = strժҪ
                        End If
                    Else
                        strժҪ = gclsInsure.GetItemInfo(int����, mobjBill.Details(Bill.Row).����ID, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 2)
                        mobjBill.Details(Bill.Row).ժҪ = strժҪ
                    End If
                    
                    Call CalcMoneys(Bill.Row)
                    
                    '�����ҽ��Calcmoney�п��ܷ���ժҪ
                    If mobjBill.Details(Bill.Row).ժҪ <> "" Then strժҪ = mobjBill.Details(Bill.Row).ժҪ
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    mrsWarn.Filter = ""
                    If mrsWarn.RecordCount > 0 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = CalcOneTotal(Bill.Row)
                        If curTotal > 0 Then
                            If mobjBill.Details(Bill.Row).���￨�� = "" Then
                                cur��� = 0
                                cur���ն� = 0
                            Else
                                cur��� = Val(Split(mobjBill.Details(Bill.Row).���￨��, ",")(2))
                                cur���ն� = Val(Split(mobjBill.Details(Bill.Row).���￨��, ",")(3))
                            End If
                            
                            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(1, mobjBill.Details(Bill.Row).����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                            '���˺�:24491
                            curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                            
                            gbytWarn = BillingWarn(mstrPrivsOpt, mobjBill.Details(Bill.Row).���� & IIf(mobjBill.Details(Bill.Row).סԺ�� = "", "", "(סԺ��:" & mobjBill.Details(Bill.Row).סԺ�� & " ����:" & mobjBill.Details(Bill.Row).���� & ")"), lng���˲���ID, GetPatiWarnRange(mobjBill.Details(Bill.Row).����ID, mobjBill.Details(Bill.Row).��ҳID), _
                                mrsWarn, cur���, cur���ն�, curTotal, mobjBill.Details(Bill.Row).������, mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, True, , curItemMoney)
                            
                            If gbytWarn = 2 Or gbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row
                                For i = 0 To Bill.Cols - 1
                                    Bill.TextMatrix(Bill.Row, i) = ""
                                Next
                                Bill.Text = "": Cancel = True
                                Bill.Col = BillCol.����: Bill.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If int���� <> 0 And mobjBill.Details(Bill.Row).���� <> 0 Then
                        If gclsInsure.GetCapability(supportʵʱ���, mobjBill.Details(Bill.Row).����ID, int����) Then
                            If gclsInsure.CheckItem(int����, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                            mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '�������ͼ��
                    Call Check��������(Bill.Row)

                    
                    Set mrsInfo = New ADODB.Recordset: mstrUseMoney = ""
                    '��ǰ������Ϣ��Ч(���ⱻ����������)
                    
                    Bill.Text = "": Bill.SetFocus
                ElseIf mobjBill.Details.Count < Bill.Row Then
                    Call zlCommFun.PressKey(vbKeyTab): Exit Sub
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    With mobjBill.Details(Bill.Row)
                        '��һ�е�����ȷ��
                        If .�շ���� = "7" And gblnPay Then Bill.ColData(BillCol.����) = BillColType.Text   '����
                        If .�շ���� = "F" Then Bill.ColData(BillCol.��־) = BillColType.CheckBox '���ӱ�־
                        
                        '���������������
                        If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                            And Not (.�շ���� = "4" And .Detail.��������) Then
                            Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '����
                            Bill.ColData(BillCol.����) = BillColType.Text '����
                        Else
                            Bill.ColData(BillCol.����) = BillColType.Text '����
                            Bill.ColData(BillCol.����) = BillColType.UnFocus '����
                        End If
                        
                        'ִ�п���
                        If InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ Then
                            Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus: .Key = 1
                        Else
                            '��FillBillComboBox������ListIndexʱ����CboClick�¼�
                            mblnEnterCell = False: Bill.Col = BillCol.ִ�п���: mblnEnterCell = True
                            Call FillBillComboBox(Bill.Row, BillCol.ִ�п���, lng���˿���ID, int������Դ, Not blnInput) 'ֱ�ӻس�ʱ����ִ�п���
                            mblnEnterCell = False: Bill.Col = BillCol.��Ŀ: mblnEnterCell = True
                            
                            blnSkip = Bill.ListCount = 1
                            If Not blnSkip And InStr(",4,5,6,7,", .�շ����) > 0 Then
                                'ָ���˹̶�ҩ��ʱ,��������ѡ��
                                Select Case .�շ����
                                    Case "4"
                                        blnSkip = glng���ϲ��� > 0 And .ִ�в���ID = glng���ϲ���
                                    Case "5"
                                        blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                    Case "6"
                                        blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                    Case "7"
                                        blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                End Select
                            End If
                            If blnSkip Then
                                Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus: .Key = 1
                            Else
                                Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox: .Key = Bill.ListCount
                            End If
                            If lngDoUnit <> .ִ�в���ID Then
                                '��ȡҩƷ���
                                If ReadDrugAndStuffStock(.ִ�в���ID, mobjBill.Details(Bill.Row).Detail) = False Then
                                    Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                        
                        '������Ŀ����
                        If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" And InStr(",5,6,7,", .�շ����) = 0 Then
                            If (gbln��������ۿ� And mobjBill.Details(Bill.Row).�������� = 0) Or Not gbln��������ۿ� Then  '(����м���,ֻȡһ��)
                                If ShouldDO(Bill.Row) Then
                                   Call SetSubItem(lng���˿���ID, int������Դ)
                                   mlngPreRow = 0 'ͨ���б仯��־������ȷ��������
                                End If
                            End If
                        End If
                    End With
                End If
                
                'ֻ����һ�θ���
                If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Details(i).�շ���� = "7" Then
                                '����ִ�иù��̣�����ᶨλ��һ����Ԫ,�ȶ�λ������,����һ����Ԫ������
                                'ѡ����øù��̣����ú���͸��س������ﲻ���ٻس��������������س���Ч��(�ؼ�ԭ��)��
                                Bill.Col = BillCol.����: Exit For
                            End If
                        Next
                    End If
                End If
                Call SetDrawDrugDeptEnabled
            Case "��"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '��ֵ�Ϸ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                        MsgBox "����Ӧ��Ϊ����������", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                
                    '����ҩ���Ǵ�����Ŀ�ſɸ��ĸ���(������ı�,����Ҳ��)
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then 'And mobjBill.Details(Bill.Row).�������� = 0 Then
                        '������ʱ��ҩƷ�����ֹ����(û�з�����ʱ��ҩƷ�����޸ĸ���������)
                        If mobjBill.Details(Bill.Row).Detail.���� Or mobjBill.Details(Bill.Row).Detail.��� Then
                            If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� > mobjBill.Details(Bill.Row).Detail.��� Then
                                MsgBox """" & mobjBill.Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                              
                        '�������ʱ�ۻ������ҩ���ĸ��������Ƿ��㹻
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" _
                                And (mobjBill.Details(i).Detail.��� Or mobjBill.Details(i).Detail.����) Then
                                If Val(Bill.Text) * mobjBill.Details(i).���� > mobjBill.Details(i).Detail.��� Then
                                    MsgBox """" & mobjBill.Details(i).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                                End If
                            End If
                        Next
                        
                        '���㲢ˢ�¸���
                        lngOld���� = mobjBill.Details(Bill.Row).����
                        mobjBill.Details(Bill.Row).���� = Bill.Text
                        Call CalcMoneys(Bill.Row)
                        
                        
                        int���� = Val(mobjBill.Details(Bill.Row).��ҩ����)
                        If int���� <> 0 And mobjBill.Details(Bill.Row).���� <> 0 Then
                            If gclsInsure.GetCapability(supportʵʱ���, mobjBill.Details(Bill.Row).����ID, int����) Then
                                If gclsInsure.CheckItem(int����, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                    mobjBill.Details(Bill.Row).���� = lngOld����
                                    Call CalcMoneys(Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If mobjBill.Details(Bill.Row).���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                mobjBill.Details(Bill.Row).���� = lngOld����
                                Call CalcMoneys(Bill.Row)
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        CalcOneTotal (Bill.Row)
                        
                        '����������ҩ����,����Ƕ�����,���޸������Ǵ����,����Ǵ���,���޸�ͬһ����Ĵ����.��Ϊ�޶�Ϊ�в�ҩ,������������
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" And mobjBill.Details(i).�������� = mobjBill.Details(Bill.Row).�������� _
                                And mobjBill.Details(i).����ID = mobjBill.Details(Bill.Row).����ID Then
                                If mobjBill.Details(i).�������� = 0 Or (mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).Detail.���д��� = 0) Then     '1��2�̶��Ͱ������Ĳ���
                                    mobjBill.Details(i).���� = Bill.Text
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            End If
                        Next
                    Else
                        sta.Panels(2) = "������Ŀ�ĸ������ܸ��ģ�"
                        Bill.Text = mobjBill.Details(Bill.Row).����: Beep '�ָ�ԭ�и���ֵ
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    With mobjBill.Details(Bill.Row)
                        '��ֵ�Ϸ���
                        If Not IsNumeric(Bill.Text) Then
                          MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                          Bill.Text = .����: Cancel = True: Exit Sub
                        End If
                        If Val(Bill.Text) = 0 Then
                          If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                              Bill.Text = .����: Cancel = True: Exit Sub
                          End If
                        End If
                        'ҩƷ����С��
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                          If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And InStr(mstrPrivsOpt, "ҩƷ����С��") = 0 Then
                              MsgBox "��û��Ȩ������С����", vbInformation, gstrSysName
                              Bill.Text = .����: Cancel = True: Exit Sub
                          End If
                        End If
                        '�������
                        If gcurMaxMoney > 0 Then
                          If CSng(Bill.Text) * .���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                              If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                  Bill.Text = .����: Cancel = True: Exit Sub
                              End If
                          End If
                        End If
                        
                        Bill.Text = FormatEx(Bill.Text, 5)
                        
                        int���� = Val(.��ҩ����)
                        If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                            dblNum = CSng(Bill.Text) * .���� * .Detail.סԺ��װ
                        Else
                            dblNum = CSng(Bill.Text) * .����
                        End If
                        
                        '�����Ϸ��Լ��
                        If CSng(Bill.Text) * .���� < 0 Then
                            'Ȩ��
                            bln�������� = True
                            If InStr(",5,6,", .�շ����) > 0 Then
                                bln�������� = (InStr(mstrPrivsOpt, ";��ҩ��������;") > 0)
                            ElseIf InStr(",7,", .�շ����) > 0 Then
                                bln�������� = (InStr(mstrPrivsOpt, ";��ҩ��������;") > 0)
                            Else
                                bln�������� = (InStr(mstrPrivsOpt, ";���Ƹ�������;") > 0)
                            End If
                            
                            If Not bln�������� Then
                                MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                                Bill.Text = .����: Cancel = True: Exit Sub
                            Else
                                If .Detail.���� Then
                                    MsgBox "����ҩƷ���������븺����", vbInformation, gstrSysName
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                                If int���� <> 0 Then
                                    If Not gclsInsure.GetCapability(support��������, .����ID, int����) Then
                                        MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                                        Bill.Text = .����: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                            
                            '���������������
                            If Not (InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ) Then
                                'dblTemp = zlGetBillOtherRowNumToTal(.����ID, .��ҳID, .�շ�ϸĿID, True, Bill.Row)
                                If Not CheckNegative(.����ID, .��ҳID, .�շ�ϸĿID, .ִ�в���ID, dblNum, .Detail.סԺ��װ, mstrPrivsOpt) Then
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        'ҩƷ�����
                        If (.�շ���� = "4" And .Detail.��������) Or (InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ) Then
                            If .Detail.���� Or .Detail.��� Then
                                '������ʱ��ҩƷ�����ֹ����
                                If .���� * CSng(Bill.Text) > .Detail.��� Then
                                    If .�շ���� = "4" Then
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .ִ�в���ID) <> 0 And Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                                    '����ҩƷ�������
                                    If .���� * CSng(Bill.Text) > .Detail.��� Then
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("""" & .Detail.���� & """�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .����: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox """" & .Detail.���� & """�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                            Bill.Text = .����: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ Then
                            'û��Ȩ��ʱ���̶�����ʾ��ʽ���
                            strҩ��IDs = Decode(.�շ����, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                            If strҩ��IDs <> "" And .���� * CSng(Bill.Text) > .Detail.��� Then
                                If gblnStock Then
                                    MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治����������!", vbInformation, gstrSysName
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                Else
                                    If MsgBox("""" & .Detail.���� & """�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        Bill.Text = .����: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        
                        dblPreTime = .����
                        .���� = Bill.Text
                                            
                        '�����������
                        If Not gbln�������� Then
                            If Not CheckLimit(mobjBill, Bill.Row, gblnסԺ��λ) Then
                                .���� = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        If .Detail.¼������ > 0 And dblNum > .Detail.¼������ Then
                            If MsgBox("��������γ�����¼������" & FormatEx(.Detail.¼������ / IIf(gblnסԺ��λ, .Detail.סԺ��װ, 1), 5) & ",�Ƿ����?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                                .���� = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        '����ʹ������
                        If int���� > 0 And .Detail.Ҫ������ Then
                            Set rsTmp = GetAuditRecord(.����ID, .��ҳID, .�շ�ϸĿID)
                            If rsTmp.RecordCount > 0 Then
                                If Not IsNull(rsTmp!��������) Then
                                    If dblNum > rsTmp!�������� Then
                                        MsgBox "��������γ�������׼��ʹ������" & FormatEx(rsTmp!�������� / IIf(gblnסԺ��λ, .Detail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                                        .���� = dblPreTime: Bill.Text = dblPreTime
                                        Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        
                        
                        '���д������ܸ�������(����Ŀ���θı�,���д���������Ҳ��)
                        If .�������� <> 0 And .Detail.���д��� <> 0 Then
                            sta.Panels(2) = "����Ŀ�ǹ��д�����Ŀ,�����β��ܹ����ġ�"
                            .���� = dblPreTime: Bill.Text = dblPreTime
                            Exit Sub
                        End If
                        
                        
                        Call CalcMoneys(Bill.Row)
                        
                        '����������(���Ѿ�������з��õ�δ��ʾǰ)
                        If MoneyOverFlow(mobjBill) Then
                            MsgBox "�����������µ��ݽ����������ʵ�������", vbInformation, gstrSysName
                            .���� = dblPreTime
                            Bill.Text = ""
                            Call CalcMoneys(Bill.Row)
                            Cancel = True: Bill.TxtVisible = False: Exit Sub
                        End If
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 Then
                            curTotal = CalcOneTotal(Bill.Row)
                            If curTotal > 0 Then
                                If .���￨�� = "" Then
                                    cur��� = 0
                                    cur���ն� = 0
                                Else
                                    cur��� = Val(Split(.���￨��, ",")(2))
                                    cur���ն� = Val(Split(.���￨��, ",")(3))
                                End If
                                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(1, .����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                
                                '���˺�:24491
                                curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                                
                                gbytWarn = BillingWarn(mstrPrivsOpt, .���� & IIf(.סԺ�� = "", "", "(סԺ��:" & .סԺ�� & " ����:" & .���� & ")"), .����ID, GetPatiWarnRange(.����ID, .��ҳID), mrsWarn, _
                                        cur���, cur���ն�, curTotal, .������, .�շ����, .Detail.�������, mstrWarn, True, , curItemMoney)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    .���� = dblPreTime
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        
                        If int���� <> 0 And mobjBill.Details(Bill.Row).���� <> 0 Then
                            If gclsInsure.GetCapability(supportʵʱ���, mobjBill.Details(Bill.Row).����ID, int����) Then
                                If gclsInsure.CheckItem(int����, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                    .���� = dblPreTime
                                    Call CalcMoneys(Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If .���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                .���� = dblPreTime
                                Call CalcMoneys(Bill.Row)
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                    End With
                        
                    Call ShowDetails(Bill.Row)
                    '��������д���������
                    For i = Bill.Row + 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).�������� = Bill.Row Then
                            '28136
                            '���������ĸ���,��Ҫ���¼��еĸ������и��³ɸ���
                            With mobjBill.Details(i)
                                If .Detail.���д��� = 0 Then  '�ǹ��д���
                                    If Abs(.����) <> Abs(.Detail.��������) Then GoTo NotCalc:
                                    .���� = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.��������
                                ElseIf .Detail.���д��� = 1 Then '�̶��Ĺ��д���
                                    .���� = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.�������� = 0, 1, .Detail.��������)
                                ElseIf .Detail.���д��� = 2 Then   '�������Ĺ��д���
                                    .���� = Val(Bill.Text) * .Detail.��������
                                Else
                                     GoTo NotCalc:
                                End If
                            End With
                            
                            Call CalcMoneys(i)
                            Call ShowDetails(i)
NotCalc:
                        End If
                    Next
                    
                 
                ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
                If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                    If CheckItemHaveSub(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '�Ϸ��Լ��
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "��Ŀ�۸�Ӧ��Ϊ������Ҫ�������ã������븺��������ʵ�֣�", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * mobjBill.Details(Bill.Row).���� > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '���û�ж�Ӧ��������Ŀ,���޷�����
                    If mobjBill.Details(Bill.Row).Detail.��� And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                        If Not (mobjBill.Details(Bill.Row).InComes(1).�ּ� = 0 And mobjBill.Details(Bill.Row).InComes(1).ԭ�� = 0) Then
                            strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).ԭ��, mobjBill.Details(Bill.Row).InComes(1).�ּ�, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(2) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).��׼����
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        
                        dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).��׼����
                        
                        mobjBill.Details(Bill.Row).InComes(1).��׼���� = Bill.Text '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
                        Call CalcMoneys(Bill.Row)
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 Then
                            curTotal = CalcOneTotal(Bill.Row)
                            If curTotal > 0 Then
                                If mobjBill.Details(Bill.Row).���￨�� = "" Then
                                    cur��� = 0
                                    cur���ն� = 0
                                Else
                                    cur��� = Split(mobjBill.Details(Bill.Row).���￨��, ",")(2)
                                    cur���ն� = Split(mobjBill.Details(Bill.Row).���￨��, ",")(3)
                                End If
                                
                                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(1, mobjBill.Details(Bill.Row).����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                gbytWarn = BillingWarn(mstrPrivsOpt, mobjBill.Details(Bill.Row).���� & IIf(mobjBill.Details(Bill.Row).סԺ�� = "", "", "(סԺ��:" & mobjBill.Details(Bill.Row).סԺ�� & " ����:" & mobjBill.Details(Bill.Row).���� & ")"), mobjBill.Details(Bill.Row).����ID, GetPatiWarnRange(mobjBill.Details(Bill.Row).����ID, mobjBill.Details(Bill.Row).��ҳID), mrsWarn, _
                                        cur���, cur���ն�, curTotal, mobjBill.Details(Bill.Row).������, mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, mstrWarn, True)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).InComes(1).��׼���� = dblPreMoney
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        Call ShowDetails(Bill.Row)
                    Else
                        Bill.Text = "0"
                        sta.Panels(2) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
                        Beep
                    End If
                End If
            Case "ִ�п���"
                If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Details(Bill.Row)
                        If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                            .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                             If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                        End If
                
                        'ҩƷ�����:��̬ҩ��,������ʱ��ҩƷҲҪ�����
                        If (.�շ���� = "4" And .Detail.��������) Or (InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ) Then
                            If .Detail.���� Or .Detail.��� Then '������ʱ��ҩƷ��治���ֹ����
                                If .���� * .���� > .Detail.��� Then
                                    If .�շ���� = "4" Then
                                        MsgBox "[" & .Detail.���� & "]Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox "[" & .Detail.���� & "]Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Cancel = True
                                End If
                            Else
                                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .ִ�в���ID) <> 0 Then
                                    If .���� * .���� > .Detail.��� Then
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Cancel = True
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                            Cancel = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                        
                        If CheckItemHaveSub(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
                        End If
                        If int���� <> 0 And mobjBill.Details(Bill.Row).���� <> 0 Then
                            If gclsInsure.GetCapability(supportʵʱ���, mobjBill.Details(Bill.Row).����ID, int����) Then
                                If gclsInsure.CheckItem(int����, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If mobjBill.Details(Bill.Row).���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                        
                    End With
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub


Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� = lngRow Then
            If mobjBill.Details(i).Detail.���д��� = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.����
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    If mobjBill.Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        Call bill_AfterAddRow(Bill.Rows - 1)
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.�Ա�
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.�Ա�
    End If
    '����:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub

Private Sub SetSubItem(lng���˿���ID As Long, int������Դ As Integer)
'����:�����շ���Ŀ��,���ص�ǰ�շ���Ŀ�Ĵ�����Ŀ�����ü�����,����ʾ�ڵ��ݿؼ���
'����:
'������:Bill_KeyDown��������Ŀ��
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long
Dim bln��������ۿ� As Boolean
Dim strժҪ As String, strPriceGrade As String

lngMainRow = Bill.Row               '�������
If gbln��������ۿ� Then            '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
    bln��������ۿ� = Not mobjBill.Details(lngMainRow).Detail.���ηѱ�
End If

With mobjBill.Details(lngMainRow)
    Set mcolDetails = New Details
    Set mcolDetails = GetSubDetails(.�շ�ϸĿID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            Call bill_AfterAddRow(Bill.Rows - 1)
        End If
        Bill.TextMatrix(Bill.Rows - 1, BillCol.�ѱ�) = "" '�б�Ҫ����
        
        'a.������ĿΪ��ҩƷ��Ŀ��ִ�п���
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).���) = 0 Then
             If mcolDetails(i).��� = .�շ���� Or mcolDetails(i).ִ�п��� = 0 Then
                '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                lngDoUnit = .ִ�в���ID
             Else
                '3.������ҩ��Ŀ��ִ�п���
                lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, _
                    mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, int������Դ, , .����ID)
             End If
        'b.������ĿΪҩƷ,���ĵ�ִ�п���
        Else
            lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, _
                mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, int������Դ, .ִ�в���ID, .����ID) '���Ĵ���ȱʡ������ִ�п�����ͬ
        End If
        
        
        '����֧����Ŀ��Ӧ���
        If Val(mobjBill.Details(lngMainRow).��ҩ����) > 0 Then
            If InStr(",5,6,7,", mcolDetails(i).���) > 0 Then
                strPriceGrade = mstrҩƷ�۸�ȼ�
            ElseIf mcolDetails(i).��� = "4" Then
                strPriceGrade = mstr���ļ۸�ȼ�
            Else
                strPriceGrade = mstr��ͨ�۸�ȼ�
            End If
            If Not CheckMediCareItem(mcolDetails(i).ID, Val(mobjBill.Details(lngMainRow).��ҩ����), mcolDetails(i).����, _
                mcolDetails(i).��� = False, , strPriceGrade) Then
                Exit Sub
            End If
        End If
        
        Call SetDetailtStock(lngDoUnit, mcolDetails(i))
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
        
        Call CalcMoney(Bill.Rows - 1, bln��������ۿ�)
        Call ShowDetails(Bill.Rows - 1)
        
        
        'CalcMoney���ȵ���GetuItemInsure���ܷ���ժҪ
         strժҪ = mobjBill.Details(Bill.Rows - 1).ժҪ
         strժҪ = gclsInsure.GetItemInfo(Val(mobjBill.Details(lngMainRow).��ҩ����), mobjBill.Details(lngMainRow).����ID, mcolDetails(i).ID, strժҪ, 2)
         mobjBill.Details(Bill.Rows - 1).ժҪ = strժҪ
        
    Next
    
    If bln��������ۿ� Then
        Call CalcMoney(lngMainRow, bln��������ۿ�) '�����������Ӧ����ʵ��,��Ϊ��û�м������ǰ�����ǰ������������.
        
        Call Calc��������ʵ��(lngMainRow)
    End If
End With

End Sub

Private Sub Calc��������ʵ��(ByVal lngMainRow As Long)
'����:����������ۿ�ʱ,����ָ�����������ID�ĵ�һ��������Ŀ���������ʵ�ս��
'����:  lngMainRow-������ID

Dim i As Long, j As Long
Dim cur����ǰӦ�պϼ� As Currency     '��¼�����������Ӧ�պϼ�
Dim cur���ۺ�ʵ�� As Currency


With mobjBill
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).�������� <> lngMainRow Then Exit For    '��ȻĿǰ�����˲������ڴ����м������������,����һ�ŵ�����������,Ϊ�˽������ܵ�����,����ȫ��ɨ��
        If i = lngMainRow Or .Details(i).�������� = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur����ǰӦ�պϼ� = cur����ǰӦ�պϼ� + .Details(i).InComes(j).Ӧ�ս��
            Next
        End If
    Next
       
    '���շ���Ŀ�漰����������ۿ�,����ҩƷ�ɱ����յĲ������贫��
    cur���ۺ�ʵ�� = CCur(Format(ActualMoney(.Details(lngMainRow).�ѱ�, .Details(lngMainRow).InComes(1).������ĿID, cur����ǰӦ�պϼ�, 0, 0, 0, 0), gstrDec))
    cur���ۺ�ʵ�� = cur���ۺ�ʵ�� - cur����ǰӦ�պϼ� + .Details(lngMainRow).InComes(1).Ӧ�ս��
    .Details(lngMainRow).InComes(1).ʵ�ս�� = Format(cur���ۺ�ʵ��, gstrDec)
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
'����:��������ִ�п��ҵı仯,ˢ�·�ҩ�����ִ�п���

    Dim i As Long, j As Long, lng���˿���ID As Long
    
    With mobjBill
        '��ȡ���д����ִ�п�������,������ȡ(��Ϊ�����ϵĴ�����Ϣ�������޸Ĺ���)
        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID)
        
        lng���˿���ID = .Details(lngRow).����ID
        If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
        
        For i = lngRow + 1 To .Details.Count
            If .Details(i).�������� = lngRow Then
                '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
                If InStr(",4,5,6,7,", .Details(i).�շ����) = 0 Then
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).ִ�п��� = 0 Then
                                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                 .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                            Else
                                '3.������ҩ��Ŀ��ִ�п���
                                .Details(i).ִ�в���ID = Get�շ�ִ�п���ID(mcolDetails(j).���, mcolDetails(j).ID, _
                                    mcolDetails(j).ִ�п���, lng���˿���ID, Get��������ID, Get������Դ(lngRow), , .Details(i).����ID)
                            End If
                        End If
                    End If
                    
                    'ˢ����ʾ����ִ�п���
                    If .Details(i).ִ�в���ID <> 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            If mbytInState = 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                '�������ֻ(��)��ʾ����
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!����
                            End If
                        Else
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                    
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
'���ܣ��жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = lngRow Then
                CheckItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim i As Long, curTotal As Currency
    Dim arrMoney As Variant, strMoney As String, arrPatiInfo As Variant
    Dim rsTmp As ADODB.Recordset, strStock As String
    Dim lng����ID As Long, strҩ��IDs As String
        
    If Row = 0 Then Exit Sub
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    
    If Not mblnEnterCell Then Exit Sub
    
    
    If mlngPreRow <> Row Then
        '���㵱ǰ���˷���
         curTotal = CalcOneTotal(Row)
    End If
    
    If mobjBill.Details.Count = 0 And mcolPatiInfo.Count > 0 Then
        '��ʾӤ����
        arrPatiInfo = Split(mcolPatiInfo("R" & Row), ",")
        Call LoadPatientBaby(cboBaby, Val(arrPatiInfo(0)), Val(arrPatiInfo(1)))
        Call zlControl.CboLocate(cboBaby, arrPatiInfo(2), True)
    End If
    
    If Not Bill.Active Or mstrInNO <> "" Then
        If mobjBill.Details.Count = 0 And mcolPatiInfo.Count > 0 Then  '����ʱ(�����)��ʾ���˷���
            If Val(mcolPatiInfo("R" & Row)) <> 0 And mlngPreRow <> Row Then
                Set rsTmp = GetMoneyInfo(Val(mcolPatiInfo("R" & Row)), , True, 2)
                If Not rsTmp Is Nothing Then
                    'sta.Panels(3).Text = Bill.TextMatrix(Row, BillCol.����) & "Ԥ��:" & Format(rsTmp!Ԥ�����, "0.00")
                    'sta.Panels(3).Text = sta.Panels(3) & "/����:" & Format(rsTmp!�������, gstrDec)
                    'sta.Panels(3).Text = sta.Panels(3) & "/���:" & Format(rsTmp!Ԥ����� - rsTmp!�������, "0.00")
                    '30604
                    Call SetStatuPatiInfor(Bill.TextMatrix(Row, BillCol.����), Val(Nvl(rsTmp!Ԥ�����)), _
                        Val(Nvl(rsTmp!�������)), Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������)))
                End If
            End If
        End If
        If Not Bill.Active Then mlngPreRow = Row: Exit Sub
    End If
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '����б༭����������ɫ
        mlngPreRow = Row
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------
    '1.�иı��������ݴ��������     mlngPreRow    ��ǰ���Ƿ�ı�
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '��ʾ���
            If InStr(",5,6,7,", .�շ����) > 0 And .�շ�ϸĿID <> 0 Then
                If Not gbln���뷢ҩ Then
                    If gbln����ҩ�� Or gbln����ҩ�� Then
                        strStock = GetStockInfo(.�շ�ϸĿID, gbln����ҩ��, gbln����ҩ��, gblnסԺ��λ)
                        If strStock <> "" Then
                            If InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0 Then
                                sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "�п��:" & strStock
                            Else
                                sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "���п��."
                            End If
                        End If
                    End If
                    If strStock = "" Then
                        '��ʱ���¿����ʾ
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        Call ShowStock(.Detail.����, .Detail.���)
                    End If
                Else
                    strҩ��IDs = Decode(.�շ����, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                    If strҩ��IDs <> "" Then
                        .Detail.��� = GetMultiStock(.�շ�ϸĿID, strҩ��IDs)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        Call ShowStock(.Detail.����, .Detail.���)
                    Else
                        sta.Panels(2) = ""
                    End If
                End If
            ElseIf .�շ���� = "4" And .Detail.�������� And .�շ�ϸĿID <> 0 Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                Call ShowStock(.Detail.����, .Detail.���)
            Else
                sta.Panels(2) = ""
            End If
            
            If mobjBill.Details.Count >= Row Then
                Call LoadPatientBaby(cboBaby, .����ID, .��ҳID)
                Call zlControl.CboLocate(cboBaby, .Ӥ����, True)
            End If
            
            Bill.ColData(BillCol.����) = BillColType.Text
            Bill.ColData(BillCol.���) = IIf(gbln�շ���� And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton
            
            '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
            If CheckItemHaveSub(Row) Or .�������� > 0 Then
                Bill.ColData(BillCol.����) = BillColType.Text_UnModify
                Bill.ColData(BillCol.���) = BillColType.Text_UnModify
                Bill.ColData(BillCol.��Ŀ) = BillColType.Text_UnModify
            End If
            
            '����Ƿǵ���״̬
            If mbytInState <> 2 Then
                If .�շ���� = "7" And gblnPay Then
                    Bill.ColData(BillCol.����) = BillColType.Text
                Else
                    Bill.ColData(BillCol.����) = BillColType.UnFocus
                End If
                
                '���������������
                If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                    And Not (.�շ���� = "4" And .Detail.��������) Then
                    Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '����
                    Bill.ColData(BillCol.����) = BillColType.Text '����
                Else
                    Bill.ColData(BillCol.����) = BillColType.Text '����
                    Bill.ColData(BillCol.����) = BillColType.UnFocus '����
                End If
                
                If .�շ���� = "F" Then
                    Bill.ColData(BillCol.��־) = BillColType.CheckBox
                Else
                    Bill.ColData(BillCol.��־) = BillColType.UnFocus
                End If
                
                If .Key = "1" Then    'ָ���˹̶�ҩ��ʱ,��������ѡ��ִ�п���
                    Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus
                Else
                    Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox
                End If
                
                If mblnOne Then Bill.ColData(BillCol.���) = BillColType.UnFocus 'ֻ����һ�����
            End If
            
            '�޸�ʱ��̬��ȡ���˷�����Ϣ,�Լӿ��ٶ�
            If mstrInNO <> "" And .���￨�� = "" Then
                '�޸�ǰ�ĵ�ǰ���ݵĲ��˷��ý��
                mcurModiMoney = GetBillMoney(2, mstrInNO, .����ID)
                
                '����ʣ���(���þ��￨����)
                Set rsTmp = Nothing
                Set rsTmp = GetMoneyInfo(.����ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
                If Not rsTmp Is Nothing Then
                   'sta.Panels(3).Text = .���� & "Ԥ��:" & Format(rsTmp!Ԥ�����, "0.00")
                   ' sta.Panels(3).Text = sta.Panels(3) & "/����:" & Format(rsTmp!�������, gstrDec)
                   ' sta.Panels(3).Text = sta.Panels(3) & "/���:" & Format(rsTmp!Ԥ����� - rsTmp!�������, "0.00")
                    
                    Call SetStatuPatiInfor(.����, Val(Nvl(rsTmp!Ԥ�����)), _
                        Val(Nvl(rsTmp!�������)), Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������)))
                    
                    
                    .���￨�� = rsTmp!Ԥ����� & "," & rsTmp!������� & "," & rsTmp!Ԥ����� - rsTmp!�������
                Else
                    .���￨�� = "0,0,0"
                End If
                '���˵��շ��ö�
                .���￨�� = .���￨�� & "," & GetPatiDayMoney(.����ID) - mcurModiMoney
            End If
        End With
    End If
    
        
    '���㲡��ʵ��ʣ���
    If mlngPreRow <> Row Then
        If Bill.TextMatrix(Row, BillCol.����) = "" Then
            sta.Panels(3).Text = "": picStatuPancl.Visible = False: lblStatuPati.Caption = ""
        Else
            If mobjBill.Details.Count >= Row Then
                If InStr(mobjBill.Details(Row).���￨��, ",") > 0 Then
                    arrMoney = Split(mobjBill.Details(Row).���￨��, ",")
                    'sta.Panels(3).Text = mobjBill.Details(Row).���� & "Ԥ��:" & Format(Val(arrMoney(0)), "0.00")
                    'sta.Panels(3).Text = sta.Panels(3) & "/����:" & Format(Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
                    'sta.Panels(3).Text = sta.Panels(3) & "/���:" & Format(Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
                    Call SetStatuPatiInfor(mobjBill.Details(Row).����, Val(arrMoney(0)), _
                                     Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0))
                                          
                End If
            Else
                '�µ�δ�������
                If mrsInfo.State = 1 Then
                    lng����ID = mrsInfo!����ID
                ElseIf mobjBill.Details.Count < Row And mobjBill.Details.Count >= Row - 1 And Row > 1 Then
                    lng����ID = mobjBill.Details(Row - 1).����ID
                End If
                If lng����ID > 0 Then
                    strMoney = GetMoneyStr(lng����ID)
                    If InStr(strMoney, ",") > 0 Then
                        arrMoney = Split(strMoney, ",")
                        'sta.Panels(3).Text = Bill.TextMatrix(Row, BillCol.����) & "Ԥ��:" & Format(Val(arrMoney(0)), "0.00")
                        'sta.Panels(3).Text = sta.Panels(3) & "/����:" & Format(Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
                        'sta.Panels(3).Text = sta.Panels(3) & "/���:" & Format(Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
                        Call SetStatuPatiInfor(Bill.TextMatrix(Row, BillCol.����), Val(arrMoney(0)), _
                                        Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0))
                    End If
                End If
            End If
        End If
        
        '������δ�������,��ָ��е�����
        If mobjBill.Details.Count < Bill.Row Then
            Bill.ColData(BillCol.����) = BillColType.Text
            Bill.ColData(BillCol.���) = IIf(gbln�շ���� And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton
        End If
    End If
     
    
    
    '-----------------------------------------------------------------
    '2.�иı��������ݴ������ʾ����
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then
        Call FillBillComboBox(Bill.Row, Bill.Col, True) '�������
    End If
    
    If gbln�շ���� And Bill.TextMatrix(Row, BillCol.���) = "" And mblnOne Then
        mrsClass.Filter = "����=" & gstr�շ����
        Bill.TextMatrix(Row, BillCol.���) = mrsClass!���
        Bill.RowData(Row) = Asc(mrsClass!����)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "����"
            Bill.TextLen = 19
        Case "�ѱ�"
            SetWidth Bill.cboHwnd, 70
            For i = 0 To Bill.ListCount - 1
                If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                    Bill.ListIndex = i: Exit For
                End If
            Next
        Case "���" '���������ʱ���ᶨλ�������
            SetWidth Bill.cboHwnd, 65
            '������Ϊ��,���Զ�Ĭ��Ϊ��һ�շ�ϸĿ�����
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "����=" & gstr�շ����
                    Bill.TextMatrix(Row, Col) = mrsClass!���
                    Bill.RowData(Row) = Asc(mrsClass!����)
                ElseIf Row > 1 Then
                    Bill.ListIndex = GetBillIndex(Bill.TextMatrix(Row - 1, Col))
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = GetBillIndex(Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "ִ�п���"
            SetWidth Bill.cboHwnd, 110
        Case "��"
            Bill.TextLen = 3: Bill.TextMask = "0123456789" & Chr(8)
        Case "����"
            Bill.TextLen = 8: Bill.TextMask = "0123456789." & Chr(8)
            If mobjBill.Details.Count >= Bill.Row Then
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                    If InStr(mstrPrivsOpt, ";ҩƷ����С��;") = 0 Then
                        Bill.TextMask = Replace(Bill.TextMask, ".", "")
                    End If
                End If
                '�ɷ����븺��
                If Not mobjBill.Details(Bill.Row).Detail.���� Then
                    If InStr(",5,6,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                        If InStr(mstrPrivsOpt, ";��ҩ��������;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    ElseIf InStr(",7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                        If InStr(mstrPrivsOpt, ";��ҩ��������;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    Else
                        If InStr(mstrPrivsOpt, ";���Ƹ�������;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    End If
                    
                    If InStr(Bill.TextMask, "-") > 0 And mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!����) Then
                            If Not gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����) Then
                                Bill.TextMask = Replace(Bill.TextMask, "-", "")
                            End If
                        End If
                    End If
                End If
            End If
        Case "����"
            Bill.TextLen = 10: Bill.TextMask = "0123456789." & Chr(8)
    End Select
    If Bill.MsfObj.ColIsVisible(Bill.Col) = False Then
        Bill.MsfObj.LeftCol = Bill.Col
    End If
    '����,����������е����ʱ,�������л�û�п�ʼ
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_KeyPress(KeyAscii As Integer)
    If Bill.TextMatrix(0, Bill.Col) = "����" And Bill.Active And Bill.ColData(Bill.Col) <> BillColType.Text_UnModify Then
         If cbo��������.ListIndex <> -1 And KeyAscii = Asc("*") Then
            KeyAscii = 0
            Call FillPatient(cbo��������.ItemData(cbo��������.ListIndex))
            If Bill.Top + Bill.CellTop + lvwPati.Height > sta.Top Then
                lvwPati.Top = Bill.Top + Bill.CellTop - lvwPati.Height - 30
            Else
                lvwPati.Top = Bill.Top + Bill.CellTop + Bill.RowHeight(1) - 15
            End If
            lvwPati.Visible = True
            lvwPati.SetFocus
        End If
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    Dim i As Long, lngParent As Long
    
    If Bill.Row <= mobjBill.Details.Count Then
        mobjBill.Details(Bill.Row).Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
        For i = Bill.Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = Bill.Row Then
                mobjBill.Details(i).Ӥ���� = mobjBill.Details(Bill.Row).Ӥ����
            End If
        Next
    End If
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub SetDefaultDoctor()
'����:����ȱʡ������
    If cbo������.ListCount = 0 Then Exit Sub
    
    If cbo������.ListCount = 1 Then
        cbo������.ListIndex = 0
    Else
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!סԺҽʦ) Then
                Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, mrsInfo!סԺҽʦ, True))
            End If
        End If
    End If
End Sub

Private Sub cboDrawDept_Click()
    Dim lng��ҩ����ID As Long
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cboDrawDept.ListIndex <> -1 Then lng��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
    If mobjBill.��ҩ����ID = lng��ҩ����ID Then Exit Sub
    mobjBill.��ҩ����ID = lng��ҩ����ID
End Sub

Private Sub cboDrawDept_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 And Not cboDrawDept.Locked Then
        lngIdx = zlControl.CboMatchIndex(cboDrawDept.hWnd, KeyAscii)
        If lngIdx = -1 And cboDrawDept.ListCount > 0 Then lngIdx = 0
        cboDrawDept.ListIndex = lngIdx
    ElseIf KeyAscii = 13 Then
        If cboDrawDept.ListIndex = -1 Then
            Beep
        Else
            mobjBill.��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, lng��������ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    If mobjBill.��������ID = lng��������ID Then Exit Sub
    mobjBill.��������ID = lng��������ID
    
    '����:
    If mrs��ҩ����.RecordCount <> 0 Then
        For i = 0 To cboDrawDept.ListCount - 1
             If cboDrawDept.ItemData(i) = lng��������ID Then
                mobjBill.��ҩ����ID = lng��������ID
                cboDrawDept.ListIndex = i: Exit For
             End If
        Next
    End If
    
    '��������ȷ��ҽ��
    If Not gblnFromDr Then
        If cbo��������.ListIndex <> -1 Then
            If gbln������ Then
                Call FillDoctor(cbo������, mrs������)
            Else
                Call FillDoctor(cbo������, mrs������, lng��������ID)
            End If
            Call SetDefaultDoctor
        Else
            cbo������.Clear
        End If
        Call cbo������_Click
    End If
        
    '�������������Ŀ��ִ�п���(�޸ĺͲ鿴ʱ����ԭ��)
    If cbo��������.ListIndex <> -1 And cbo��������.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
               '�������շ���Ŀ
                If InStr(",4,5,6,7,", .Detail.���) = 0 And .Detail.ִ�п��� = 6 Then '6-�����˿���
                    .ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    'ˢ����ʾ����ִ�п���
                    If i <= Bill.Rows - 1 And .ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                End If
            End With
        Next
    End If
    
    If Not gblnFromDr Then '�������ɿ����˶�����ʱ
        If mobjBill.Details.Count > 0 And cbo��������.ListIndex <> mlngPreUnit And Visible And Not txtIn.Enabled Then
            MsgBox "���ѣ����Ѿ����Ŀ�������Ϊ""" & zlStr.NeedName(cbo��������.Text) & """,ע���鵥���еĲ����Ƿ����ڸÿ��ң�", vbInformation, gstrSysName
        End If
    End If
    
    mlngPreUnit = cbo��������.ListIndex
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    If cbo��������.Text <> "" And cbo��������.ListIndex < 0 Then cbo��������.Text = ""
End Sub

Private Sub cbo������_Click()
    Dim lng������ID As Long, lng��������ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text)) Then Exit Sub
    
    mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    If gblnFromDr Then
        If cbo������.ListIndex <> -1 Then
            lng������ID = cbo������.ItemData(cbo������.ListIndex)
            If cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
            
            Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lng������ID)
            If lng��������ID > 0 Then Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
            If cbo��������.ListIndex = -1 Then Call SetDefaultDept(cbo��������, mrs��������, mrs������, lng������ID)
        Else
            cbo��������.Clear
        End If
        Call cbo��������_Click
    End If
    
    '��ʿ���
    If Bill.Active Then
        If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
            And Bill.RowData(Bill.Rows - 1) <> 0 Then
            '�����Ч����
            Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = ""
            Bill.RowData(Bill.Rows - 1) = 0
        ElseIf Bill.Col = BillCol.��� Then
            Call Bill_EnterCell(Bill.Row, Bill.Col) 'ˢ��
        End If
    End If
    
    '��ʿ���:�жϷǷ�����
    If CheckInhibitiveByNurse(mobjBill, mrs������) Then
        MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
    End If
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo������.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cbo������_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Text <> "" Then
        If cbo.FindIndex(cbo������, zlStr.NeedName(cbo������.Text), True) = -1 Then cbo������.ListIndex = -1: cbo������.Text = ""
    End If
    If cbo������.Text = "" Then Call cbo������_KeyPress(vbKeyReturn)
    '����������ȷ��������ʱ,���ܴ�ʱ��ѡ������,��ȥ�����������Һ�����ѡ
    If gblnFromDr And gbln������ And cbo������.ListIndex = -1 And mobjBill.Details.Count > 0 Then Cancel = True
End Sub

Private Sub chkCancel_Click()
    Dim i As Long
    
    mstrInNO = ""
    Call NewBill
    Call ClearRows
    Call Bill.ClearBill
    
    Bill.AllowAddRow = (chkCancel.Value = 0)
    
    Call SetDrawDrugDeptVisible
        
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        
        picUnit.Enabled = False
        fraAppend.Enabled = False
        chkIn.Enabled = False
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = BillColType.Text_UnModify
        Next
        Call ShowDeleteCol(True)
        Bill.Active = True
        
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then cbo������.Visible = False: lbl������.Visible = False
        Call SetDisible
        cboNO.Locked = False
        cboNO.SetFocus
    Else
        chkCancel.ForeColor = 0
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
        Call ShowDeleteCol(False)
        
        If gbytBilling = 2 Then
            Call SetDisible
            Bill.Active = False
            cboNO.Locked = False
            cboNO.SetFocus
        Else
            Call SetDisible(True)
            picUnit.Enabled = True
            fraAppend.Enabled = True
            chkIn.Enabled = True
            Bill.Active = True
            cboNO.Locked = True
            Bill.SetFocus
        End If
        
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then cbo������.Visible = True: lbl������.Visible = True
        Call cbo��������_Click
    End If
End Sub

Private Sub chkIn_Click()
    sta.Panels(2) = ""
    If chkIn.Value = Checked Then
        txtIn.Enabled = True
        txtIn.BackColor = &H80000005
        sta.Panels(2) = "������Ҫ����ļ��ʵ����ݺ���"
        txtIn.SetFocus
    Else
        txtIn.Text = ""
        txtIn.Enabled = False
        txtIn.BackColor = &HE0E0E0
        Bill.SetFocus
    End If
End Sub

Private Sub chk�Ӱ�_Click()
    Dim blnAdd As Boolean
    
    If mbytInState = 1 Or chkCancel.Value = 1 Or gbytBilling = 2 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk�Ӱ�.Visible Then Exit Sub
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk�Ӱ�.Value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Checked
        End If
    End If
    If chk�Ӱ�.Value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Unchecked
        End If
    End If
    mobjBill.�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    Call CalcMoneys
    Call ShowDetails
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count > 0 And Bill.Active And mbytInState = 0 And mstrInNO = "" Then
        Call Form_KeyDown(vbKeyF6, 0): Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, BillCol.��Ŀ) <> "" Then Bill.TextMatrix(i, Bill.Cols - 1) = ""
    Next
End Sub
Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������(�����������Ҫ����,�����ڸ�������,���ֹ
    '���:
    '����:lngRow-���ظ�����������
    '����:������������û�����븽������,����true,���򷵻�False
    '����:
    '�޸�:���˺�(�˺�ʱ,���Ӷ�λ����),���Ӳ���;strBackNo
    '����:2009/7/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, lngRow As Long   'ָ����
    Dim i As Long, j As Long
    
    
    For i = 1 To mobjBill.Details.Count
        lngCount = 0: lngRow = 0
        For j = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).����ID = mobjBill.Details(j).����ID Then
                If mobjBill.Details(j).�շ���� = "F" Then
                   If mobjBill.Details(j).���ӱ�־ = 0 Then lngCount = 0: Exit For     '������Ҫ����,�򲻼��,ֱ�ӷ���true
                   lngCount = lngCount + 1  '��ʾ��������
                   If lngRow <= 0 Then lngRow = j
                End If
            End If
        Next
        If lngCount <> 0 Then Exit For
    Next
    
    If lngCount <> 0 Then
          MsgBox "�����в�����Ҫ����,�����ڸ�������,���飡", vbInformation, gstrSysName
          If Bill.Rows > lngRow Then Bill.Row = lngRow
          If Bill.Visible Then Bill.SetFocus
          Exit Function
    End If
    CheckMainOperation = True
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset, rsExamine As ADODB.Recordset, rsFeeItem As ADODB.Recordset
    Dim strSQL As String, strInfo As String, strTmp As String, strAddDate As String, strRows As String
    Dim arrPati() As Variant, strMoney As String, str���ܺ� As String, str�������� As String
    Dim strPatis As String, i As Long, j As Long, lng����ID As Long, lng��ҳID As Long, lng���� As Long
    Dim curModiMoney As Currency, Curdate As Date, cur��� As Currency, dbl���� As Double
    Dim strInsure As String, arrInsure As Variant
    Dim dblTotal As Double, blnTrans As Boolean
    Dim colStock As Collection
    Dim arrSMSQL As Variant, str��������IDs As String, str������s As String
    Dim cllPro As Collection
    Dim rsItems As ADODB.Recordset
    
    If mbytInState = 3 Or (mbytInState = 0 And chkCancel.Visible And chkCancel.Value = 1) Then
        If mbytInState = 0 And mstrInNO = "" Then
            MsgBox "û�ж�ȡ��������,�������ʣ�", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        For i = 1 To UBound(marrSerial)
            If Bill.TextMatrix(i, Bill.Cols - 1) = "��" Then
                strRows = strRows & "," & marrSerial(i)
            End If
        Next
        If strRows = "" Then
            MsgBox "������ѡ��һ��Ҫ���ʵķ��ã�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If zlCheckIsExistsApplied(mstrInNO, strRows, str��������IDs, str������s) Then
            '����:47416
            If MsgBox("ע��:" & vbCrLf & "    ����" & mstrInNO & "�д����������ʵ���Ŀ,���ʺ�,�����Զ�ȡ��" & vbCrLf & "�����˵�������Ŀ,�Ƿ��������?" & vbCrLf & "����������: " & str������s, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        '������ѡ����
        strRows = Mid(strRows, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strRows, ",")) + 1 = i Then strRows = ""
        
        If strRows <> "" And InStr(1, mstrPrivsOpt, ";��������;") = 0 Then
            MsgBox "��û�в������ʵ�Ȩ�ޣ�ֻ�ܶԸõ���ȫ�����ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�ж��Ƿ����ҽ�����˷���,�Լ��Ƿ�����������(ֻҪ��һ��������,������)
        'ȥ����ҽ������ƥ����
        If gbytBilling = 0 Then
            Call GetBillInsures(strInsure, mstrInNO, mstrTime)
            If strInsure <> "" Then
                arrInsure = Split(strInsure, ",")
                If strRows <> "" Then
                    For i = 0 To UBound(arrInsure)
                        If gclsInsure.GetCapability(support���������ϴ�, , arrInsure(i)) Then
                            If Not gclsInsure.GetCapability(support�����ݳ�������, , arrInsure(i)) Then
                                MsgBox "��Ϊҽ��������Ҫ,�õ����е���Ŀ����ȫ�����ʣ�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            End If
        End If
        
        
        Set rsFeeItem = GetNOFeeItem(mstrInNO, 2, strRows)
        If rsFeeItem.RecordCount > 0 Then
            For i = 1 To Bill.Rows - 1
                rsFeeItem.Filter = "���=" & marrSerial(i)
                If rsFeeItem.RecordCount > 0 Then
                    If Not (InStr(",5,6,7,", rsFeeItem!�շ����) > 0 And gbln���뷢ҩ) Then
                        strTmp = mcolPatiInfo("R" & i)
                        Set rsTmp = GetPatientFeeItemTotal(Split(strTmp, ",")(0), Split(strTmp, ",")(1), mstrInNO)
                        rsTmp.Filter = "�շ�ϸĿid=" & rsFeeItem!�շ�ϸĿID & " And ִ�в���id=" & rsFeeItem!ִ�в���ID
                        If rsTmp.RecordCount > 0 Then
                            If Bill.TextMatrix(i, BillCol.����) * Bill.TextMatrix(i, BillCol.����) > rsTmp!���� Then
                                MsgBox "��" & i & "�������������ڿ���������" & rsTmp!���� & ".", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        Else
                            MsgBox "��" & i & "�п���������Ϊ��.", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
         '����:47416
        Set cllPro = New Collection
        If str��������IDs <> "" Then
            strSQL = "zl_���˷�������_Delete('" & str��������IDs & "')"
            zlAddArray cllPro, strSQL
        End If
        strSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "','" & strRows & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        zlAddArray cllPro, strSQL
        
        cmdOK.Enabled = False
        On Error GoTo errH
            blnTrans = True
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            'Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
            'ҽ�����������ϴ�(ֻҪ��һ���ɹ����ύ)
            If gbytBilling = 0 And strInsure <> "" Then
                For i = 0 To UBound(arrInsure)
                    If gclsInsure.GetCapability(support���������ϴ�, , arrInsure(i)) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , arrInsure(i)) Then
                        If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrInsure(i)) Then
                            If i = 0 Then gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
                        End If
                    End If
                Next
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ�����������ϴ�
        If gbytBilling = 0 And strInsure <> "" Then
            For i = 0 To UBound(arrInsure)
                If gclsInsure.GetCapability(support���������ϴ�, , arrInsure(i)) And gclsInsure.GetCapability(support������ɺ��ϴ�, , arrInsure(i)) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrInsure(i)) Then
                        MsgBox "������ " & GetInsureName(Val(arrInsure(i))) & " �ķ�����ҽ������ʧ�ܣ���Щ���������ʡ�", vbInformation, gstrSysName
                    End If
                End If
            Next
        End If
        
        cmdOK.Enabled = True
        If mbytInState = 0 Then
            mstrInNO = "": mstr����IDs = ""
            cboNO.Text = ""
            Call ClearRows
            Call Bill.ClearBill
            Call NewBill
            
            chkCancel.Value = 0
            
            If gbytBilling = 2 Then
                cboNO.SetFocus
            Else
                Bill.SetFocus
            End If
        Else
           gblnOK = True: Unload Me: Exit Sub
        End If
    ElseIf mbytInState = 2 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "������Ϸ��ķ���ʱ�䣡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        strInfo = Check����ʱ��(CDate(txtDate.Text), cboNO.Text)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If Not SaveModi() Then Exit Sub
        gblnOK = True: Unload Me: Exit Sub
    ElseIf Bill.Active And chkCancel.Value = 0 Then '�������뵥��״̬
        If mobjBill.Details.Count = 0 Then
            MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbExclamation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Checkִ�п���
        If i <> 0 Then
            MsgBox "�����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If mobjBill.��������ID = 0 Then
            MsgBox "��ѡ�񿪵����ң�", vbExclamation, gstrSysName
            cbo��������.SetFocus: Exit Sub
        End If
        
        '�Ƿ���
        dbl���� = 0
        For i = 1 To mobjBill.Details.Count
            '27467,52828
            If mobjBill.Details(i).���� <> 0 And dbl���� = 0 Then
                dbl���� = mobjBill.Details(i).����
            End If
            If mobjBill.Details(i).�շ�ϸĿID = 0 Then
                MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbExclamation, gstrSysName
                Bill.SetFocus: Exit Sub
            ElseIf InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                '�ռ�ҩƷ�ķ�ҩҩ��
                strTmp = strTmp & "," & mobjBill.Details(i).�շ�ϸĿID
            End If
        Next
        '27467,52828
        If mbytInState = 0 And FormatEx(dbl����, 7) = 0 Then
            MsgBox "����������Ҫ��һ����Ϊ�������,���飡", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
         '���ҩƷ�ķ�ҩҩ����Ӧ�ķ������
        If strTmp <> "" And Not gbln���뷢ҩ Then
            strTmp = Mid(strTmp, 2)
            Set rsTmp = GetServiceDept(strTmp)
            If Not rsTmp Is Nothing Then
                strTmp = ""
                For i = 1 To mobjBill.Details.Count
                    If InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                        strInfo = mobjBill.Details(i).�շ�ϸĿID
                        '�ȼ���Ƿ�������Ĵ洢�ⷿ
                        rsTmp.Filter = "�շ�ϸĿID=" & strInfo & " And ִ�п���id=" & mobjBill.Details(i).ִ�в���ID
                        If rsTmp.RecordCount = 0 Then
                            strTmp = strTmp & "," & i
                        Else
                            '�ټ���Ƿ�������ķ������(û�����÷�����ҵ�,��������IDΪ��)
                            rsTmp.Filter = "(" & rsTmp.Filter & " And ��������ID=" & mobjBill.Details(i).����ID & ") Or (" & rsTmp.Filter & " And ��������ID=0)"
                            If rsTmp.RecordCount = 0 Then
                                strTmp = strTmp & "," & i
                            End If
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    strTmp = Mid(strTmp, 2)
                    MsgBox "����,��" & strTmp & "��ҩƷ�Ƿ�Υ�����¹���:" & vbCrLf & vbCrLf & _
                        "A.ѡ���ִ�п��Ҳ���ҩƷ�Ĵ洢�ⷿ" & vbCrLf & _
                        "B.���˿��Ҳ�����ҩƷ�ڴ˴洢�ⷿ�ķ������.", _
                        vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        '���÷���ʱ����
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbExclamation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strSQL = ""
        For i = 1 To mobjBill.Details.Count
            If InStr(strSQL & ",", "," & mobjBill.Details(i).����ID & ",") = 0 Then
                strInfo = Check����ʱ��(CDate(txtDate.Text), mobjBill.Details(i).����ID)
                If strInfo <> "" Then
                    MsgBox strInfo, vbInformation, gstrSysName
                    txtDate.SetFocus: Exit Sub
                End If
                strSQL = strSQL & "," & mobjBill.Details(i).����ID
            End If
        Next
        
        If mobjBill.������ = "" And gbln������ Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        '��ʿ���:�жϷǷ�����
        If CheckInhibitiveByNurse(mobjBill, mrs������) Then
            MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        
        'ҽ���������ʼ��    ��Ϊ����Ա���������ҽ�����˵ĸ�������,�ٻ�Ϊҽ������,����Ҫ�ټ��һ��
        If InStr(mstrPrivsOpt, ";��������;") > 0 Then     '����������һ�ָ�������Ȩ��,�ſ����Ǹ���
            For i = 1 To mobjBill.Details.Count
                If Val(mobjBill.Details(i).��ҩ����) <> 0 Then
                    If mobjBill.Details(i).���� * mobjBill.Details(i).���� < 0 Then
                        If Not gclsInsure.GetCapability(support��������, mobjBill.Details(i).����ID, Val(mobjBill.Details(i).��ҩ����)) Then
                            MsgBox "�����е� " & i & " ���Ǹ���,����ҽ����֧�ָ������ʣ�", vbInformation, gstrSysName
                            Bill.SetFocus: Exit Sub
                        End If
                    End If
                End If
            Next
        End If
        
        '��Ժǿ�Ƽ���Ȩ�޼��
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If InStr(strPatis & ",", "," & .����ID & ",") = 0 Then
                    strPatis = strPatis & "," & .����ID
                    If Not PatiCanBilling(.����ID, .��ҳID, mstrPrivsOpt) Then Exit Sub
                    If zlPatiIS�����ѱ�Ŀ(.����ID, .��ҳID) = True Then     '����:28725
                        Exit Sub
                    End If
                    If zlIsAllowFeeChange(.����ID, .��ҳID) = False Then
                        Exit Sub
                    End If
                End If
            End With
        Next
       
              
        '����ְ����
        '���ѻ�ҽ������
        i = CheckDuty(, False)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.��Ŀ: Bill.SetFocus
            Exit Sub
        End If
        
        '���в�����Ŀ
        i = CheckDuty(, True)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.��Ŀ: Bill.SetFocus
            Exit Sub
        End If
                
        'ҽ��������Ŀ�Ƿ�������飬����ʱ�Ѽ�飬����ʱ�ټ������Ϊ��
        '1���������������ʱֻ��������2.���뵥��ʱδ���
        'ע��:���ܴ���ҽ���ͷ�ҽ�����˻�ϵ����,�Լ�ֻ��һ�з��õ����
        strInfo = "": strTmp = "": lng����ID = 0: lng��ҳID = 0: lng���� = 0
        str�������� = zlStr.NeedName(cbo��������.Text)
        For i = 1 To mobjBill.Details.Count
            lng��ҳID = mobjBill.Details(i).��ҳID
            lng���� = Val(mobjBill.Details(i).��ҩ����)
            
            If lng����ID <> mobjBill.Details(i).����ID And lng���� > 0 Then
                Set rsTmp = GetAuditRecord(lng����ID, lng��ҳID)
                Set rsExamine = GetExamineItem(strTmp, lng����)
                For j = 1 To rsExamine.RecordCount
                    rsTmp.Filter = "��ĿID=" & rsExamine!�շ�ϸĿID
                    If rsTmp.RecordCount = 0 Then
                        strInfo = strInfo & "," & GetRowByFeeItemID(mobjBill.Details, rsExamine!�շ�ϸĿID, lng����ID)
                    ElseIf Not IsNull(rsTmp!��������) Then
                        If mobjBill.Details(i).���� * mobjBill.Details(i).���� * IIf(gblnסԺ��λ, mobjBill.Details(i).Detail.סԺ��װ, 1) > rsTmp!�������� Then
                            MsgBox "��" & i & "���շ���Ŀ�����γ�������׼��ʹ������" & FormatEx(rsTmp!�������� / IIf(gblnסԺ��λ, mobjBill.Details(i).Detail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    rsExamine.MoveNext
                Next
                strTmp = ""
                
                
                If gclsInsure.GetCapability(supportʵʱ���, lng����ID, lng����) Then
                    If gclsInsure.CheckItem(lng����, 1, 2, MakeDetailRecord(mobjBill, mobjBill.������, str��������, 2, gbytBilling)) = False Then
                       Exit Sub
                    End If
                End If
            End If
            
            lng����ID = mobjBill.Details(i).����ID
            If lng���� > 0 Then
                strTmp = IIf(strTmp = "", "", strTmp & ",") & mobjBill.Details(i).�շ�ϸĿID
            End If
        Next
        
        If strTmp <> "" Then
            Set rsTmp = GetAuditRecord(lng����ID, lng��ҳID)
            Set rsExamine = GetExamineItem(strTmp, lng����)
            For j = 1 To rsExamine.RecordCount
                rsTmp.Filter = "��ĿID=" & rsExamine!�շ�ϸĿID
                If rsTmp.RecordCount = 0 Then
                    strInfo = strInfo & "," & GetRowByFeeItemID(mobjBill.Details, rsExamine!�շ�ϸĿID, lng����ID)
                ElseIf Not IsNull(rsTmp!��������) Then
                    i = GetRowByFeeItemID(mobjBill.Details, rsExamine!�շ�ϸĿID, lng����ID)
                    If mobjBill.Details(i).���� * mobjBill.Details(i).���� * IIf(gblnסԺ��λ, mobjBill.Details(i).Detail.סԺ��װ, 1) > rsTmp!�������� Then
                        MsgBox "��" & i & "���շ���Ŀ�����γ�������׼��ʹ������" & FormatEx(rsTmp!�������� / IIf(gblnסԺ��λ, mobjBill.Details(i).Detail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                rsExamine.MoveNext
            Next
        End If
        
        If strInfo <> "" Then
            MsgBox "��" & Mid(strInfo, 2) & "���շ���ĿҪ������,��ǰ����δ����׼ʹ��!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�������ò��˲�������
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).�շ����) = 0 Then
                If CheckItemHaveSub(i) Then
                    If Not CheckFeeItemLimitDept(mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).����ID, mobjBill.Details(i).����ID) Then
                        MsgBox "��" & i & "�е��շ���Ŀ�Ըò��˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                        Bill.Row = i: Bill.MsfObj.TopRow = i
                        Bill.Col = BillCol.��Ŀ: Bill.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        '�������ͼ��
        If Not Check�������� Then Exit Sub
                
        '���ʷ��౨��:������˼��ʱ���
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            ReDim arrPati(7) '������Ϣ����
            'ѭ�������ж�ÿ�����˴���
            For i = 1 To mobjBill.Details.Count
                '�ռ����в�����Ϣ
                If mobjBill.Details(i).����ID <> arrPati(0) Then
                    arrPati(0) = mobjBill.Details(i).����ID  '����ID
                    arrPati(1) = CStr(mobjBill.Details(i).���� & IIf(mobjBill.Details(i).סԺ�� = "", "", "(סԺ��:" & mobjBill.Details(i).סԺ�� & " ����:" & mobjBill.Details(i).���� & ")")) '����
                    arrPati(2) = CCur(CalcOneTotal(CLng(i), False)) '���ݽ��
                    arrPati(3) = CCur(mobjBill.Details(i).������) '������
                    arrPati(4) = GetMedPayMode(mobjBill.Details(i).ҽ�Ƹ���, mrsMedPayMode)
                    
                    'ˢ�¶�ȡ:Ԥ�����,�������;ʣ����;���շ���
                    curModiMoney = 0
                    If mstrInNO <> "" Then
                        curModiMoney = GetBillMoney(2, mstrInNO, mobjBill.Details(i).����ID)
                    End If
                    
                    strMoney = "0,0,0"
                    Set rsTmp = GetMoneyInfo(mobjBill.Details(i).����ID, IIf(gbytBilling = 0, curModiMoney, 0), True, 2)
                    If Not rsTmp Is Nothing Then
                        strMoney = rsTmp!Ԥ����� & "," & rsTmp!������� & "," & rsTmp!Ԥ����� - rsTmp!�������
                    End If
                    strMoney = strMoney & "," & GetPatiDayMoney(mobjBill.Details(i).����ID) - mcurModiMoney '���˵��շ��ö�
                    
                    For j = 1 To mobjBill.Details.Count
                        If mobjBill.Details(j).����ID = mobjBill.Details(i).����ID Then
                            mobjBill.Details(j).���￨�� = strMoney
                        End If
                    Next
                                    
                    'ʣ����,���շ��ö�
                    arrPati(5) = Val(Split(strMoney, ",")(2))
                    arrPati(6) = Val(Split(strMoney, ",")(3))
                    
                    '��������
                    arrPati(7) = Val(mobjBill.Details(i).��ҩ����)
                                    
                    cur��� = CCur(arrPati(5))
                    If gbln�����������۷��� Then cur��� = CCur(arrPati(5)) - GetPriceMoneyTotal(1, CLng(arrPati(0))) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                End If
                
                '���н��б���
                If CCur(arrPati(2)) > 0 Then
                    gbytWarn = BillingWarn(mstrPrivsOpt, CStr(arrPati(1)), mobjBill.Details(i).����ID, GetPatiWarnRange(mobjBill.Details(i).����ID, mobjBill.Details(i).��ҳID), mrsWarn, _
                            cur���, CCur(arrPati(6)), CCur(arrPati(2)), CCur(arrPati(3)), mobjBill.Details(i).�շ����, mobjBill.Details(i).Detail.�������, mstrWarn, True)
                    If gbytWarn = 2 Or gbytWarn = 3 Then
                        Bill.Row = i: Exit Sub
                    End If
                End If
            Next
        End If
        
        'ҩƷ���ɼ��
        strInfo = CheckDisable(mobjBill)
        If strInfo <> "" Then
            If strInfo Like "*(�������)*" Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            Else
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        '�����������
        If Not gbln�������� Then
            If Not CheckLimit(mobjBill, , gblnסԺ��λ) Then Exit Sub
        End If
        
        '��������ʱ��ҩƷͬһҩ���Ƿ����ظ�����
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.���� Or .Detail.���) _
                    And (InStr(",5,6,7,", .�շ����) > 0 Or .�շ���� = "4" And .Detail.��������) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And .ִ�в���ID = mobjBill.Details(j).ִ�в���ID And .����ID = mobjBill.Details(j).����ID Then
                            If .�շ���� = "4" Then
                                MsgBox "�� " & j & " �еķ�����ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����룬��ϲ���", vbInformation, gstrSysName
                            Else
                                MsgBox "�� " & j & " �еķ�����ʱ��ҩƷ""" & .Detail.���� & """��ͬһ��ҩ�����ظ����룬��ϲ���", vbInformation, gstrSysName
                            End If
                            Exit Sub
                        End If
                    Next
                End If
            End With
        Next
        
        'ҩƷ�����,71188:������,2014-04-03,�Բ������ѵ�ҲҪ���м��
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                If InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ Then
                    If .Detail.���� Or .Detail.��� Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ����ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 1 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            If MsgBox("�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,Ҫ������?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                        End If
                    End If
                ElseIf InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ And gblnStock Then
                    '���ݶ���Ŀ���Ǳ��ز���ָ����ҩ���Ŀ��֮��
                    strInfo = Decode(.Detail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                    If strInfo <> "" Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, 0)
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, 0)
                        
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ҩƷ""" & .Detail.���� & "]�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & _
                                "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                ElseIf .�շ���� = "4" And .Detail.�������� Then
                    If .Detail.���� Or .Detail.��� Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ����������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 1 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            If MsgBox("�� " & i & " ����������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,Ҫ������?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                        End If
                    End If
                End If
            End With
        Next
    
        '���ۼ��,105875
        If Not gobjPublicDrug Is Nothing Then
            'Private Function zlCheckPriceAdjustBySell(ByVal lngҩƷid As Long, ByVal lngҩ��id As Long) As Boolean
            '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
            '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
            'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
            '���۳���ʱֻ�ж�ҩ��
            '���أ�True-�����������۳��⣻false-���ܽ������۳���
            For i = 1 To mobjBill.Details.Count
                With mobjBill.Details(i)
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        If gobjPublicDrug.zlCheckPriceAdjustBySell(.�շ�ϸĿID, .ִ�в���ID) = False Then
                            Exit Sub
                        End If
                    End If
                End With
            Next
        End If
        
        '���˺�:22441,����������͸����������
        If CheckMainOperation = False Then Exit Sub
        
        '��Ŀ���������(��Ҫ��Ϊ�����������۲���)
        If Check������� > 0 Then Exit Sub
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 1, _
            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling)) = False Then
            Exit Sub
        End If
        
        '�����˷Ѽ��
        If Not CheckBillNegative Then Exit Sub
        
        '����������ϵ����Ч��
        'ҩƷ�Զ���ҩ
        mblnSendMateria = False
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If .�շ���� = "4" And .Detail.�������� Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    If Not CheckValidity(.�շ�ϸĿID, .ִ�в���ID, dblTotal) Then Exit Sub
             
                ElseIf InStr(1, ",5,6,7,", .�շ����) > 0 Then
                    '��ӡ��ҩ��,����ͨ����,�һ��۵�����
                    If gbytSendMateria <> 0 And mbytUseType = 0 And gbytBilling = 0 Then
                        'ȫ��ҩƷ��ȷ����ҩ���Ĳ��Զ���ҩ(���뷢ҩʱ,û��ȷ��ҩ��)
                        mblnSendMateria = .ִ�в���ID <> 0
                    End If
                End If
            End With
        Next
        If InStr(mstrPrivsOpt, ";ҩƷ��ҩ;") = 0 Then mblnSendMateria = False
        
        If mstrInNO <> "" Then
            If HaveExecute(2, mstrInNO, 2) Then
                MsgBox "�õ��ݰ�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ġ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If mblnSendMateria And gbytSendMateria = 2 Then
            If MsgBox("������ɺ��Զ�ִ�з�ҩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
        mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate     'ע��:��ӡ��ҩ��ʱҪ�õ����ʱ��
        If zlGetSaveDataItems_Plugin(mobjBill, rsItems, True) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(mlngModule, 2, False, gbytBilling = 1, "", rsItems) = False Then Exit Sub
        
        cmdOK.Enabled = False
        If Not SaveBill Then
            cmdOK.Enabled = True
            Exit Sub
        Else
            Call zlChargeSaveAfter_Plugin(mlngModule, 0, 0, False, 2, mobjBill.NO)
            If gbytBilling = 0 And gbln���ʴ�ӡ Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_113" & 3 + mbytUseType, Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
            ElseIf gbytBilling = 1 And gbln���۴�ӡ Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
            End If
                        
            '��ӡ��ҩ��
            If mblnSendMateria Then
                If MsgBox("����""" & mobjBill.NO & """��ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), 1)
                End If
            End If
        
            cmdOK.Enabled = True
            If mstrInNO = "" Then
                sta.Panels(2) = "��һ�ŵ���:" & mobjBill.NO
                Call ClearRows: Call Bill.ClearBill
                Call NewBill: mstrInNO = ""
                Bill.SetFocus
            Else '�޸�
                '���˺� ����:27083 ����:2009-12-25 10:09:21
                gblnOK = True: Unload Me: Exit Sub
            End If
        End If
    ElseIf Not Bill.Active Then '���סԺ����״̬
        If mstrInNO = "" Then
            MsgBox "û��סԺ���۵���,�������룡", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        'ȡ������˵������
        strSQL = ""
        For i = 1 To UBound(marrSerial)
            strSQL = strSQL & "," & marrSerial(i)
        Next
        strSQL = Mid(strSQL, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        
        '���ñ���
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            If Not AuditingWarn(mstrPrivsOpt, mrsWarn, mstrInNO, strSQL) Then Exit Sub
        End If
        
        'ȥ����ҽ������ƥ����
        Call GetBillInsures(strInsure, mstrInNO, , True)
        If strInsure <> "" Then arrInsure = Split(strInsure, ",")
        
        'ҩƷ�Զ���ҩ
        mblnSendMateria = False
        If gbytSendMateria <> 0 And mbytUseType = 0 And InStr(mstrPrivsOpt, ";ҩƷ��ҩ;") > 0 Then
            For i = 1 To Bill.Rows - 1
                If InStr(",����ҩ,�г�ҩ,�в�ҩ,", "," & Bill.TextMatrix(i, BillCol.���) & ",") > 0 Then '���ȡ����ʱû�д洢������,��Ϊ���������ж�
                    'ȫ��ҩƷ��ȷ����ҩ���Ĳ��Զ���ҩ(���뷢ҩʱ,û��ȷ��ҩ��)
                    mblnSendMateria = Trim(Bill.TextMatrix(i, BillCol.ִ�п���)) <> ""
                End If
            Next
        End If
        If mblnSendMateria And gbytSendMateria = 2 Then
            If MsgBox("������˺��Զ�ִ�з�ҩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        cmdOK.Enabled = False
        arrSMSQL = Array()
        Curdate = zlDatabase.Currentdate
        strAddDate = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSQL = "zl_סԺ���ʼ�¼_Verify('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "','" & strSQL & "',NULL," & strAddDate & ")"
        str���ܺ� = zlDatabase.GetNextNo(20)
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '׼���Զ���ҩ(����ͨ����),�����������в��ܶ�������
            If mblnSendMateria Then
                Set rsTmp = Get����ҩ�嵥(mstrInNO, Format(Curdate, "yyyy-MM-dd HH:mm:ss"), True)
                If rsTmp.RecordCount > 0 Then
                    ReDim arrSMSQL(rsTmp.RecordCount - 1)
                    For i = 0 To rsTmp.RecordCount - 1
                        arrSMSQL(i) = "ZL_ҩƷ�շ���¼_���ŷ�ҩ(" & rsTmp!�ⷿID & "," & rsTmp!ID & ",'" & UserInfo.���� & "'," & strAddDate & ",Null,Null,Null," & str���ܺ� & ")"
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            'ִ���Զ���ҩ
            For i = 0 To UBound(arrSMSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSMSQL(i)), Me.Caption)
            Next
            
            'ҽ���ϴ�(ֻҪ��һ���ɹ����ύ)
            If strInsure <> "" Then
                For i = 0 To UBound(arrInsure)
                    If gclsInsure.GetCapability(support�����ϴ�, , arrInsure(i)) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , arrInsure(i)) Then
                        strInfo = ""
                        If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , arrInsure(i)) Then
                            If i = 0 Then gcnOracle.RollbackTrans
                            If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                            If i = 0 Then cmdOK.Enabled = True: Exit Sub
                        End If
                    End If
                Next
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ���ϴ�
        If strInsure <> "" Then
            For i = 0 To UBound(arrInsure)
                If gclsInsure.GetCapability(support�����ϴ�, , arrInsure(i)) And gclsInsure.GetCapability(support������ɺ��ϴ�, , arrInsure(i)) Then
                    strInfo = ""
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , arrInsure(i)) Then
                        If strInfo <> "" Then
                            MsgBox strInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "�����е� " & GetInsureName(Val(arrInsure(i))) & " ������ҽ������ʧ��,��Щ��������ˣ�", vbInformation, gstrSysName
                        End If
                    End If
                End If
            Next
        End If
        
        On Error GoTo 0
        
        If gbytBilling = 2 And gbln��˴�ӡ And mblnPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mstrInNO, "�Ǽ�ʱ��=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
        End If
        
        '��ӡ��ҩ��
        If mblnSendMateria Then
            If MsgBox("����""" & mstrInNO & """��ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & mstrInNO, "�Ǽ�ʱ��=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), 1)
            End If
        End If
        
        cmdOK.Enabled = True
        mstrInNO = "": cboNO.Text = ""
        Call ClearRows: Call Bill.ClearBill
        Call NewBill: cboNO.Locked = False
        cboNO.SetFocus
    End If
    gblnOK = True
    Call SetDrawDrugDeptEnabled
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    cmdOK.Enabled = True
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, BillCol.��Ŀ) <> "" Then Bill.TextMatrix(i, Bill.Cols - 1) = "��"
    Next
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        
    If mbytInState = 0 And gblnFromDr And mobjBill.Details.Count = 0 Then
        cbo������.SetFocus
    ElseIf mbytUseType = 1 And mbytInState = 0 Then
        Bill.SetFocus
    ElseIf gbytBilling = 2 Then
        cboNO.SetFocus
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mbytInState = 3 Then
        cmdOK.SetFocus
    Else
        Bill.SetFocus
    End If
    Call SetDrawDrugDeptVisible
    Call SetDrawDrugDeptEnabled
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',:��;��?��|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub



Private Sub Form_Load()
    Dim i As Long, tmpBill As ExpenseBill
    
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    glngFormW = 12000: glngFormH = 7290
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    RestoreWinState Me, App.ProductName, mbytInState
    sta.Visible = True
    
    gblnOK = False: mblnEnterCell = True
    mlngPreUnit = -1
    
    
    
    '��ʼ����������
    Set mobjBill = New ExpenseBill
    
    Call zlLoadDrawDeptData(mbytUseType, mlngDeptID)
    
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Then
        If Not InitData Then Unload Me: Exit Sub
    Else
        If Init�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mstrPrivs, mbytUseType, mlngDeptID) = False Then
            Exit Sub
        End If
    End If
    mstrUnitIDs = GetUserUnits
    Set mcolPatiInfo = New Collection
    
    
    Call InitFace
    Call NewBill
    
    
    If mbytInState <> 0 Then '��ʾ�����������ʵ���(1,2,3)
        If Not ReadBill(mstrInNO, (mbytInState = 3)) Then Unload Me: Exit Sub
        cboNO.Text = mstrInNO
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then cbo������.Visible = False: lbl������.Visible = False
    Else '����
        mstrҩƷ�۸�ȼ� = gstrҩƷ�۸�ȼ�
        mstr���ļ۸�ȼ� = gstr���ļ۸�ȼ�
        mstr��ͨ�۸�ȼ� = gstr��ͨ�۸�ȼ�
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸ĵ���
            Set mobjBill = ImportBill(mstrInNO, True, Me, True, gblnסԺ��λ, True, , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
            If mobjBill.NO = "" Then
                MsgBox "��ȡ����ʧ�ܡ�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                Call ReCalcInsure '���¼���ͳ����
                
                cboNO.Text = mobjBill.NO '��ʾԭ����
                
                Bill.ClearBill
                Bill.Rows = mobjBill.Details.Count + 1
                Call InitBillColumnColor
                
                txtDate.Text = Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss")
                chk�Ӱ�.Value = mobjBill.�Ӱ��־

                Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
                                
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                If gintPriceGradeStartType < 2 Then
                    If gbln��������ۿ� Then Call CalcMoneys
                Else
                    'ÿһ�и��ݼ۸�ȼ�����۸�
                    For i = 1 To mobjBill.Details.Count
                        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, _
                            mobjBill.Details(i).����ID, mobjBill.Details(i).��ҳID, mobjBill.Details(i).ҽ�Ƹ���, _
                            mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                        Call CalcMoneys(i)
                    Next
                End If
                Call ShowDetails
                Call SetIntureColor
            End If
        Else
            If mbytUseType = 1 And mlng����ID <> 0 Then
                Bill.Row = 1: Bill.Col = BillCol.����
                Bill.Text = "-" & mlng����ID
                Call Bill_KeyDown(13, 0, False)
                Bill.Text = ""
                Bill.TxtVisible = False
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngCancelW As Long
        
    On Error Resume Next
    
    Bill.Height = Me.ScaleHeight - Bill.Top - sta.Height - fraAppend.Height - picAppend.Height - IIf(fraDrawDept.Visible, fraDrawDept.Height, 0) + 75
    Bill.Left = 0: Bill.Width = Me.ScaleWidth
    
    If chkCancel.Visible Or lblFlag.Visible Then lngCancelW = chkCancel.Width
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    chkCancel.Left = fraTitle.Width - chkCancel.Width - 60
    lblFlag.Left = chkCancel.Left + (chkCancel.Width - lblFlag.Width) / 2
    
    cboNO.Left = fraTitle.Width - lngCancelW - 60 - cboNO.Width - 30
    lblNO.Left = cboNO.Left - lblNO.Width - 45
        
    fraAppend.Top = Bill.Top + Bill.Height - 75
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    
    fraDrawDept.Top = fraAppend.Top + fraAppend.Height - 150
    fraDrawDept.Width = Me.ScaleWidth - fraDrawDept.Left
    
    
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
    
    If cbo������.Container Is picUnit Then
        cbo��������.Left = lblDate.Left - cbo��������.Width - 300
        lbl��������.Left = cbo��������.Left - lbl��������.Width - 45
    Else
        cbo������.Left = lblDate.Left - cbo������.Width - 300
        lbl������.Left = cbo������.Left - lbl������.Width - 45
    End If
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 500
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mbytInState
    
    mbytInState = Empty
    mstrInNO = Empty
    mblnNOMoved = False '�����˳������,����Ӱ���������
    
    mlngDelRow = 0
    mlngUnitID = Empty
    mstrTime = ""
    mblnDelete = False
    gbytBilling = 0
    mbytUseType = 0
    mlngDeptID = 0
    mlng����ID = 0
    mstr����IDs = ""
    
    mlngҩƷ���ID = 0
    mlng�������ID = 0
    
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Set mrsWarn = Nothing
    Set mrsMedPayMode = Nothing
    
    If Not OS.IsDesinMode Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwPati.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwPati.SortOrder = lvwDescending
    Else
        lvwPati.SortOrder = lvwAscending
    End If
    lvwPati.Sorted = True
    intIdx = ColumnHeader.Index
        
    If Not lvwPati.SelectedItem Is Nothing Then lvwPati.SelectedItem.EnsureVisible
End Sub

Private Sub lvwPati_DblClick()
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    Bill.Text = "-" & Mid(lvwPati.SelectedItem.Key, 2)
    lvwPati.Visible = False
    Bill.SetFocus
    Call zlCommFun.PressKey(13)
End Sub

Private Sub lvwPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then lvwPati_DblClick
End Sub

Private Sub lvwPati_LostFocus()
    lvwPati.Visible = False
End Sub

 

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If gbln�����л� Then    '35242
        If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
            '�л����������ƥ�䷽ʽ
            Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            If Panel.Key = "PY" Then
                sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Else
                sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            End If
            zlDatabase.SetPara "���뷽ʽ", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
            gbytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))
        End If
    End If
End Sub

Private Sub tmrStatuPati_Timer()
  If picStatuPancl.Visible Then Call MoveStatuPatiInfor
 
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim tmpBill As New ExpenseBill
    Dim i As Long, strSQL As String
    Dim strInfo As String
    Dim Curdate As Date     '��������ǰʱ��
    
    On Error GoTo errH
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii)
    Else
        txtIn.Text = GetFullNO(txtIn.Text, 14)
        
        '�������۲���Ȩ��
        strInfo = Check���۲���(txtIn.Text, mstrPrivsOpt)
        If strInfo <> "" Then
            MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set tmpBill = ImportBill(txtIn.Text, True, Me, False, gblnסԺ��λ, False, mlngUnitID, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
        If tmpBill.NO = "" Then
            MsgBox "��ȡ����ʧ�ܡ�", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtIn: txtIn.SetFocus
        Else
            '�����޸ļ���ʾ
            Screen.MousePointer = 11
            Set mobjBill = New ExpenseBill
            Set mobjBill = tmpBill
            
            Call ReCalcInsure '���¼���ͳ����
            
            Curdate = zlDatabase.Currentdate
            mobjBill.NO = cboNO.Text
            mobjBill.�Ǽ�ʱ�� = Curdate
            mobjBill.����Ա��� = UserInfo.���
            mobjBill.����Ա���� = UserInfo.����
            mobjBill.�Ӱ��־ = chk�Ӱ�.Value
            If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then mobjBill.������ = ""
            
            'ȡ��ǰʱ��
            txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
            
            Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
            
            Bill.Redraw = False
            Bill.ClearBill
            Bill.Rows = mobjBill.Details.Count + 1
            
            Call InitBillColumnColor
            
            If gintPriceGradeStartType < 2 Then
                Call CalcMoneys
            Else
                'ÿһ�и��ݼ۸�ȼ�����۸�
                For i = 1 To mobjBill.Details.Count
                    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, _
                        mobjBill.Details(i).����ID, mobjBill.Details(i).��ҳID, mobjBill.Details(i).ҽ�Ƹ���, _
                        mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                    Call CalcMoneys(i)
                Next
            End If
            Call ShowDetails
            Call SetIntureColor
            
            Bill.Redraw = True
            chkIn.Value = 0
            Screen.MousePointer = 0
            
            '���ʷ��౨��
            mstrWarn = ""
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboNO_GotFocus()
    zlControl.TxtSelAll cboNO
    
    If gbytBilling = 2 Or chkCancel.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, strOper As String
    Dim vDate As Date, intTmp As Integer, str����IDs As String
    Dim strInfo As String, intInsure As Integer
    Dim strInsure As String, arrInsure As Variant
    Dim i As Long, blnFlagPrint As Boolean
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 14)
        
        If chkCancel.Value = 1 Then
            '����
            
            '�Ƿ���ת������ݱ���
            If gbytBilling = 0 Then
                If zlDatabase.NOMoved("סԺ���ü�¼", cboNO.Text, , 2, Me.Caption) Then
                    If Not ReturnMovedExes(cboNO.Text, 2, Me.Caption) Then Exit Sub
                    mblnNOMoved = False
                End If
            End If
            
            '�����˻���ȫ��˵Ĳ���������
            If Not BillIdentical(cboNO.Text) Then
                MsgBox "�����а������ݲ�ȫ����˻�ֶ����˵����ݣ����������������ʡ�" & _
                    vbCrLf & "���˻ع��������˳���Ӧ�ĵ������ݣ�Ȼ�������ʡ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '����Ȩ��
            If Not ReadBillInfo(2, cboNO.Text, 2, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If mbytUseType = 0 And InStr(mstrPrivs, ";���в���Ա;") <= 0 Then
                If UserInfo.���� <> strOper Then
                    MsgBox "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ�������!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(5, strOper, vDate, "����", cboNO.Text) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '��Ŀ����Ȩ��
            If mbytUseType = 0 Or mbytUseType = 1 Then
                If Not CheckDelPriv(cboNO.Text, mstrPrivsOpt) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
        
            '���۲���Ȩ��
            strInfo = Check���۲���(cboNO.Text, mstrPrivsOpt)
            If strInfo <> "" Then
                MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '�Ƿ���ִ��
            intTmp = BillCanDelete(cboNO.Text, 2, True, , mstrPrivsOpt, blnFlagPrint)
            If intTmp <> 0 Then
                Select Case intTmp
                    Case 1 '�õ��ݲ�����
                        MsgBox "ָ�������е����ݲ�����,������û������շ���Ŀ������Ȩ�ޣ�", vbInformation, gstrSysName
                    Case 2 '�Ѿ�ȫ����ȫִ��
                        MsgBox "ָ�������е������Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
                    Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                        MsgBox "ָ�������е�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If blnFlagPrint Then
                If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
                        
            '��Ժ���˲���Ȩ���ж�
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "����", , str����IDs) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
                        
            '�Ƿ��ѽ����ж�
            intTmp = HaveBilling(2, cboNO.Text, False)
            If intTmp <> 0 Then
                Call GetBillInsures(strInsure, cboNO.Text, , , True)
                If strInsure <> "" Then
                    arrInsure = Split(strInsure, ",")
                    For i = 0 To UBound(arrInsure)
                        If arrInsure(i) <> 0 Then
                            If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , arrInsure(i)) Then
                                'ҽ�����˵ĵ���,�̶�Ϊ�ѽ��ʵĽ�ֹ����
                                If intTmp = 1 Then
                                    MsgBox "��ҽ�����ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                                Else
                                    MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbInformation, gstrSysName
                                End If
                            End If
                        Else
                            Select Case gbytBillOpt
                                Case 0
                                Case 1
                                    If MsgBox("�ü��ʵ����д����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                                    End If
                                Case 2
                                    If intTmp = 1 Then
                                        MsgBox "�ü��ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                                    Else
                                        MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbInformation, gstrSysName
                                    End If
                            End Select
                        End If
                    Next
                End If
            End If
            
            intInsure = BillExistInsure(cboNO.Text) '�ж��Ƿ���ҽ�����˼ǵ���,���ʱ�������ֻҪ��ҽ������
            'ҽ�����ʲ�����Ը�����¼��������
            If intInsure <> 0 Then
                If CheckNONegative(cboNO.Text) Then
                    MsgBox "�õ��ݴ��ڸ������ʼ�¼,���������ҽ�����ʲ�����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
                        
            '�Ƿ������������¼
            If CheckRecalcRecord(cboNO.Text) Then
                MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
                    "����ǰ�밴�ѱ�������ã������˽����������ʵ��ݵĴ����Żݽ�", vbInformation, Me.Caption
            End If
        ElseIf mobjBill.Details.Count = 0 Then
            '���ʻ��۵�(�������)
            
            If Not BillExistMoney(cboNO.Text, 2) Then
                MsgBox "�õ��ݷ����Ѿ�ȫ�����ʻ򵥾ݲ����ڣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '��Ժ���˲���Ȩ���ж�
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "���", , str����IDs) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        End If
        
        mstr����IDs = str����IDs
        
        If chkCancel.Value = 1 Then '��ȡ���ʵ�
            blnRead = ReadBill(cboNO.Text, True)
        ElseIf mobjBill.Details.Count = 0 Then '��ȡסԺ���۵�
            blnRead = ReadBill(cboNO.Text, False)
        End If
        
        If blnRead Then
            
            mstrInNO = cboNO.Text 'ȷ��ʱ��mstrInNOΪ׼
            If chkCancel.Value = 0 Then '���۵�
                Bill.Active = False
            Else '����
                'Call SetDisible 'cboNO�ڻ�ȡ�����unLock
                Bill.Active = True
            End If
            cmdOK.SetFocus
        Else
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim lngColor As Long, i As Long
    Dim lngRow As Long, lngCol As Long
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        'Bill.RemoveMSFItem Row'������AllowAddRow����
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
     '�Զ�������һ�еĲ�����Ϣ
    If Row > 1 Then
        Bill.Redraw = False
        lngRow = Bill.Row: lngCol = Bill.Col
        
        Bill.Col = BillCol.����: Bill.Row = Row - 1
        lngColor = Bill.MsfObj.CellForeColor
        
        Bill.Col = BillCol.����: Bill.Row = Row
        Bill.MsfObj.CellForeColor = lngColor
        
        Bill.Row = lngRow: Bill.Col = lngCol
        Bill.Redraw = True
        
        Bill.TextMatrix(Row, BillCol.����) = Bill.TextMatrix(Row - 1, BillCol.����)
        Bill.TextMatrix(Row, BillCol.�Ա�) = Bill.TextMatrix(Row - 1, BillCol.�Ա�)
        Bill.TextMatrix(Row, BillCol.����) = Bill.TextMatrix(Row - 1, BillCol.����)
        Bill.TextMatrix(Row, BillCol.����) = Bill.TextMatrix(Row - 1, BillCol.����)
        'ȱʡ�ѱ�
        Bill.TextMatrix(Row, BillCol.�ѱ�) = Bill.TextMatrix(Row - 1, BillCol.�ѱ�)
    End If
    
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        If mbytInState <> 2 Then
            .ColData(BillCol.����) = BillColType.Text      '�����λ��������,��ı�
            .ColData(BillCol.���) = IIf(gbln�շ���� And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            .ColData(BillCol.��Ŀ) = BillColType.CommandButton
            .ColData(BillCol.����) = BillColType.UnFocus '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = BillColType.UnFocus '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.��־) = BillColType.UnFocus '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
        End If
        
        '����б༭����������ɫ
        
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
    End With
    
   
    
    On Error Resume Next
    Bill.Text = "": Bill.SetFocus
    
    Set mrsInfo = New ADODB.Recordset: mstrUseMoney = "" '��ǰ������Ϣ��Ч(���ⱻ����������)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo��������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo������.ListIndex >= 0 Then lngҽ��ID = cbo������.ItemData(cbo������.ListIndex)
    If mrs�������� Is Nothing Then Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lngҽ��ID)
    
    If zlSelectDept(Me, mlngModule, cbo��������, mrs��������, cbo��������.Text) = False Then
        Call Beep: mobjBill.��������ID = 0
        KeyAscii = 0: Exit Sub
    End If
    mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Exit Sub




'
'
'
'
'
'    Dim lngIdx As Long
'
'    If KeyAscii >= 32 And Not cbo��������.Locked Then
'        lngIdx = zlControl.CboMatchIndex(cbo��������.hwnd, KeyAscii)
'        If lngIdx = -1 And cbo��������.ListCount > 0 Then lngIdx = 0
'        cbo��������.ListIndex = lngIdx
'
'    ElseIf KeyAscii = 13 Then
'        If cbo��������.ListIndex = -1 Then
'            Beep
'        Else
'            mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
'            Call zlcommfun.PressKey(vbKeyTab)
'        End If
'    End If
End Sub
Private Function isCheck������Exists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo������.ListCount - 1
        If zlStr.NeedName(cbo������.List(i)) = str���� Then
            If blnLocateItem Then cbo������.ListIndex = i
            isCheck������Exists = True
            Exit Function
        End If
    Next
End Function
Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, rsTemp As ADODB.Recordset, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    
    If KeyAscii = 13 Then
        If cbo������.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cbo������.Text)
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then Call zlControl.CboSetIndex(cbo������.hWnd, -1)
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1
            strFilter = IIf(gbln��ʿ, "��Ա����<>''", "��Ա����<>'��ʿ'")
            '���˺�:22383
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs������)
            Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
            Dim strCompents As String 'ƥ�䴮
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs������.Filter = strFilter: iCount = 0
            With mrs������
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs������.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If Nvl(!���) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!���)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!����)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(Nvl(!���)) Like strText & "*" Then
                            If isCheck������Exists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!����)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                        End If
                    End Select
                    mrs������.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
            '���˺�:ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck������Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                Select Case intInputType
                Case 0 '����ȫ����
                    rsTemp.Sort = "���"
                Case 1 '����ȫƴ��
                    rsTemp.Sort = "����"
                Case Else
                    '����ѡ������
                    If gbyt��������ʾ = 1 Then '����
                        rsTemp.Sort = "����"
                    Else
                        rsTemp.Sort = "���"
                    End If
                End Select
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo������, rsTemp, True, "", "ȱʡ,ְ��,���ȼ���", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheck������Exists(Nvl(rsReturn!����), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo������: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
            
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.������ = ""
            If gblnFromDr Then Exit Sub
        Else
            mobjBill.������ = zlStr.NeedName(cbo������.Text)
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF3
            If chkIn.Visible And chkIn.Enabled Then chkIn.Value = IIf(chkIn.Value = 1, 0, 1)
        Case vbKeyF6 '�����ǰ��������,�����µ�״̬
            If mbytInState = 0 Then
                If MsgBox("ȷʵҪ�����ǰ�����е�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If

                If chkCancel.Value = Checked Then '�˾ݵ�״̬
                    Call ClearRows: Call Bill.ClearBill
                    chkCancel.Value = Unchecked
                    Call NewBill
                    Call SetDisible(True)
                    If Bill.Enabled Then Bill.SetFocus
                ElseIf Bill.Active Then '�������뵥��״̬
                    Call ClearRows: Call Bill.ClearBill
                    Call NewBill   '����ԭ���ݺ�
                    If Bill.Enabled Then Bill.SetFocus
                End If
            End If
        Case vbKeyF7 '�л����뷨
            If gbln�����л� Then
                If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                    If sta.Panels("WB").Bevel = sbrRaised Then
                        Call sta_PanelClick(sta.Panels("WB"))
                    Else
                        Call sta_PanelClick(sta.Panels("PY"))
                    End If
                End If
            End If
        Case vbKeyF8 '��(�Զ������¼�)
            If chkCancel.Visible And chkCancel.Enabled Then chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
        Case vbKeyA, vbKeyR
            'ȫѡ��ȫ��
            If Shift = vbCtrlMask Then
                If KeyCode = vbKeyA And cmdSelALL.Visible And cmdSelALL.Enabled Then
                    Call cmdSelAll_Click
                ElseIf KeyCode = vbKeyR And cmdClear.Visible And cmdClear.Enabled Then
                    Call cmdClear_Click
                End If
            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then Call LocateNewRow
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            ElseIf lvwPati.Visible Then
                lvwPati.Visible = False
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim Curdate As Date     '��������ǰʱ��
    On Error GoTo errH
    
    Curdate = zlDatabase.Currentdate
    '�Զ�ʶ��Ӱ�
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(Curdate) Then chk�Ӱ�.Value = Checked
    End If
    
    '��ͬҩ��ҩƷ�����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '��ѡ�ѱ�
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ� Order by ����"
    Set mrsLevel = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsLevel, strSQL, Me.Caption)
    If mrsLevel.EOF Then
        MsgBox "û�г�ʼ���ѱ����ȵ��ѱ�����н������ã�", vbInformation, gstrSysName
        Exit Function
    End If
        
    If Init�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mstrPrivs, mbytUseType, mlngDeptID) = False Then
        Exit Function
    End If
        
    If gstr�շ���� = "" Then
        strSQL = "Select ����,���� as ��� from �շ���Ŀ��� Where ����<>'1' Order by ���"
    Else
        strSQL = "" & _
        "   Select /*+ RULE */   A.����,A.���� as ��� " & _
        "   From �շ���Ŀ��� A," & _
        "          (Select Column_Value From Table(Cast(f_str2list([1]) As Zltools.t_strlist))) J " & _
        "   Where A.����=J. Column_Value " & _
        "   Order by ���"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(gstr�շ����, "'", ""))
    
    If mrsClass.EOF Then
        MsgBox "û�����ÿ��õ��շ����,�����ڱ��ز��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    '��ֻ��һ�ֿ�ѡ�շ����ʱ,�����û�ѡ��
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(gstr�շ����, "'5'") > 0 Or InStr(gstr�շ����, "'6'") > 0 Or InStr(gstr�շ����, "'7'") > 0 Or gstr�շ���� = "" Then
        mlngҩƷ���ID = ExistIOClass(10)
        If mlngҩƷ���ID = 0 Then
            MsgBox "����ȷ���������ݵ�������,���ȵ�ҩƷ���������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr�շ����, "'4'") > 0 Or gstr�շ���� = "" Then
        mlng�������ID = ExistIOClass(42)
        If mlng�������ID = 0 Then
            MsgBox "����ȷ�����ĵ��ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'ִ�в���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        "From ���ű� A,��������˵�� B " & _
        "Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "and B.����ID=A.ID and B.������� IN(2,3) " & _
        "Order by B.�������,A.����"
    Set mrsUnit = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsUnit, strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbytInState = 0 Then Set mrsWarn = GetUnitWarn
    Set mrsInfo = New ADODB.Recordset
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str��� As String, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'���ܣ���ȡ����������ͬ�����Ŀ��ִ�п���ID
    Dim i As Long
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).�շ���� = str��� _
            And mobjBill.Details(i).ִ�в���ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                Exit Function
            End If
        End If
    Next
    
    '�������������,��ȡ��������������ƥ���ִ�п���
    If str��� = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).ִ�в���ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional ByVal lng����ID As Long, Optional ByVal int��Դ As Integer, Optional blnEnter As Boolean)
'���ܣ����ݵ��������������б������
'������blnEnter=�Ƿ񰴽�����д���,����ִ�п��ұ��ֲ���
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strSQL As String, strIDs As String
    Dim bln��ʿ As Boolean, strTmp As String
    Dim lng����ID As Long, lngListIndex As Long
    
    
    On Error GoTo errHandle
    
    Bill.Clear '����б������
    Select Case Bill.TextMatrix(0, lngCol)
        Case "�ѱ�"
            mrsLevel.Filter = adFilterNone
            If mrsLevel.RecordCount <> 0 Then
                For i = 1 To mrsLevel.RecordCount
                    Bill.AddItem mrsLevel!���� & "-" & mrsLevel!����
                    mrsLevel.MoveNext
                Next
            End If
             Bill.cboStyle = DropOlnyDown
        Case "���"
            Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
        
            mrsClass.Filter = adFilterNone
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                For i = 1 To mrsClass.RecordCount
                    If Not (bln��ʿ And InStr(",E,M,4,", mrsClass!����) = 0) Then
                        Bill.AddItem Bill.ListCount + 1 & "-" & mrsClass!���
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!����)  '����������ASCII��
                    End If
                    mrsClass.MoveNext
                Next
            End If
             Bill.cboStyle = DropOlnyDown
        Case "ִ�п���"
             Bill.cboStyle = DropDownAndEdit
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .�շ����) > 0 And .�շ���� <> "" Then
                        Call GetWorkUnit(.�շ�ϸĿID, .�շ����)
                        If mrsWork.RecordCount > 0 Then
                            'ȡ��һ��ҩ��ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                lng����ID = GetLastDeptID(.�շ����, lngRow, Mid(strIDs, 2))
                            End If
                            If lng����ID = 0 Then lng����ID = .ִ�в���ID
                                                    
                            'ȷ����ǰ�е�ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!���� & "-", "") & mrsWork!����
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        
                        If int��Դ = 0 Then int��Դ = Get������Դ(lngRow)
                        
                        If lng����ID = 0 Then
                            lng����ID = .����ID
                            If lng����ID = 0 Then lng����ID = Get��������ID
                        End If
                        
                        lng����ID = .����ID
                        If lng����ID = 0 Then lng����ID = Get����ID(lng����ID)
                        If lng����ID = 0 Then lng����ID = lng����ID
                        
                        '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                        Select Case .Detail.ִ�п���
                            Case 0 '����ȷ
                                mrsUnit.Filter = 0
                            Case 1 '���˿���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 2 '���˲���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 3 '����Ա����
                                mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                            Case 4 'ָ������
                                strSQL = "" & _
                                "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                                "   From �շ�ִ�п��� A,���ű� C" & _
                                "   Where A.�շ�ϸĿID=[1]��And A.ִ�п���ID+0=C.ID " & _
                                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                                " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, int��Դ, lng����ID)
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!ִ�п���ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .ִ�в���ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                                End If
                            Case 5 'Ժ��ִ��(Ԥ��,������δ��)
                            Case 6 '�����˿���
                               mrsUnit.Filter = "ID=" & Get��������ID & " Or ID=" & .ִ�в���ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                        If Not mrsUnit.EOF Then
                            mrsUnit.MoveFirst: lngListIndex = -1
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                                '���˺�:28947
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                                'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.ListCount - 1) = mrsUnit!ID
                                    
                                    '����ȱʡִ�п���
                                    If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng����ID Then lngListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '����һ�з�ҩƷ��ͬ
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).ִ�в���ID And mobjBill.Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).�շ����) = 0 Then
                                                lngListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng����ID And Bill.ListIndex = -1 Then
                                               lngListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                            '28378 ����Ҫ����Bill_CboClick�¼�,���,���ܽ�Bill.Listindex����ѭ��(��Ϊ�¼��а����˶�mrsUnit�Ĺ��˴�������ɼ�¼������)
                            If lngListIndex >= 0 Then Bill.ListIndex = lngListIndex
                        End If
                            
                        If Not blnEnter And .Detail.ִ�п��� = 4 Then    'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.����ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        If Bill.ListIndex = -1 Then '���û����ȡ���е�ִ�п���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .ִ�в���ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .ִ�в���ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
'���ܣ����ݱ�Ҫ��ɵĹ������ý��沼��
    Dim arrHead() As String, i As Long
    
    '���õ��ݱ��ʽ
    With Bill
        .LocateCol = BillCol.���� 'ȱʡ��λ��������
        .PrimaryCol = BillCol.���� '������Ϊ����
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        .MsfObj.MergeCells = flexMergeRestrictRows
        .MsfObj.MergeRow(0) = True
                
        If mbytInState = 0 And gbytBilling <> 2 Then
            .ColData(BillCol.����) = BillColType.Text  '�������룡����
            .ColData(BillCol.�Ա�) = BillColType.UnFocus   '�Ա�����
            .ColData(BillCol.����) = BillColType.UnFocus  '��������
            .ColData(BillCol.����) = BillColType.UnFocus  '��������
            .ColData(BillCol.�ѱ�) = BillColType.UnFocus  '�ѱ�����
            
            .ColData(BillCol.���) = IIf(gbln�շ���� And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            
            .ColData(BillCol.��Ŀ) = 1  '��Ŀ����,��Ť��ѡ
            '���˺�:27990 2010-02-22 17:00:04
            .ColData(BillCol.��Ʒ��) = BillColType.UnFocus  '��Ʒ������
            .ColData(BillCol.���) = BillColType.UnFocus  '�������
            .ColData(BillCol.��λ) = BillColType.UnFocus  '��λ����
            .ColData(BillCol.����) = BillColType.UnFocus '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = BillColType.Text   '��/������
            .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.Ӧ�ս��) = BillColType.UnFocus  'Ӧ�ս������
            .ColData(BillCol.ʵ�ս��) = BillColType.UnFocus  'ʵ�ս������
            .ColData(BillCol.ִ�п���) = BillColType.ComboBox   'Ĭ��ȡ�������һ���һ����
            .ColData(BillCol.��־) = BillColType.UnFocus  '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
            .ColData(BillCol.����) = BillColType.UnFocus
        End If
        
                
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
    
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    If gTy_System_Para.bytҩƷ������ʾ <> 2 Then
        '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        Bill.ColWidth(BillCol.��Ʒ��) = 0
    Else
        If Bill.ColWidth(BillCol.��Ʒ��) = 0 Then
             Bill.ColWidth(BillCol.��Ʒ��) = GetOrigColWidth(BillCol.��Ʒ��)
        End If
    End If
    '��ȡ����ƥ�䷽ʽ
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln�����л� '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln�����л�
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If
    
    '����
    Select Case gbytBilling
        Case 0
            lblTitle.Caption = gstrUnitName & "סԺ���ʱ�"
        Case 1
            lblTitle.Caption = gstrUnitName & "סԺ���ʱ�(����)"
        Case 2
            lblTitle.Caption = gstrUnitName & "סԺ���ʱ�(���)"
    End Select
    
    txt����.Text = gstrDec: txt����.Text = gstrDec
    
    If mbytInState = 0 And (gbytBilling = 0 Or gbytBilling = 1) Then
        chkIn.Visible = True
        txtIn.Visible = True
    Else
        txt����.Left = Val(txt����.Tag) - chkIn.Width - txtIn.Width
        lbl����.Left = txt����.Left - lbl����.Width - 45
        txt����.Left = Val(txt����.Tag) - chkIn.Width - txtIn.Width
        lbl����.Left = txt����.Left - lbl����.Width - 45
    End If
    
    Select Case mbytInState
        Case 0 'ִ��
            '55380
            If mstrInNO <> "" Or _
                (InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
                And InStr(mstrPrivsOpt, ";��������;") = 0 _
                And InStr(mstrPrivsOpt, ";��������;") = 0) Then
                chkCancel.Visible = False
            End If
            Select Case gbytBilling
                Case 0, 1 'ִ�м��ʡ�����
                    Call SetShowCol
                Case 2 'ִ�����
                    Call SetDisible
                    cboNO.Locked = False
                    picUnit.Enabled = False
                    fraAppend.Enabled = False
            End Select
        Case 1 '����
            Call SetDisible
            chkCancel.Visible = False
            If mblnDelete Then lblFlag.Visible = True
            fraTitle.Enabled = False
            picUnit.Enabled = False
            fraAppend.Enabled = False
            
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
        Case 2 '����
            Call SetDisible
            txtDate.Enabled = True
            chkCancel.Visible = False
            fraTitle.Enabled = False
            picUnit.Enabled = False
            
        Case 3 '����
            Call SetDisible
            chkCancel.Visible = False
            fraTitle.Enabled = False
            picUnit.Enabled = False
            fraAppend.Enabled = False
            
            Call ShowDeleteCol(True)
            Bill.Active = True
    End Select
    
    '�������������뿪����λ��
    If gblnFromDr Then
        Call ExChangeLocate(cbo��������, cbo������)
        Call ExChangeLocate(lbl��������, lbl������)
        cbo��������.TabStop = False
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'��������Ϊ�����޸�״̬
    cboNO.Locked = Not bln
    cbo��������.Locked = Not bln
    chk�Ӱ�.Enabled = bln
    cboBaby.Enabled = bln
    cbo������.Locked = Not bln
    txtDate.Enabled = bln
    Bill.Active = bln
End Sub

Private Function GetPatient(ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '���ܣ���ȡ������Ϣ
    '������blnCard=�Ƿ���￨ˢ��
    '����:blnOutMsg-�Ѿ���ʾ,�����ⲻ���ٵ�Msgbox
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String, bln���в��� As Boolean
    Dim strPati As String, strIF As String, strWhere As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsOutSel As ADODB.Recordset
    mstrUseMoney = ""
    On Error GoTo errH
        
    'a.�Ƿ����ǿ�Ƽ���Ȩ��
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
    Else
        strIF = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
    End If
    
    'b.�Ƿ���Լ����в�������
    bln���в��� = True
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";���в���;") <= 0 Then
        bln���в��� = False
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.��ǰ����ID+0=[3]"
        Else
            strIF = strIF & " And B.��ǰ����ID+0 IN (Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.�Ƿ����۲��˼���Ȩ��
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.��������,0)=0"
    End If
    
    strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "   A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬," & _
            "   nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            "   B.����,Nvl(B.��������,0) as ��������,B.��������,B.��˱�־" & _
            " From ������Ϣ A,������ҳ B,������� X" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
            "       And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2 And A.ͣ��ʱ�� is NULL " & strIF
            
    If blnCard Then '���￨��
        strInput = UCase(strInput)
        strWhere = strWhere & " And A.���￨��=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "/" Then   '��λ��
        '41654 And IsNumeric(Mid(strInput, 2))
        strInput = Mid(strInput, 2)
        If mlngUnitID = 0 Then '������ȷ��������ͨ������ȷ������
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "   A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬," & _
            "   nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            "   B.����,Nvl(B.��������,0) as ��������,B.��������,B.��˱�־" & _
            "   From ������Ϣ A,������ҳ B,��λ״����¼ C,������� X" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
            " And Nvl(B.��ҳID,0)<>0 And A.����ID=C.����ID And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2 And A.ͣ��ʱ�� is NULL" & _
            " And C.����ID=[3] And C.����=[2] " & strIF
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(ҽ������)
        strWhere = strWhere & " And A.�����=[1]"
    Else '��������
        If zlSelectChargePatiFromInputName(Me, mstrPrivsOpt, strInput, bln���в���, mstrUnitIDs, gintOutDay, lng����ID, strErrMsg, Bill.TxtHwnd, Bill.RowHeight(Bill.Row)) = False Then
            If strErrMsg = "" Then blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
            MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
        End If
        strInput = "-" & lng����ID
        strWhere = strWhere & " And A.����ID=[1]"
    End If
    
    strSQL = strSQL & vbCrLf & strWhere
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput, mlngUnitID, mstrUnitIDs)
    
    If Not mrsInfo.EOF Then
        'ȥ����ҽ������ƥ����
        If zlPatiIS�����ѱ�Ŀ(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = True Then      '����:28725
            Set mrsInfo = New ADODB.Recordset
            blnOutMsg = True
            Exit Function
        End If
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), Val(Nvl(mrsInfo!��˱�־))) = False Then
            Set mrsInfo = New ADODB.Recordset
            blnOutMsg = True
            Exit Function
        End If
        sta.Panels(2) = ""
        If cbo��������.ListIndex <> -1 Then
            If mrsInfo!����ID <> cbo��������.ItemData(cbo��������.ListIndex) And mrsInfo!����ID <> cbo��������.ItemData(cbo��������.ListIndex) Then
                MsgBox "���ѣ���סԺ���˲�����""" & zlStr.NeedName(cbo��������.Text) & """��", vbInformation, gstrSysName
            End If
        End If
        
        '��ȡ�۸�ȼ�
        If gintPriceGradeStartType >= 2 Then
            Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, _
                Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), Nvl(mrsInfo!ҽ�Ƹ��ʽ), _
                mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
        End If
        
        '������Ϣ�������շ���Ŀÿ�ζ�ȡ,��Ϊÿ�еĲ��˿��ܲ�ͬ�����ҿ����޸�����������
        GetPatient = True
        Exit Function
    End If
    
    Set mrsInfo = New ADODB.Recordset
    
    If strWhere = "" Then Exit Function '������������ֱ���˳�
    
    'δ�ҵ����ˣ���Ҫ�Ըò��˵ľ��������Ϣ������ʾ
    strSQL = _
    " Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,a.��Ժ,B.��Ժ����,B.��Ժ����,X.�������,B.״̬, " & _
    "       nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,nvl(b.����,A.����) as ����,B.�ѱ�,Nvl(B.��������,0) as ��������,B.��������" & _
    " From ������Ϣ A,������ҳ B,������� X" & _
    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
    "   And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) and X.����(+)=1 and X.����(+)=2 And A.ͣ��ʱ�� is NULL " & strWhere
    
    Set rsOutSel = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput)
    If rsOutSel.EOF Then Exit Function
    
    '1.�������
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";���в���;") <= 0 Then
        If InStr(1, "," & mstrUnitIDs & ",", "," & Val(rsOutSel!����ID) & ",") = 0 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "�������㸺��Ĳ���,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    
    '2.���۲��˼��(�Ƿ����۲��˼���Ȩ��)
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        If Val(Nvl(rsOutSel!��������)) = 2 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��ΪסԺ���۲���,�㲻�߱���סԺ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        If Val(Nvl(rsOutSel!��������)) = 1 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��Ϊ�������۲���,�㲻�߱����������ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    Else
        If Val(Nvl(rsOutSel!��������)) <> 0 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��Ϊ" & IIf(Val(Nvl(rsOutSel!��������)) = 1, "����", "סԺ") & "���۲���,�㲻�߱��������סԺ ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    
    '124007
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strErrMsg = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) <> 0) Then
              
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                    strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                    strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) = 0) Then
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    Else
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����)) Then
            If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
            Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
            End If
        End If
    End If
    
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        blnOutMsg = True
        Exit Function
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'���ܣ���������¼���ָ���л������еĽ��
'������lngRow=ָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, j As Long, k As Long
    Dim blnExist As Boolean
    
    Dim strMainRows As String
    Dim bln��������ۿ� As Boolean
    
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Details.Count, lngRow)
        
        bln��������ۿ� = False
        If gbln��������ۿ� Then                    '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
            If mobjBill.Details(i).�������� > 0 Then    '����
                bln��������ۿ� = Not mobjBill.Details(mobjBill.Details(i).��������).Detail.���ηѱ�
                If bln��������ۿ� And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).��������      '��������һ�е�ʱ��
            Else
                If CheckItemHaveSub(i) Then                          '����������
                     bln��������ۿ� = Not mobjBill.Details(i).Detail.���ηѱ�
                     If bln��������ۿ� Then strMainRows = strMainRows & "," & i  'һҳ�����ж��������,�ȼ�¼�����к�,���������������ۿ�
                End If
            End If
        End If
                    
        Call CalcMoney(i, bln��������ۿ�)
    Next
    
    '������������,������bln��������ۿ۱���,��Ϊ�������������Ǵ������ʱ�Ѹı�
    If gbln��������ۿ� Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc��������ʵ��(Split(strMainRows, ",")(i))
        Next
    End If
    
    Set mcolMoneys = New BillInComes
    '�������ܷ�Ŀ
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '�����Ƿ��Ѿ��������������Ŀ,������ϼ�,��������
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).������ĿID = mobjBill.Details(i).InComes(j).������ĿID Then
                    blnExist = True: Exit For
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + mobjBill.Details(i).InComes(j).ʵ�ս��
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��
                End With
            End If
        Next
    Next
End Sub

Private Sub CalcMoney(lngRow As Long, Optional bln��������ۿ� As Boolean)
'���ܣ���������¼���ָ���еĽ��
'������lngRow=ָ����
'˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
'      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details(lngRow).InComes(1)
'      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
'      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strInfo As String, i As Long
    Dim intInsure As Integer, dblMoney As Double '�û�����ı�۽��
    
    Dim dblAllTime As Double, dbl�Ӱ�Ӽ��� As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        strPriceGrade = mstrҩƷ�۸�ȼ�
    ElseIf mobjBill.Details(lngRow).�շ���� = "4" Then
        strPriceGrade = mstr���ļ۸�ȼ�
    Else
        strPriceGrade = mstr��ͨ�۸�ȼ�
    End If
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).�շ�ϸĿID)
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where b.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    strSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,B.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        " And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID, strPriceGrade)
    If rsTmp.EOF Then
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        Exit Sub
    End If

    With mobjBill.Details(lngRow)
        '�Ȼ�ȡ����Ա��ǰ����ı�۽��
        If InStr(",5,6,7,", .�շ����) > 0 Or (.�շ���� = "4" And .Detail.��������) Then
            '����ҩƷʱ��(�����򲻷���)
            '��Ȼ�м�¼(�������Ŀʱ���ж�)
            dblAllTime = .���� * .����
            If gblnסԺ��λ And InStr(",5,6,7,", .�շ����) > 0 Then
                dblAllTime = dblAllTime * .Detail.סԺ��װ '���ʱ�۰��ۼ��������м���
            End If
            If dblAllTime <> 0 Or Not .Detail.��� Then
                Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                            Me.Caption, .�շ�ϸĿID, .ִ�в���ID, dblAllTime)
                If rsPrice.EOF Then
                    '��ȡ�۸�ʧ��
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        MsgBox "�� " & lngRow & " ��ҩƷ""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                    Else
                        MsgBox "�� " & lngRow & " ����������""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                    End If
                Else
                    strPrice = Nvl(rsPrice!Price) & "|||"
                    varPrice = Split(strPrice, "|")
                    dblMoney = Val(varPrice(0))
                    dblʣ������ = Val(varPrice(2))
                    
                    If dblʣ������ <> 0 And .Detail.��� Then
                        '����δ�ֽ����
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                            MsgBox "�� " & lngRow & " ��ʱ��ҩƷ""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                        Else
                            MsgBox "�� " & lngRow & " ��ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                        End If
                        dblMoney = 0
                    End If
                End If
            Else
                dblMoney = 0
            End If
        Else
            If .Detail.��� Then
                If .InComes.Count = 0 Then '��һ�μ�����ȡȱʡֵ
                    dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                Else                        '��ȡ����Ա��ǰ����ı�۽��
                    dblMoney = .InComes(1).��׼����
                    '����û�����ı�۲������۷�Χ����ȡȱʡֵ
                    If CheckScope(Val(Nvl(rsTmp!ԭ��)), Val(Nvl(rsTmp!�ּ�)), dblMoney) <> "" Then
                        dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                    End If
                End If
            End If
        End If
    End With
    
    '�����ԭ�м�¼
    Set mobjBill.Details(lngRow).InComes = New BillInComes
    
    '��д���з��ü�¼
    For i = 1 To rsTmp.RecordCount
        Set mobjBillIncome = New BillInCome
        With mobjBillIncome
            .������ĿID = rsTmp!������ĿID
            .������Ŀ = rsTmp!����
            .�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
            .ԭ�� = Val(Nvl(rsTmp!ԭ��))
            .�ּ� = Val(Nvl(rsTmp!�ּ�))
            
            If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                If gblnסԺ��λ Then
                    .��׼���� = Format(dblMoney * mobjBill.Details(lngRow).Detail.סԺ��װ, gstrFeePrecisionFmt)
                Else
                    .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                End If
            Else
                If mobjBill.Details(lngRow).Detail.��� Then
                    .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                Else
                    .��׼���� = Format(Nvl(rsTmp!�ּ�, 0), gstrFeePrecisionFmt)
                End If
            End If
            
            'Ӧ�ս��=���� * ���� * ����
            .Ӧ�ս�� = .��׼���� * IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����) * mobjBill.Details(lngRow).����
            
            '�������������ü���(����������Ŀ)
            If mobjBill.Details(lngRow).���ӱ�־ = 1 And mobjBill.Details(lngRow).�շ���� = "F" Then
                .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
            End If
            
            '�Ӱ�����ʼ���
            dbl�Ӱ�Ӽ��� = 0
            If mobjBill.�Ӱ��־ = 1 And mobjBill.Details(lngRow).Detail.�Ӱ�Ӽ� Then
                dbl�Ӱ�Ӽ��� = IIf(IsNull(rsTmp!�Ӱ�Ӽ���), 0, rsTmp!�Ӱ�Ӽ��� / 100)
                .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * dbl�Ӱ�Ӽ���
            End If
            
            .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
            dblAllTime = mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
            If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                If gblnסԺ��λ Then dblAllTime = dblAllTime * mobjBill.Details(lngRow).Detail.סԺ��װ
            End If
            
            If mobjBill.Details(lngRow).Detail.���ηѱ� Or bln��������ۿ� Or .Ӧ�ս�� = 0 Then
                .ʵ�ս�� = .Ӧ�ս��
            Else
                If .Ӧ�ս�� = 0 Then
                    .ʵ�ս�� = 0
                    mobjBill.Details(lngRow).�ѱ� = mobjBill.�ѱ�
                Else
                    'ҩƷ���ɱ��ۼ���,��������
                    .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.Details(lngRow).�ѱ�, .������ĿID, .Ӧ�ս��, _
                        mobjBill.Details(lngRow).�շ�ϸĿID, mobjBill.Details(lngRow).ִ�в���ID, dblAllTime, dbl�Ӱ�Ӽ���), gstrDec))   '��ǰ���˵ķѱ�
                End If
            End If
            
            '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
            intInsure = GetPatiInsure(mobjBill.Details(lngRow).����ID)
            If intInsure > 0 Then
                strInfo = gclsInsure.GetItemInsure(mobjBill.Details(lngRow).����ID, mobjBill.Details(lngRow).�շ�ϸĿID, .ʵ�ս��, False, intInsure, _
                    mobjBill.Details(lngRow).ժҪ & "||" & dblAllTime)
                If strInfo <> "" Then
                    mobjBill.Details(lngRow).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                    mobjBill.Details(lngRow).���մ���ID = Val(Split(strInfo, ";")(1))
                    .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                    mobjBill.Details(lngRow).���ձ��� = CStr(Split(strInfo, ";")(3))
                    
                    If UBound(Split(strInfo, ";")) >= 4 Then
                        If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).ժҪ = CStr(Split(strInfo, ";")(4))
                        If UBound(Split(strInfo, ";")) >= 5 Then
                            If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.���� = Split(strInfo, ";")(5)
                        End If
                    End If
                End If
            End If
            
            mobjBill.Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, , .ͳ����
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiInsure(Optional ByVal lng����ID As Long) As Integer
'����:�õ���������
'������lng����ID=����ʱ��ȡ��һ��ҽ�����˵�����
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If lng����ID <> 0 Then
            If mobjBill.Details(i).����ID = lng����ID Then
                GetPatiInsure = Val(mobjBill.Details(i).��ҩ����)
                Exit Function
            End If
        Else
            If Val(mobjBill.Details(i).��ҩ����) > 0 Then
                GetPatiInsure = Val(mobjBill.Details(i).��ҩ����)
                Exit Function
            End If
        End If
    Next
End Function

Private Function GetMultiInsures() As String
'����:�õ������а����Ķ����������
    Dim strInsure As String, i As Long
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).����ID <> 0 And Val(mobjBill.Details(i).��ҩ����) <> 0 Then
            If InStr(strInsure & ",", "," & Val(mobjBill.Details(i).��ҩ����) & ",") = 0 Then
                strInsure = strInsure & "," & Val(mobjBill.Details(i).��ҩ����)
            End If
        End If
    Next
    GetMultiInsures = Mid(strInsure, 2)
End Function

Private Sub ShowDetails(Optional lngRow As Long = 0)
'���ܣ�ˢ����ʾָ���л������е�����
'������lngRow=ָ����,Ϊ0��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long

    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
    
    txt����.Text = Format(GetBillTotal(mobjBill), gstrDec)
End Sub

Private Sub ShowDetail(lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ����
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim dbl���� As Double, cur��� As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '���������
    For i = 0 To Bill.Cols - 1
        '����ʱ�շ�������
        If i > 5 Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).�շ����)
    
    'ˢ�µ�����
    For i = 0 To Bill.Cols - 1
        If i = 0 Then Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).����
        If i = 1 Then Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).�Ա�
        If i = 2 Then Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).����
        Select Case Bill.TextMatrix(0, i)
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).����
            Case "�ѱ�"
                '������ݻ������Ŀֻ(��)��ʾ����
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).�ѱ�
            Case "���"
                '������ݻ������Ŀֻ(��)��ʾ����
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.�������
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "���"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���
            Case "��Ʒ��"   '���˺�:27990 2010-02-22 17:00:49
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.��Ʒ��
            Case "��λ"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 And gblnסԺ��λ Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.סԺ��λ
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���㵥λ
                End If
            Case "��"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����)
            Case "����"
                '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).����, 5)
            Case "����"
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                dbl���� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl���� = dbl���� + mobjBill.Details(lngRow).InComes(j).��׼����
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl����, gstrFeePrecisionFmt)
            Case "Ӧ�ս��"
                'Ӧ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ִ�п���"
                If mobjBill.Details(lngRow).ִ�в���ID <> 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).ִ�в���ID
                    If mrsUnit.RecordCount <> 0 Then
                        If mbytInState = 0 Then
                            Bill.TextMatrix(lngRow, i) = mrsUnit!���� & "-" & mrsUnit!����
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(lngRow, i) = mrsUnit!����
                        End If
                    Else
                        Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "��־"
                If mobjBill.Details(lngRow).�շ���� = "F" And mobjBill.Details(lngRow).���ӱ�־ = 1 Then
                    Bill.TextMatrix(lngRow, i) = "��"
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Integer = 0)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    'ȡ������ҩ�ĸ���
    intPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "7" And i <> lngRow Then
            intPay = mobjBill.Details(i).����
            Exit For
        End If
    Next
    If Detail.��� <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '���ø���RowData
                Bill.RowData(lngRow) = Asc(Detail.���)
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = IIf(Detail.�������� = 0, 1, Detail.��������)
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Details(bytParent).����
                End If
            Else
                
                If InStr(",5,6,7,", Detail.���) > 0 Then
                    dblTime = 0
                Else
                    dblTime = 1
                End If
            End If
            If bytParent <> 0 Then
                '���˺�:mobjBill.Details(bytParent).��ҩ���� ����:
                '����:
                mobjBill.Details.Add Detail, .ID, CInt(lngRow), bytParent, mobjBill.Details(bytParent).����ID, _
                mobjBill.Details(bytParent).��ҳID, mobjBill.Details(bytParent).����ID, _
                mobjBill.Details(bytParent).����ID, mobjBill.Details(bytParent).����, _
                mobjBill.Details(bytParent).�Ա�, mobjBill.Details(bytParent).����, mobjBill.Details(bytParent).סԺ��, _
                mobjBill.Details(bytParent).����, mobjBill.Details(bytParent).�ѱ�, mobjBill.Details(bytParent).��������, _
                .���, .���㵥λ, mobjBill.Details(bytParent).��ҩ����, intPay, dblTime, 0, lngDoUnit, tmpIncomes, mobjBill.Details(bytParent).���￨��, , mobjBill.Details(bytParent).������, _
                mobjBill.Details(bytParent).ҽ�Ƹ���
            Else
                mobjBill.Details.Add Detail, .ID, CInt(lngRow), bytParent, 0, 0, 0, 0, "", "", "", 0, 0, "", 0, _
                .���, .���㵥λ, "", intPay, dblTime, 0, lngDoUnit, tmpIncomes
            End If
        End With
    Else '��������Ѿ�����,���޸�
        
        If InStr(",5,6,7,", Detail.���) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .���� = intPay
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = dblTime
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select count(����ID) as NUM From �շѴ�����Ŀ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID)
    
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).�������� = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(ByVal lng��Ŀid As Long) As Details
'���ܣ�����һ���շ�ϸĿ�Ĵ�����Ŀ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
    Dim objDetail As New Detail
    
    Set GetSubDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    If lngMediCareNO > 0 Then
        strSQL = _
        "Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�,G.Ҫ������," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D.סԺ��װ) as סԺ��װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D.סԺ��λ) as סԺ��λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1,����֧����Ŀ G" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And A.ID=E.����ID(+)" & _
        "       And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "       And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And C.����ID=[1] And A.ID=G.�շ�ϸĿID(+) And G.����(+)=[2] " & _
        " Order by ����"
    Else
        strSQL = _
        " Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�,0 as Ҫ������," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D.סԺ��װ) as סԺ��װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D.סԺ��λ) as סԺ��λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And A.ID=E.����ID(+)" & _
        "       And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "       And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And C.����ID=[1]  " & _
        " Order by ����"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
            .���� = rsTmp!����
            .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
            .��� = Nvl(rsTmp!���)
            .סԺ��װ = Nvl(rsTmp!סԺ��װ, 1)
            .סԺ��λ = Nvl(rsTmp!סԺ��λ)
            .���㵥λ = Nvl(rsTmp!���㵥λ)
            .���� = Nvl(rsTmp!����, 0) = 1
            .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
            .��� = rsTmp!���
            .������� = rsTmp!�������
            .���� = rsTmp!����
            .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
            .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
            .������� = Nvl(rsTmp!�������, 0)
            .���д��� = Nvl(rsTmp!���д���, 0)
            .�������� = Nvl(rsTmp!��������, 1)
            .���� = Nvl(rsTmp!��������)
            .�������� = Nvl(rsTmp!��������, 0) = 1
            .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
            .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
            .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
            GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                .סԺ��װ, .סԺ��λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .�������, .����, .����ժҪ, .���д���, .��������, .��������, , , , , , .Ҫ������, , .��ҩ��̬, .��Ʒ��
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
'���ܣ�ɾ��ָ���շ���Ŀ��
'˵������ʱ����������е�ɾ��,��Ҫ�����������д�����ϵ����Ӧ�ĵ���
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).�������� > lngRow Then
            mobjBill.Details(i).�������� = mobjBill.Details(i).�������� - 1
        End If
        mobjBill.Details(i).��� = mobjBill.Details(i).��� - 1 '������кŶ�Ӧ
    Next
    mobjBill.Details.Remove lngRow
    If lngRow = 1 And mobjBill.Details.Count = 0 And Bill.Rows = 2 Then
        For i = 0 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill()
'���ܣ���ʼ��һ���µĵ���(�������)
    Dim Curdate As Date     '��������ǰʱ��
    
    '���ʷ��౨��
    mstrWarn = ""
    Set mrsInfo = New ADODB.Recordset
    Set mobjBill = New ExpenseBill
    Set mcolPatiInfo = New Collection

    mstrUseMoney = "": sta.Panels(3).Text = "": picStatuPancl.Visible = False: lblStatuPati.Caption = ""
    mcurModiMoney = 0
    mlngPreRow = 0
    cboNO.Text = ""
    
    Call LoadPatientBaby(cboBaby, 0, 0)
    Call cbo��������_Click
    
    Curdate = zlDatabase.Currentdate
    chk�Ӱ�.Value = IIf(OverTime(Curdate), 1, 0)
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    txt����.Text = gstrDec: txt����.Text = gstrDec: lbl����.Caption = "����"
    
    With mobjBill
        .�����־ = 2
        .�ಡ�˵� = True
        .������ = UserInfo.����
        .������ = zlStr.NeedName(cbo������.Text)
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
        .����ʱ�� = CDate(txtDate.Text)
        .�Ӱ��־ = chk�Ӱ�.Value
        .Ӥ���� = 0
        If cbo��������.ListIndex = -1 Then
            .��������ID = 0
        Else
            .��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        If cboDrawDept.ListIndex = -1 Then
            .��ҩ����ID = 0
        Else
            .��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
        End If
        
    End With
End Sub

Private Function SaveBill() As Boolean
'����:���浱ǰ����ļ��ʵ���(����סԺ���ʡ����ۡ�������ߵ��޸�)
'���:mobjBill=���ݶ���
'����:�����Ƿ�ɹ�
    Dim i As Long, j As Long, arrSQL As Variant, arrSMSQL As Variant
    Dim int��� As Integer, int�к� As Integer, strNO As String, strTmp As String
    Dim intParent As Integer, intParentNO As Integer
    Dim dbl���� As Double, dbl���� As Double, str��Ϣ As String, str���ܺ� As String
    Dim strDelInsure As String, arrDelInsure As Variant
    Dim strInsure As String, arrInsure As Variant
    Dim blnModiBill As Boolean
    Dim strSQL As String, strStuffDept As String '��¼���Ϸ��ϲ���
    
    Dim strAddDate As String '���ʷ���,�Զ���ҩ,���ϵ�ʱ��
    Dim blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    
    mobjBill.NO = zlDatabase.GetNextNo(14)
    strAddDate = "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    
    Call zlReSetDrawDrugDept
    
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    arrSMSQL = Array()
    
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.���� <> 0 Then
            intParent = 0: intParentNO = int���
            For Each mobjBillIncome In mobjBillDetail.InComes
                int��� = int��� + 1 '��ǰ��¼���
                
                '�������弰������ϸ
                With mobjBill
                    gstrSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & mobjBillDetail.����ID & "," & _
                        mobjBillDetail.��ҳID & "," & mobjBillDetail.סԺ�� & "," & "'" & mobjBillDetail.���� & "','" & _
                        mobjBillDetail.�Ա� & "','" & mobjBillDetail.���� & "','" & mobjBillDetail.���� & "','" & mobjBillDetail.�ѱ� & "'," & _
                        IIf(mobjBillDetail.����ID = 0, .��������ID, mobjBillDetail.����ID) & "," & mobjBillDetail.����ID & "," & .�Ӱ��־ & "," & _
                        mobjBillDetail.Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
                End With
                
                '�շ�ϸĿ����
                With mobjBillDetail
                    '�����������
                    If .��� <> int�к� Then
                        int�к� = .���
                        
                        '���´����������
                        If mobjBill.Details(.���).�������� = 0 Then
                            For i = .��� + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).�������� = .��� Then
                                    mobjBill.Details(i).�������� = int��� '������Ŀ�ж��������Ŀ(������)ʱ,ȡ��һ�����
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
    
                    gstrSQL = gstrSQL & IIf(.������Ŀ��, 1, 0) & "," & IIf(.���մ���ID = 0, "NULL", .���մ���ID) & ",'" & .���ձ��� & "',"
                    
                    dbl���� = .����
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                        dbl���� = Format(.���� * .Detail.סԺ��װ, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & "," & IIf(.ִ�в���ID = 0, "NULL", .ִ�в���ID) & ","
                    
                    '�ռ����Ϸ��ϲ���,�Ա��Զ�����
                    If gbytBilling = 0 And gint���ķ��Ͽ��� <> 0 Then
                        'gint���ķ��Ͽ���:0-���Զ����ϣ�1-�Զ����ϣ�2-�����ҿ���ʱ�Զ�����
                        If .ִ�в���ID <> 0 And .�շ���� = "4" And .Detail.�������� _
                            And ((gint���ķ��Ͽ��� = 2 And .ִ�в���ID = mobjBill.��������ID) Or gint���ķ��Ͽ��� = 1) Then
                            If InStr("," & strStuffDept, "," & .ִ�в���ID & ",") = 0 Then
                                strStuffDept = strStuffDept & "," & .ִ�в���ID
                            End If
                        End If
                    End If
                End With
                
                '������Ŀ����
                With mobjBillIncome
                    intParent = intParent + 1
                    dbl���� = .��׼����
                    If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 And gblnסԺ��λ Then
                        dbl���� = Format(.��׼���� / mobjBillDetail.Detail.סԺ��װ, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .������ĿID & "," & _
                        "'" & .�վݷ�Ŀ & "'," & dbl���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & "," & _
                        IIf(.ͳ���� = 0, "NULL", .ͳ����) & ","
                End With
                                                
                '��������
                gstrSQL = gstrSQL & _
                    "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & strAddDate & "," & _
                    "'" & mstrInNO & "'," & IIf(gbytBilling = 1, 1, 0) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                    "1," & IIf(mobjBillDetail.�շ���� = "4", mlng�������ID, mlngҩƷ���ID) & ",Null,'" & mobjBillDetail.ժҪ & "'," & _
                    "Null,Null,Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.���� & "',0," & mobjBill.��ҩ����ID & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
            Next
        End If
    Next
    
    '�޸�ǰ�˳�ԭ����
    If mstrInNO <> "" Then
        '���ж��Ƿ�ҽ�����˼ǵ���,�����Ϸ��Լ��(�����޸�ʱ������һ������ж�)
        If gbytBilling = 0 Then
            'ȥ����ҽ������ƥ����
            Call GetBillInsures(strDelInsure, mstrInNO)
            If strDelInsure <> "" Then arrDelInsure = Split(strDelInsure, ",")
        End If
        
        gstrSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        If gstrSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
        End If
    End If

    If UBound(arrSQL) >= 0 Then
        '��SQL���а��շ�ϸĿID����
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        '�жϼ��ʱ�֮���Ƿ���ҽ������(��ҩ���ڼ�¼����)
        strInsure = GetMultiInsures
        If strInsure <> "" Then arrInsure = Split(strInsure, ",")

        'ִ��SQL���
        On Error GoTo errH
        gcnOracle.BeginTrans
            blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            'ִ���Զ�����
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                For i = 0 To UBound(Split(strStuffDept, ","))
                    strSQL = "zl_�����շ���¼_��������(" & Split(strStuffDept, ",")(i) & ",26,'" & mobjBill.NO & _
                        "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1," & strAddDate & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Next
            End If
            
            '׼���Զ���ҩ(����ͨ����),�����������в��ܶ�������
            If mblnSendMateria Then
                Set rsTmp = Get����ҩ�嵥(mobjBill.NO, Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), True)
                If rsTmp.RecordCount > 0 Then
                    str���ܺ� = zlDatabase.GetNextNo(20)
                    ReDim arrSMSQL(rsTmp.RecordCount - 1)
                    For i = 0 To rsTmp.RecordCount - 1
                        arrSMSQL(i) = "ZL_ҩƷ�շ���¼_���ŷ�ҩ(" & rsTmp!�ⷿID & "," & rsTmp!ID & ",'" & UserInfo.���� & "'," & strAddDate & ",Null,Null,Null," & str���ܺ� & ")"
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            'ִ���Զ���ҩ
            For i = 0 To UBound(arrSMSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSMSQL(i)), Me.Caption)
            Next
            
            'ҽ���ӿ�
            '1.ҽ�����������ϴ�(ֻҪ��һ���ɹ����ύ)
            blnModiBill = False
            If mstrInNO <> "" And gbytBilling = 0 And strDelInsure <> "" Then
                For i = 0 To UBound(arrDelInsure)
                    If gclsInsure.GetCapability(support���������ϴ�, , arrDelInsure(i)) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , arrDelInsure(i)) Then
                        If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrDelInsure(i)) Then
                            If i = 0 Then gcnOracle.RollbackTrans: Exit Function
                        Else
                            blnModiBill = True '��������ʱ�ɹ��ϴ�����ϸ
                        End If
                    End If
                Next
            End If
            
            '2.����ʵʱ�ϴ�(ֻҪ��һ���ɹ����ύ)
            If gbytBilling = 0 And strInsure <> "" Then
                For i = 0 To UBound(arrInsure)
                    If gclsInsure.GetCapability(support�����ϴ�, , arrInsure(i)) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , arrInsure(i)) Then
                        str��Ϣ = ""
                        If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , arrInsure(i)) Then
                            '������޸�,ֻҪ����ʱ�гɹ��ϴ����ύ
                            If i = 0 And Not blnModiBill Then gcnOracle.RollbackTrans
                            If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
                            If i = 0 And Not blnModiBill Then Exit Function
                        End If
                    End If
                Next
            End If
        gcnOracle.CommitTrans
        blnTrans = False
        
        '1.ҽ�����������ϴ�
        If mstrInNO <> "" And gbytBilling = 0 And strDelInsure <> "" Then
            For i = 0 To UBound(arrDelInsure)
                If gclsInsure.GetCapability(support���������ϴ�, , arrDelInsure(i)) And gclsInsure.GetCapability(support������ɺ��ϴ�, , arrDelInsure(i)) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrDelInsure(i)) Then
                        MsgBox "�����е� " & GetInsureName(Val(arrDelInsure(i))) & " ���ʷ�����ҽ������ʧ��,��Щ���������ʣ�", vbInformation, gstrSysName
                    End If
                End If
            Next
        End If
        
        '2.����ʵʱ�ϴ�
        If gbytBilling = 0 And strInsure <> "" Then
            For i = 0 To UBound(arrInsure)
                If gclsInsure.GetCapability(support�����ϴ�, , arrInsure(i)) And gclsInsure.GetCapability(support������ɺ��ϴ�, , arrInsure(i)) Then
                    str��Ϣ = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , arrInsure(i)) Then
                        If str��Ϣ <> "" Then
                            MsgBox str��Ϣ, vbInformation, gstrSysName
                        Else
                            MsgBox "������ " & GetInsureName(Val(arrInsure(i))) & " �ķ�����ҽ������ʧ��,��Щ�����ѱ��棡", vbInformation, gstrSysName
                        End If
                    End If
                End If
            Next
        End If
        
        '���뵥����ʷ��¼(�������͵���)
        For i = 0 To cboNO.ListCount - 1
            strNO = strNO & "," & cboNO.List(i)
        Next
        strNO = mobjBill.NO & strNO
        cboNO.Clear
        For i = 0 To UBound(Split(strNO, ","))
            cboNO.AddItem Split(strNO, ",")(i)
            If i = 9 Then Exit For 'ֻ��ʾ10��
        Next
        
        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
    End If
    
    SaveBill = True
    Exit Function
errH:
    If Err.Description Like "*��ǰ���㵥�۲�һ��*" Then
       If blnTrans Then gcnOracle.RollbackTrans
       
       If MsgBox("ĳЩ����ҩƷ�۸��ѷ����仯��Ҫ�Զ�����۸���", vbYesNo + vbQuestion + vbDefaultButton1, App.ProductName) = vbYes Then
           Call CalcMoneys
           Call ShowDetails
           Exit Function
       End If
    Else
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Boolean
'���ܣ����ݵ��ݺŶ�ȡһ�ŵ��ݲ�����������
'������strNO=���ݺ�
'      blnDelete=True:���ʵ���ʱ����,False:���ĵ���ʱ����
    Dim rsTmp As ADODB.Recordset
    Dim curTotal As Currency, blnDo As Boolean, arrInsure As Variant
    Dim i As Long, lng����ID As Long, intSign As Integer
    Dim strSQL As String, strSQL1 As String, strSQL2 As String, strInsure As String, strFeeKind As String, strUserUnitIDs As String
        
    On Error GoTo errH
    
    mblnPrint = False
    
     '������֮ǰ�Ѽ��,������һ������Ȩ��
    If blnDelete Then
        '55380
        Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
        blnYP = zlStr.IsHavePrivs(mstrPrivsOpt, "ҩƷ����")
        blnZL = zlStr.IsHavePrivs(mstrPrivsOpt, "��������")
        blnWC = zlStr.IsHavePrivs(mstrPrivsOpt, "��������")
        If blnYP And blnWC And blnZL Then
            '����,������
        ElseIf blnYP And blnWC And Not blnZL Then
            strFeeKind = " And �շ����   In('4','5','6','7')"
        ElseIf blnYP And Not blnWC And blnZL Then
            strFeeKind = " And �շ����   <>'4'"
        ElseIf blnYP And Not blnWC And Not blnZL Then
            strFeeKind = " And �շ���� In('5','6','7')"
        ElseIf Not blnYP And blnWC And blnZL Then
            strFeeKind = " And �շ���� Not In('5','6','7')"
        ElseIf Not blnYP And Not blnWC And blnZL Then
            strFeeKind = " And �շ���� Not In('4','5','6','7')"
        ElseIf Not blnYP And blnWC And Not blnZL Then
            strFeeKind = " And �շ���� ='4'"
        End If
    End If
    
    Call ClearRows: Call Bill.ClearBill: mlngPreRow = 0 '���¶�ȡ����ʱ����ʼ���кű�־
    
    '��ȡ��������
    strNO = GetFullNO(strNO, 14)
    
    strSQL = _
    " Select A.��������ID,Nvl(A.�Ӱ��־,0) as �Ӱ��־," & _
    "       A.������,A.������,A.����Ա����,A.����ʱ��,A.���˲���ID " & _
    " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " ,��Ա�� C " & _
    " Where NO=[1] And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=1 And Nvl(A.����Ա����,A.������)=C.����" & _
    "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
    "       And Rownum=1 And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
            IIf(mbytInState = 0 And gbytBilling = 0, " And A.����Ա���� is Not Null", "") & _
            IIf(mbytInState = 0 And gbytBilling = 1, " And A.����Ա���� is Null And A.������ is Not NULL", "") & _
            IIf(mbytInState = 0 And gbytBilling = 2, " And A.����Ա���� is Null And A.������ is Not NULL", "")
    
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    End If
    
    If rsTmp.EOF Then
        MsgBox "û���ҵ��õ��ݣ�����õ����Ƿ�����סԺ���ʱ�.", vbInformation, gstrSysName
        Exit Function
    Else
        If blnDelete Then
            If InStr(mstrPrivsOpt, ";ȫԺ����;") = 0 Then
                strUserUnitIDs = GetUserUnits(True)
                If InStr("," & strUserUnitIDs & ",", "," & rsTmp!��������ID & ",") = 0 Then
                    MsgBox "��û��Ȩ�޶��������ҵĵ������ʣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            If mbytUseType = 0 Or mbytUseType = 1 Then
                If InStr(mstrPrivs, ";���в���;") = 0 And mlngUnitID > 0 Then
                    If InStr(1, "," & mstrUnitIDs & ",", "," & IIf(IsNull(rsTmp!���˲���ID), 0, rsTmp!���˲���ID) & ",") = 0 Then
                        MsgBox "��û��Ȩ�޶�ȡ���������ĵ��ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    cboNO.Text = strNO
    
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    chk�Ӱ�.Value = IIf(IsNull(rsTmp!�Ӱ��־), 0, rsTmp!�Ӱ��־)
                
    Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, Nvl(rsTmp!������), Nvl(rsTmp!��������ID, 0))
    
    '-----------------------------------------------------------------------------------
    '��ȡ�����շ�ϸĿ
    If blnDelete Then
         '�˷ѵ����迼�Ǻ󱸱�,ǰ��Ĳ����ѽ�ֹ
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        
        '��ȡ������ԭʼ��¼�ķ���ID
        strSQL1 = _
            " Select A.ID,A.���,A.�շ�ϸĿID," & _
            " Nvl(A.����,1)*A.����" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & " as ԭʼ����" & _
            " From סԺ���ü�¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
            " And A.�շ�ϸĿID=B.ҩƷID(+) And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=1" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
            IIf(mstr����IDs <> "", " And Instr([2],','||A.����ID||',')>0", "")
        
        '��ȡҩƷ�շ���¼�е�׼����
        strSQL2 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & ") as ׼������" & _
            " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            " And A.ҩƷID=B.ҩƷID(+) And A.���� IN(10,26) And A.����� is NULL" & _
            " Group by A.����ID"
        
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
        strSQL = "Select Nvl(�۸񸸺�,���) From סԺ���ü�¼ " & _
            " Where ��¼����=2 And �����־=2 And Nvl(�ಡ�˵�,0)=1" & _
            " And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1" & _
            IIf(mstrTime <> "", " And �Ǽ�ʱ��=[3]", "") & strFeeKind
            
        '����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
        Call GetBillInsures(strInsure, strNO)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , arrInsure(i)) Then
                    blnDo = True: Exit For 'ֻҪ��һ��������������
                End If
            Next
        Else
            blnDo = gbytBillOpt = 2
        End If
        If blnDo Then
            strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From סԺ���ü�¼ " & _
                " Where NO=[1] And ��¼���� IN(2,12)" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
                    
        '��Ϊ�ǽ�Ҫ��������ʣ�������ģ����Բ�����ֱ����ʱ�����ƣ����������
        strSQL = _
            " Select A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
            " A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
            " A.�շ�ϸĿID,C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From סԺ���ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=1" & _
            " And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            IIf(mstr����IDs <> "", " And Instr([2],','||A.����ID||',')>0", "") & _
            " Group by A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
            " A.�շ�ϸĿID,C.����,C.����,B.����,B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X.סԺ��λ"
            
        '��������
        '��"׼������=ԭʼ����"ʱ,�����ű���
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSQL = _
            " Select A.���,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
            " A.�շ�ϸĿID,A.����,A.���,A.����,A.���,A.��������,A.���㵥λ," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Avg(A.����),1) as ׼�˸���," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
            " Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
            " A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��,A.ִ�в���,A.���ӱ�־" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.���=B.��� And B.ID=C.����ID(+)" & _
            " Group by A.���,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�,A.�շ�ϸĿID,A.����,A.���," & _
            " A.����,A.���,A.��������,A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�в���,A.���ӱ�־" & _
            " Having Sum(A.����*A.����)<>0"
        If strInsure <> "" Then
            'ҽ�����˷��ÿ��ܲ�������,��������������(׼������=ԭʼ����)
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If Not gclsInsure.GetCapability(support�����ֳ�����ϸ, , arrInsure(i)) Then
                    strSQL = strSQL & " And (Nvl(C.׼������,Sum(A.����*A.����))=B.ԭʼ����" & _
                        " Or A.����ID+0 IN(Select ����ID From ������Ϣ Where ���� is NULL And ����ID=A.����ID))"
                    Exit For  'ֻҪ��һ��������,��������������
                End If
            Next
        End If
            
        strSQL = _
        " Select A.���,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��," & _
        "       A.���,A.��������,A.���㵥λ,A.׼�˸��� as ����,A.׼������ as ����,A.����," & _
        "       A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
        "       A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
        "       A.ִ�в���,A.���ӱ�־" & _
        " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        " Order by LPAD(A.����,10,' '),A.����ID,A.���"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '��ȡ���ʻ��۵�(�������),ֻ��ȡδ��˲���
        '���ÿ����ں󱸱���
        strSQL = _
            " Select Nvl(A.�۸񸸺�,A.���) as ���,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����," & _
            " A.����,A.�ѱ�,A.�շ�ϸĿID,C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From סԺ���ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=2 And Nvl(A.�ಡ�˵�,0)=1" & _
            " And A.��¼״̬=0 And �����־=2 And A.NO=[1]" & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.��¼״̬,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
            " A.�շ�ϸĿID,C.����,C.����,B.����,B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X.סԺ��λ"
            
        strSQL = "" & _
        " Select A.���,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
        "       A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������,A.���㵥λ," & _
        "       A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���,A.���ӱ�־" & _
        " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        " Order by LPAD(A.����,10,' '),A.����ID,A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        strSQL = _
            " Select Nvl(A.�۸񸸺�,A.���) as ���,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
            " A.�շ�ϸĿID,C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(" & intSign & "*A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��,Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & ",�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=2 And Nvl(A.�ಡ�˵�,0)=1 And �����־=2" & _
            " And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
            " A.�շ�ϸĿID,C.����,C.����,B.����,B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X.סԺ��λ"
            
        strSQL = "" & _
        " Select A.���,A.����ID,A.��ҳID,A.Ӥ����,A.����,A.�Ա�,A.����,A.����,A.�ѱ�," & _
        "       A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������,A.���㵥λ," & _
        "       A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���,A.���ӱ�־" & _
        " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        " Order by LPAD(A.����,10,' '),A.����ID,A.���"
        
    End If
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, "," & mstr����IDs & ",", CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, "," & mstr����IDs & ",")
    End If
    
    If rsTmp.EOF Then Exit Function
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    ReDim marrSerial(1 To rsTmp.RecordCount)
    Set mcolPatiInfo = New Collection
    
    For i = 1 To rsTmp.RecordCount
        If gbytBilling = 2 And Not mblnPrint Then mblnPrint = True

        marrSerial(i) = rsTmp!��� '���ڼ������ʻ򻮼����
        mcolPatiInfo.Add rsTmp!����ID & "," & Val("" & rsTmp!��ҳID) & "," & Val("" & rsTmp!Ӥ����), "R" & i
        
        Bill.TextMatrix(i, BillCol.����) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        Bill.TextMatrix(i, BillCol.�Ա�) = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
        Bill.TextMatrix(i, BillCol.����) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        Bill.TextMatrix(i, BillCol.����) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        Bill.TextMatrix(i, BillCol.�ѱ�) = IIf(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
        Bill.TextMatrix(i, BillCol.���) = rsTmp!���
        Bill.TextMatrix(i, BillCol.��Ŀ) = rsTmp!����
        '���˺�:27990 2010-02-22 17:01:30
        Bill.TextMatrix(i, BillCol.��Ʒ��) = Nvl(rsTmp!��Ʒ��)
        Bill.TextMatrix(i, BillCol.���) = IIf(IsNull(rsTmp!���), "", rsTmp!���)
        Bill.TextMatrix(i, BillCol.��λ) = IIf(IsNull(rsTmp!���㵥λ), "", rsTmp!���㵥λ)
        Bill.TextMatrix(i, BillCol.����) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        Bill.TextMatrix(i, BillCol.����) = FormatEx(rsTmp!����, 5)
        Bill.TextMatrix(i, BillCol.����) = Format(rsTmp!����, gstrFeePrecisionFmt)
        Bill.TextMatrix(i, BillCol.Ӧ�ս��) = Format(rsTmp!Ӧ�ս��, gstrDec)
        Bill.TextMatrix(i, BillCol.ʵ�ս��) = Format(rsTmp!ʵ�ս��, gstrDec)
        Bill.TextMatrix(i, BillCol.ִ�п���) = Nvl(rsTmp!ִ�в���)
        Bill.TextMatrix(i, BillCol.��־) = IIf(rsTmp!���ӱ�־ = 1, "��", "")
        Bill.TextMatrix(i, BillCol.����) = IIf(IsNull(rsTmp!��������), "", rsTmp!��������)
        
        '�������ʱ�־
        If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
            If mlngDelRow = 0 Or mlngDelRow <> 0 And mlngDelRow = rsTmp!��� Then
                Bill.TextMatrix(i, Bill.Cols - 1) = "��"
            End If
        End If
        
        curTotal = curTotal + rsTmp!ʵ�ս��
        rsTmp.MoveNext
    Next
    '����б༭����������ɫ
    Call InitBillColumnColor
    
    Bill.Redraw = True
    
    ReadBill = True
    txt����.Text = Format(curTotal, gstrDec)
    Call Bill_EnterCell(Bill.Row, Bill.Col)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetShowCol()
'���ܣ������еĿ���(���ʱչ��)
    mrsClass.Filter = "����='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.����) = 0
    ElseIf Bill.ColWidth(BillCol.����) = 0 Then
        Bill.ColWidth(BillCol.����) = 300
    End If
End Sub
Private Sub InitBillColumnColor()
    
    Bill.SetColColor BillCol.���, &HE7CFBA
    Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE7CFBA
    Bill.SetColColor BillCol.ִ�п���, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.��־, &HE0E0E0
End Sub
Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double
    Dim int���� As Integer
    Dim lngִ�п��� As Long, strִ�п��� As String
    'ҩƷ�����
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                    lngִ�п��� = .ִ�в���ID: strִ�п��� = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                
                    If InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then
                            dblStock = dblStock / .Detail.סԺ��װ
                        End If
                        .Detail.��� = dblStock  '��¼��ǰ��ҩƷ���
                        Call ShowStock(.Detail.����, .Detail.���)
                        
                        'ҩ���ı�,ʵ��ҩƷ���¼���۸�
                        'If .Detail.��� Then    '����ѱ�ļ��㷽ʽ�ǳɱ��ۼ��շ�,����Ҫ����۸�,����򻯲����ж�
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call CalcOneTotal(Bill.Row)
                        'End If
                    ElseIf .�շ���� = "4" And .Detail.�������� Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = dblStock
                        Call ShowStock(.Detail.����, .Detail.���)
                        
                        '���ϲ��Ÿı�,ʱ���������¼���۸�
                        If .Detail.��� Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call CalcOneTotal(Bill.Row)
                        End If
                    ElseIf InStr(",4,5,6,7,", .�շ����) = 0 Then
                        If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                    End If
                    int���� = Val(mobjBill.Details(Bill.Row).��ҩ����)
                    If int���� <> 0 And mobjBill.Details(Bill.Row).���� <> 0 Then
                        If gclsInsure.GetCapability(supportʵʱ���, mobjBill.Details(Bill.Row).����ID, int����) Then
                            If gclsInsure.CheckItem(int����, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = strִ�п���: .ִ�в���ID = lngִ�п���
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.cboObj.Text = strִ�п���: .ִ�в���ID = lngִ�п���
                            Exit Sub
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Public Function GetBillIndex(strFind As String) As Long
'���ܣ������δ�����ComboBox������ֵ
'������cbo=ComboBox,strFind=�����ַ���
    Dim i As Long
    If strFind = "" Then GetBillIndex = -1: Exit Function
    For i = 0 To Bill.ListCount - 1
        If InStr(Bill.List(i), strFind) > 0 Then
            GetBillIndex = i
            Exit Function
        End If
    Next
    GetBillIndex = -1
End Function

Private Function CalcOneTotal(lngRow As Long, Optional blnShow As Boolean = True) As Currency
'���ܣ�������ʵ��е�ǰ�в��˵��ڵ�ǰ�����еķ��úϼ�
'˵����������Ϊ׼
    Dim i As Long, strName As String, curTotal As Currency
    Dim tmpBillInCome As New BillInCome
    
    If Bill.TextMatrix(lngRow, BillCol.����) = "" Then
        If blnShow Then txt����.Text = gstrDec
    Else
        If mobjBill.Details.Count = 0 Then
            strName = Bill.TextMatrix(lngRow, BillCol.����) '��������
            If blnShow Then lbl����.Caption = strName
            For i = 1 To Bill.Rows - 1
                If IsNumeric(Bill.TextMatrix(i, BillCol.ʵ�ս��)) Then
                    If Bill.TextMatrix(i, BillCol.����) = strName Then
                        curTotal = curTotal + CCur(Bill.TextMatrix(i, BillCol.ʵ�ս��))
                    End If
                End If
            Next
        Else
            If mobjBill.Details.Count >= lngRow Then
                strName = mobjBill.Details(lngRow).����ID & mobjBill.Details(lngRow).����   '��������
                If blnShow Then lbl����.Caption = mobjBill.Details(lngRow).����
            ElseIf mrsInfo.State = 1 Then
                strName = mrsInfo!����ID & mrsInfo!����
                If blnShow Then lbl����.Caption = mrsInfo!����
            ElseIf Bill.TextMatrix(lngRow, BillCol.����) <> "" And mobjBill.Details.Count < lngRow And mobjBill.Details.Count >= lngRow - 1 And lngRow > 1 Then
                strName = mobjBill.Details(lngRow - 1).����ID & mobjBill.Details(lngRow - 1).����
                If blnShow Then lbl����.Caption = mobjBill.Details(lngRow - 1).����
            End If
            For i = 1 To Bill.Rows - 1
                If mobjBill.Details.Count >= i Then
                    If mobjBill.Details(i).����ID & mobjBill.Details(i).���� = strName Then
                        For Each tmpBillInCome In mobjBill.Details(i).InComes
                            curTotal = curTotal + tmpBillInCome.ʵ�ս��
                        Next
                    End If
                End If
            Next
        End If
        If blnShow Then txt����.Text = Format(curTotal, gstrDec)
    End If
    CalcOneTotal = curTotal
End Function

Private Function GetDetailNum(lngRow As Long) As Double
'���ܣ���ȡ����ָ��ϸĿ���ܼ�������(����������)
'������lngRow=��ǰ������
    Dim rsTmp As ADODB.Recordset
    Dim lngNum As Long, i As Long
    Dim strSQL As String
    Dim lng����ID As Long, lng��ҳID As Long
        
    If lngRow <= mobjBill.Details.Count Then
        lng����ID = mobjBill.Details(lngRow).����ID
        lng��ҳID = mobjBill.Details(lngRow).��ҳID
        
        '��ǰ�����е�����
        For i = 1 To mobjBill.Details.Count
            If i <> lngRow And mobjBill.Details(i).�շ�ϸĿID = mobjBill.Details(lngRow).�շ�ϸĿID And mobjBill.Details(i).����ID = lng����ID Then
                lngNum = lngNum + mobjBill.Details(i).���� * IIf(mobjBill.Details(i).���� = 0, 1, mobjBill.Details(i).����)
            End If
        Next
        '���ݿ��е�����
        strSQL = _
            "Select Sum(A.����*Nvl(A.����,1)" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & ") as NUM" & _
            " From סԺ���ü�¼ A,ҩƷ��� B" & _
            " Where A.�۸񸸺� is Null And A.���ʷ���=1" & _
            IIf(gbytBilling = 0, " And A.��¼״̬<>0", "") & _
            " And A.����ID=[1] And Nvl(A.��ҳID,0)=[2]" & _
            " And A.�շ�ϸĿID=B.ҩƷID(+) And A.�շ�ϸĿID+0=[3]"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, mobjBill.Details(lngRow).�շ�ϸĿID)
        If Not rsTmp.EOF Then
            lngNum = lngNum + Nvl(rsTmp!Num, 0)
        End If
        GetDetailNum = lngNum
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWorkUnit(ByVal lngҩƷID As Long, ByVal str��� As String) As Boolean
'���ܣ�ȡ���пɹ�ѡ���ҩ��
    Dim strSQL As String, strҩ�� As String, bytDay As Byte
    Dim int������� As Integer, str������� As String
    Dim int������Դ As Integer, lng��������ID As Long
    
    '������Ŀ��Ȩ��ȷ��ҩ���ķ������
    int������� = Get�������(lngҩƷID)
    'int������� = mobjDetail.�������  '�޸�,����ʱû�и�ֵ
    If int������� = 1 Then
        str������� = "1,3"
    ElseIf int������� = 2 Then
        str������� = "2,3"
    ElseIf int������� = 3 Then
        If InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
            str������� = "1,2,3"
        Else
            str������� = "2,3"
        End If
    Else
        str������� = "2,3"
    End If
        
    'ȷ��������Դ
    int������Դ = Get������Դ(Bill.Row)
    
    'ȷ�����˿���
    If mrsInfo.State = 1 Then
        lng��������ID = Nvl(mrsInfo!����ID, 0)
    ElseIf Bill.TextMatrix(Bill.Row, BillCol.����) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then
        lng��������ID = mobjBill.Details(Bill.Row - 1).����ID
    Else
        lng��������ID = mobjBill.Details(Bill.Row).����ID
    End If
    
    '��������
    If lng��������ID = 0 And cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    
    If str��� = "4" Then
        strSQL = _
        "Select Distinct c.Id, c.����, c.����, c.����, b.��������, b.�������" & vbNewLine & _
        "From �շ�ִ�п��� A, ��������˵�� B, ���ű� C" & vbNewLine & _
        "Where a.ִ�п���id + 0 = b.����id And b.�������� = '���ϲ���' And b.������� IN(" & str������� & ") And b.����id = c.Id And" & vbNewLine & _
        "      (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null) And" & vbNewLine & _
        "      (a.������Դ Is Null Or a.������Դ = [1]) And" & vbNewLine & _
        "      (a.��������id Is Null Or a.��������id = [2] Or Exists (Select 1 From �������Ҷ�Ӧ Where ����id = [2] And a.��������id = ����id)) And a.�շ�ϸĿid = [3]" & vbNewLine & _
        "Order By b.�������, c.����"
    Else
        '��ҩƷ����ȷ��ҩ������
        Select Case str���
            Case "5"
                strҩ�� = "��ҩ��"
            Case "6"
                strҩ�� = "��ҩ��"
            Case "7"
                strҩ�� = "��ҩ��"
        End Select
    
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not gblnҩ���ϰల�� Then
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
            "       And B.������� IN(" & str������� & ") And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
            "       And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
            "       And B.������� IN(" & str������� & ") And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And D.����ID=C.ID And D.����=[5]" & _
            "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
            "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
            "       And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        End If
    End If
    
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, int������Դ, lng��������ID, lngҩƷID, strҩ��, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FillPatient(lng����ID As Long)
    Dim i As Long, j As Long, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem
    Dim str��Ժ As String
    On Error GoTo errH
    
    str��Ժ = "    Exists(Select 1 From ��Ժ���� ZY Where ZY.����ID=B.����ID)"
    '�Ƿ����ǿ�Ƽ���Ȩ��
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If gintOutDay = 0 Then
            strSQL = "   And " & str��Ժ
        Else
            strSQL = " And (" & str��Ժ & " Or B.��Ժ����>Trunc(Sysdate)-" & gintOutDay & ")"
        End If
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        If gintOutDay = 0 Then
            strSQL = " And (" & str��Ժ & " And B.״̬<>3 Or X.�������<>0 And " & str��Ժ & "  And B.״̬=3)"
        Else
            strSQL = " And (" & str��Ժ & " And B.״̬<>3 Or X.�������<>0 And (" & str��Ժ & " And B.״̬=3 Or B.��Ժ����>Trunc(Sysdate)-" & gintOutDay & "))"
        End If
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If gintOutDay = 0 Then
            strSQL = " And (" & str��Ժ & "  And B.״̬<>3 Or X.�������=0 And " & str��Ժ & "  And B.״̬=3)"
        Else
            strSQL = " And (" & str��Ժ & "  And B.״̬<>3 Or X.�������=0 And (" & str��Ժ & "  And B.״̬=3 Or B.��Ժ����>Trunc(Sysdate)-" & gintOutDay & "))"
        End If
    Else
        'û��Ȩ�޶Գ�Ժ��Ԥ��Ժ���˽���
        strSQL = " And " & str��Ժ & "  And B.״̬<>3"
    End If
    
    '���۲��˼���Ȩ��
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        strSQL = strSQL & " And Nvl(B.��������,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        strSQL = strSQL & " And Nvl(B.��������,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strSQL = strSQL & " And Nvl(B.��������,0) IN(0,2)"
    Else
        strSQL = strSQL & " And Nvl(B.��������,0)=0"
    End If
    
    lvwPati.ListItems.Clear
    
    strSQL = "Select A.����ID,A.סԺ��,nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,A.����," & _
            " B.��Ժ���� as ��λ,B.��Ժ����,B.����,B.��������,B.��������" & _
            " From ������Ϣ A,������ҳ B,������� X" & _
            " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID" & _
            " And Nvl(B.��ҳID,0)<>0 And A.��ҳID=B.��ҳID" & strSQL & _
            " And A.����ID=X.����ID(+)  And X.����(+)=1 And X.����(+)=2 And B.��Ժ����ID = [1]" & _
            " Order by A.סԺ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If IIf(IsNull(rsTmp!��������), 0, rsTmp!��������) = 0 Then
                Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID, rsTmp!����ID, , 1)
            Else
                Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID, rsTmp!����ID, , 2)
            End If
            objItem.SubItems(1) = IIf(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
            objItem.SubItems(2) = rsTmp!����
            objItem.SubItems(3) = IIf(IsNull(rsTmp!��λ), "", rsTmp!��λ)
            objItem.SubItems(4) = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
            objItem.SubItems(5) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
            objItem.SubItems(6) = IIf(IsNull(rsTmp!��Ժ����), "��", "")
                        
            objItem.ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
            Next
            
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMoneyStr(lng����ID As Long) As String
'���ܣ���������ID��ȡ���˷�����Ϣ
    Dim i As Long
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).����ID = lng����ID Then
            GetMoneyStr = mobjBill.Details(i).���￨��
            Exit For
        End If
    Next
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'���ܣ���ʾ\�������ʱ�־��
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "����" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "����"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 450
            Bill.ColData(Bill.Cols - 1) = BillColType.CheckBox
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���) - 100
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ) - 200
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���) - 150
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���)
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ)
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���)
            Bill.Redraw = True
        End If
    End If
    
    cmdSelALL.Visible = blnShow
    cmdClear.Visible = blnShow
    
    If blnShow Then
        chkIn.Visible = False
        txtIn.Visible = False
        txt����.Left = Val(txt����.Tag) - chkIn.Width - txtIn.Width
        lbl����.Left = txt����.Left - lbl����.Width - 45
        txt����.Left = Val(txt����.Tag) - chkIn.Width - txtIn.Width
        lbl����.Left = txt����.Left - lbl����.Width - 45
    Else
        If mbytInState = 0 And (gbytBilling = 0 Or gbytBilling = 1) Then
            chkIn.Visible = True
            txtIn.Visible = True
            txt����.Left = Val(txt����.Tag)
            lbl����.Left = txt����.Left - lbl����.Width - 45
            txt����.Left = Val(txt����.Tag)
            lbl����.Left = txt����.Left - lbl����.Width - 45
        Else
            chkIn.Visible = False
            txtIn.Visible = False
            txt����.Left = Val(txt����.Tag) - chkIn.Width - txtIn.Width
            lbl����.Left = txt����.Left - lbl����.Width - 45
            txt����.Left = Val(txt����.Tag) - chkIn.Width - txtIn.Width
            lbl����.Left = txt����.Left - lbl����.Width - 45
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'���ܣ���ȡָ���е�ԭʼ�п�
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Function SaveModi() As Boolean
'���ܣ����浱ǰ�޸ĵķ��õ���
    Dim strSQL As String
    
    strSQL = "zl_���˷��ü�¼_Update('" & cboNO.Text & "',2,'" & zlStr.NeedName(cbo������.Text) & "'," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),NULL,2)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To Bill.Cols - 1
        If Bill.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Function GetInputDetail(ByVal lng��Ŀid As Long, ByVal int���� As Integer) As Detail
'���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
       
    If int���� > 0 Then
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,M.Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C.סԺ��װ) as סԺ��װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C.סԺ��λ) as סԺ��λ,D.��������,A.¼������,C.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,����֧����Ŀ M" & _
        " Where A.���=B.���� And A.ID=C.ҩƷID(+) And A.ID=D.����ID(+)" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.ID=M.�շ�ϸĿID(+) And M.����(+)=[2]" & vbNewLine & _
        "       And A.ID=[1]"
    Else
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,0 as Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C.סԺ��װ) as סԺ��װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C.סԺ��λ) as סԺ��λ,D.��������,A.¼������,C.��ҩ��̬" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1" & _
        " Where A.���=B.���� And A.ID=C.ҩƷID(+) And A.ID=D.����ID(+)" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, int����)
    With objDetail
        .ID = rsTmp!ID
        .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0) '�����ж������ظ�
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���� = rsTmp!����
        .��� = Nvl(rsTmp!���)
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .סԺ��λ = Nvl(rsTmp!סԺ��λ)
        .סԺ��װ = Nvl(rsTmp!סԺ��װ, 1)
        .���� = Nvl(rsTmp!����, 0) = 1 '�Ƿ�ҩ������
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1 '��ҩƷ�����Ƿ�ʱ��
        .���� = Nvl(rsTmp!��������)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .������� = Nvl(rsTmp!�������, 0)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
        .�������� = Nvl(rsTmp!��������, 0) = 1
        .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
        .¼������ = Val("" & rsTmp!¼������)
        .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
        .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True, Optional strҽ�Ƹ��� As String) As Integer
'���ܣ����ָ��ҩƷ�е�ְ���Ƿ��뵱ǰҽ����ְ����ƥ��
'������tmpDetail=�������Ŀ,����Ϊ������;
'      blnCommon=�Ƿ������ļ���ֻ���ҽ�������Ѳ��˵ļ��
'      strҽ�Ƹ���=����һ�й��ѻ�ҽ�����˼��ʱ,Ҫ����
'���أ���ƥ�����,0Ϊ��ȷ
'˵����ְ��1=����,2=����,3=�м�,4=����/ʦ��,5=Ա/ʿ,9=��Ƹ
    Dim i As Long, intְ��A As Integer, intְ��B As Integer
    Dim strTmp As String, bytҽ�Ƹ����� As Byte, strAllDuty As String
    
    
    If cbo������.ListIndex = -1 Then Exit Function
    strAllDuty = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    Call GetOperatorInfo(mrs������, mobjBill.������, , intְ��A)
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                If Not blnCommon Then
                    If mobjBill.Details(i).ҽ�Ƹ��� <> "" Then
                        'ҽ���򹫷Ѳ���
                        '����:45605
                         If zlIsCheckMedicinePayMode(mobjBill.Details(i).ҽ�Ƹ���) Then

                            intְ��B = Val(Right(mobjBill.Details(i).Detail.����ְ��, 1))
                            If intְ��B > 0 Then
                                If intְ��A = 0 Then
                                    strTmp = "�� " & i & " �в���:" & mobjBill.Details(i).���� & ",ҽ�Ƹ��ʽΪ:" & mobjBill.Details(i).ҽ�Ƹ��� & "," & _
                                        vbCrLf & "ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                                    CheckDuty = 1
                                ElseIf intְ��B < intְ��A Then
                                    strTmp = "�� " & i & " �в���:" & mobjBill.Details(i).���� & ",ҽ�Ƹ��ʽΪ:" & mobjBill.Details(i).ҽ�Ƹ��� & "," & _
                                        vbCrLf & "ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                                    CheckDuty = i: Exit For
                                End If
                            End If
                        End If
                    End If
                Else
                    intְ��B = Val(Left(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.���) = 0 Then Exit Function
        If Not blnCommon Then
            If strҽ�Ƹ��� = "" Then Exit Function
            'ҽ���򹫷Ѳ���
            '����:45605
             If zlIsCheckMedicinePayMode(strҽ�Ƹ���) = False Then Exit Function
            intְ��B = Val(Right(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "��ǰ����ҽ�Ƹ��ʽΪ:" & strҽ�Ƹ��� & "," & _
                        vbCrLf & "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "��ǰ����ҽ�Ƹ��ʽΪ:" & strҽ�Ƹ��� & "," & _
                        vbCrLf & "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        Else
            intְ��B = Val(Left(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function Check��������(Optional intRow As Integer) As Boolean
'���ܣ����ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    Dim strSQL As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim rsҽ�� As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim blnҽ�� As Boolean, bln���� As Boolean
    
    Check�������� = True
        
    On Error GoTo errH
    '�޷����
    If intRow > 0 Then
        If mobjBill.Details(intRow).ҽ�Ƹ��� = "" Then Exit Function
        'ҽ���򹫷Ѳ���
        '����:45605
        If zlIsCheckMedicinePayMode(mobjBill.Details(intRow).ҽ�Ƹ���, blnҽ��, bln����) = False Then Exit Function
        bytType = IIf(blnҽ��, 1, 2)
        
        '��ȡ�������
        If bytType = 1 Then
            strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstrҽ���������� & ") Order by ����"
        Else
            strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstr���ѷ������� & ") Order by ����"
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If rsTmp.EOF Then Exit Function
    
        If mobjBill.Details(intRow).Detail.���� = "" Then
            MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������δ���ã�", vbInformation, gstrSysName
            Check�������� = False
        Else
            rsTmp.Filter = "����='" & mobjBill.Details(intRow).Detail.���� & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������Ϊ""" & _
                    mobjBill.Details(intRow).Detail.���� & """,����" & _
                    IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�", vbInformation, gstrSysName
                Check�������� = False
            End If
        End If
    Else
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstrҽ���������� & ") Order by ����"
        Call zlDatabase.OpenRecordset(rsҽ��, strSQL, Me.Caption)
        
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstr���ѷ������� & ") Order by ����"
        Call zlDatabase.OpenRecordset(rs����, strSQL, Me.Caption)
        
        For i = 1 To mobjBill.Details.Count
        
            Call zlIsCheckMedicinePayMode(mobjBill.Details(i).ҽ�Ƹ���, blnҽ��, bln����)
            bytType = IIf(blnҽ��, 1, IIf(bln����, 2, 0))
            
            If InStr(",1,2,", bytType) > 0 Then
                Set rsTmp = Nothing
                If bytType = 1 Then
                    rsҽ��.Filter = 0
                    Set rsTmp = rsҽ��
                Else
                    rs����.Filter = 0
                    Set rsTmp = rs����
                End If
                If Not rsTmp.EOF Then
                    If mobjBill.Details(i).Detail.���� = "" Then
                        If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������δ���ã�" & vbCrLf & "ȷʵҪ���浥����", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Check�������� = False: Exit For
                        End If
                    Else
                        rsTmp.Filter = "����='" & mobjBill.Details(i).Detail.���� & "'"
                        If rsTmp.EOF Then
                            If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������Ϊ""" & _
                                mobjBill.Details(i).Detail.���� & """,����" & _
                                IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�" & vbCrLf & "ȷʵҪ���浥����", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Check�������� = False: Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReCalcInsure()
'���ܣ��޸ĵ���ʱ,���¼���ͳ������������Ϣ
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String
    Dim intInsure As Integer
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).����ID <> 0 And Val(mobjBill.Details(i).��ҩ����) > 0 Then
            For j = 1 To mobjBill.Details(i).InComes.Count
                intInsure = Val(mobjBill.Details(i).��ҩ����)
                If intInsure <> 0 Then
                    dblAllTime = mobjBill.Details(i).���� * mobjBill.Details(i).����
                    If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                        If gblnסԺ��λ Then dblAllTime = dblAllTime * mobjBill.Details(i).Detail.סԺ��װ
                    End If
                
                    strInfo = gclsInsure.GetItemInsure(mobjBill.Details(i).����ID, mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).InComes(j).ʵ�ս��, False, intInsure, _
                        mobjBill.Details(i).ժҪ & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Details(i).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(i).���մ���ID = Val(Split(strInfo, ";")(1))
                        mobjBill.Details(i).InComes(j).ͳ���� = Val(Split(strInfo, ";")(2))
                        mobjBill.Details(i).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Function PhysicExist(objDetail As Detail, intRow As Integer, lng����ID As Long) As Boolean
'���ܣ��ж�ָ��ҩƷ�ڵ������Ƿ��Ѿ�����
'������objDetail=��Ŀ,intRow=Ҫ�жϵ���
'˵����ʱ�ۻ����ҩƷ��ͬһҩ����ֹ�ظ�����(�������ʾ,����ʱ��ֹ)
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.���� Or mobjBill.Details(i).Detail.���) _
                    And (objDetail.���� Or objDetail.���) Then
                    If objDetail.��� = "4" Then
                        If MsgBox("��������""" & objDetail.���� & """�ڵ������Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺����������Ϊ������ʱ�۲���,�ظ�����ʱ���뱣֤���ǵ�ִ��ҩ����ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ������Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺��ҩƷΪ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵ�ִ��ҩ����ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                ElseIf mobjBill.Details(i).����ID = lng����ID Then
                    If objDetail.��� = "4" Then
                        If MsgBox("�ò����Ѿ�������������""" & objDetail.���� & """,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("�ò����Ѿ�����ҩƷ""" & objDetail.���� & """,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function


Private Function Checkִ�п���() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).ִ�в���ID = 0 Or Bill.TextMatrix(i, BillCol.ִ�п���) = "" Then
            If Not (InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 And gbln���뷢ҩ) Then
                Checkִ�п��� = i: Exit Function
            End If
        End If
    Next
End Function

Private Function Check�������() As Integer
'���ܣ������ʱ��в��˵ļ��ʷ�����Ŀ�ķ�������Ƿ�һ��
'˵������Ϊ�������������۲���,�����д˼��
'���أ���һ�µķ�����,Ϊ0ʱ����
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    With mobjBill
        For i = 1 To .Details.Count
            If InStr(",0,2,", .Details(i).��������) > 0 Then
                'סԺ���˻�סԺ���۲���,������ֻ�������������Ŀ
                If .Details(i).Detail.������� = 1 Then
                    MsgBox "�� " & i & " ����Ŀ""" & .Details(i).Detail.���� & """������������,����""" & .Details(i).���� & """����ʹ��.", vbInformation, gstrSysName
                    Check������� = i: Exit Function
                End If
            ElseIf InStr(",1,-1,", .Details(i).��������) > 0 Then
                '������Ժ����(ҽ������)���������۲���,������ֻ������סԺ����Ŀ
                If .Details(i).Detail.������� = 2 Then
                    MsgBox "�� " & i & " ����Ŀ""" & .Details(i).Detail.���� & """��������סԺ,����""" & .Details(i).���� & """����ʹ��.", vbInformation, gstrSysName
                    Check������� = i: Exit Function
                End If
            End If
            If .Details(i).Detail.������� = 0 Then
                MsgBox "�� " & i & " ����Ŀ""" & .Details(i).Detail.���� & """�������ڲ���,����""" & .Details(i).���� & """����ʹ��.", vbInformation, gstrSysName
                Check������� = i: Exit Function
            End If
        Next
    End With
End Function

Private Sub SetIntureColor()
'���ܣ����뵥�ݺ��ҽ����������Ϊ��ɫ
    Dim intRow As Integer, intCol As Integer, i As Integer
    
    intRow = Bill.Row: intCol = Bill.Col
    Bill.Col = BillCol.����
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).��ҩ���� <> "" Then
            Bill.Row = i
            Bill.MsfObj.CellForeColor = vbRed
        End If
    Next
    Bill.Row = intRow: Bill.Col = intCol
End Sub
Private Function Get��������ID() As Long
    If cbo��������.ListIndex <> -1 Then
        Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        Get��������ID = UserInfo.����ID
    End If
End Function
Private Function Get������Դ(ByVal lngRow As Long) As Integer
'���ܣ���ȡ��ǰ���˵���Դ(��Ϊ���Զ��������۲��˼���)
    Dim int�������� As Integer
    
    int�������� = -2
    If mobjBill.Details.Count >= lngRow Then
        int�������� = mobjBill.Details(lngRow).��������
    ElseIf mrsInfo.State = 1 Then
        int�������� = mrsInfo!��������
    ElseIf Bill.TextMatrix(lngRow, BillCol.����) <> "" And lngRow > 1 Then
        int�������� = mobjBill.Details(lngRow - 1).��������
    End If
    If int�������� <> -2 Then
        If int�������� = 0 Or int�������� = 2 Then
            Get������Դ = 2
        ElseIf int�������� = 1 Or int�������� = -1 Then
            Get������Դ = 1 '���ﲡ��(ҽ������)���������۲���
        End If
    Else
        Get������Դ = 2 'ȱʡΪ2
    End If
End Function

Private Sub zlReSetDrawDrugDept()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ӧ�Ĺ���,���»�ȡ��ҩ����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-29 18:23:12
    '����:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    '4)  סԺ���ʡ����ҷ�ɢ���ʣ������ɲ���ʹ�ã�Ҳ������ҽ������ʹ�á�
    '    a)  �жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
    '    b)  �������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    If mbytUseType = 2 Then
        'ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
        mobjBill.��ҩ����ID = mlngDeptID: Exit Sub
    End If
    If mrs��ҩ����.RecordCount = 0 Then
        '�жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
        mobjBill.��ҩ����ID = mobjBill.����ID: Exit Sub
    End If
    '�������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    If mrs��ҩ����.RecordCount = 1 Then
        'ֻ��һ������,�϶�����
        If mrs��ҩ����.EOF Then mrs��ҩ����.MoveFirst
         mobjBill.��ҩ����ID = Val(Nvl(mrs��ҩ����!ID)): Exit Sub
    End If
    'ѡ��Ŀ������ĸ������ĸ�
    With cboDrawDept
        If .ListIndex < 0 Then Exit Sub
        If mobjBill.��ҩ����ID <> .ItemData(.ListIndex) Then mobjBill.��ҩ����ID = .ItemData(.ListIndex): Exit Sub
    End With
End Sub
Private Sub zlLoadDrawDeptData(ByVal bytUseType As Byte, Optional ByVal lngDeptID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:bytUseType:���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
    '����:
    '����:
    '����:���˺�
    '����:2009-07-29 15:05:18
    '����:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    '4)  סԺ���ʡ����ҷ�ɢ���ʣ������ɲ���ʹ�ã�Ҳ������ҽ������ʹ�á�
    '    a)  �жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
    '    b)  �������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    
    On Error GoTo errHandle
    
    'ҽ������
    If bytUseType = 2 Then
        '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
        strSQL = "Select ID,����,���� From ���ű� where id=[2]"
    Else
        strSQL = _
            " Select distinct  A.ID, A.����,A.����   " & vbNewLine & _
            " From ���ű� A, ��������˵�� B,������Ա C" & vbNewLine & _
            " Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)  " & _
            "       And A.ID = B.����id and a.id=C.����ID and C.��Աid=[1] " & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "       AND B.�������� IN('���','����','����','����','Ӫ��') " & _
            " Order by ����"
    End If
    Set mrs��ҩ���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, lngDeptID)
    With mrs��ҩ����
        cboDrawDept.Clear
        Do While Not .EOF
            cboDrawDept.AddItem IIf(zlIsShowDeptCode, Nvl(!����) & "-", "") & Nvl(!����)
            cboDrawDept.ItemData(cboDrawDept.NewIndex) = Val(Nvl(!ID))
            If Val(Nvl(!ID)) = UserInfo.����ID Then cboDrawDept.ListIndex = cboDrawDept.NewIndex
            .MoveNext
        Loop
        If .RecordCount <> 0 And cboDrawDept.ListIndex < 0 Then cboDrawDept.ListIndex = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetDrawDrugDeptVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ҩ���ŵ�visibled����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-29 19:07:38
    '����:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    ' mbytUseType As Byte '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
    
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    If mbytUseType = 2 Then
        fraDrawDept.Visible = False
    ElseIf chkCancel.Value = 1 Then
        '����Ҳ���ܿ���
        fraDrawDept.Visible = False
    Else
        'mbytInState As Byte '0-ִ��,1-����,2-����,3-����
        ' gbytBilling:0-����,1-����,2-���
        fraDrawDept.Visible = mrs��ҩ����.RecordCount > 1 And (mbytInState = 0 And gbytBilling <> 2)
    End If
    Call Form_Resize
End Sub
Private Sub SetDrawDrugDeptEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ҩ���ŵ�Enabled����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-31 11:55:07
    '����:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnHaveDrug As Boolean '����ҩƷ
    
    '���û�����ò��ŵ�ѡ��,��ֱ���˳�
    If fraDrawDept.Visible = False Then cboDrawDept.Enabled = False: lblDrawDrugDept.Enabled = False: Exit Sub
    blnHaveDrug = False
    For i = 1 To mobjBill.Details.Count
        If InStr(1, ",5,6,7,", "," & mobjBill.Details(i).�շ���� & ",") > 0 Then
            blnHaveDrug = True
            Exit For
        End If
    Next
    cboDrawDept.Enabled = blnHaveDrug: lblDrawDrugDept.Enabled = blnHaveDrug
End Sub
Public Function zl��ȡ��ҩ��̬(ByVal lng����ID As Long, Optional ByVal lngRow As Long = -1, Optional blnOnly�г�ҩ As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����Ƿ�¼�����в�ҩ��
    '���:intPage-��ǰ�ڼ�ҳ
    '     blnOnly�г�ҩ-���ж��Ƿ����г�ҩ(���䷽ʱ�ж���Ч):ԭ�����л�ҩ���䷽���Ѿ�����,�Ͳ���Ҫ���
    '     lngRow-��ǰ��������
    '����:
    '����:¼�����в�ҩ��,�򷵻���ҩ��̬����(0-ɢװ,1-��Ƭ,2-����),���򷵻�-1 ��ʾ��û��¼����ҩ��̬��Ŀ
    '����:���˺�
    '����:2010-02-02 11:44:17
    '����:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl��ȡ��ҩ��̬ = -1
    '���δָ��ҳ,���õ�ǰҳ
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly�г�ҩ, ",6,", ",6,7,")
    Err = 0: On Error GoTo ErrHand:
    
    With mobjBill.Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).�շ���� & ",") > 0 And .Item(i).�շ�ϸĿID <> 0 And i <> lngRow And .Item(i).����ID = lng����ID Then
                zl��ȡ��ҩ��̬ = .Item(i).Detail.��ҩ��̬
                Exit Function
            End If
        Next
    End With
ErrHand:
End Function
Private Function zlGetBillOtherRowNumToTal(lng����ID As Long, lng��ҳID As Long, lngϸĿID As Long, _
    Optional blnOnly�������� As Boolean, Optional ByVal lngCurRow As Long = 0) As Double
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ��������������еĺϼƽ��
    '��Σ�lng����ID-����ID
    '         lngCurRow-��ǰ��(Ϊ��ʱ,Ϊ������)
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-05-05 16:09:12
    '˵����29412
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl���� As Double
    
    dbl���� = 0
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .����ID = lng����ID And .��ҳID = lng��ҳID And .�շ�ϸĿID = lngϸĿID And i <> lngCurRow Then
                If blnOnly�������� Then
                    If .���� < 0 And .ִ�в���ID <> 0 Then
                        dbl���� = dbl���� + .���� * .����
                    End If
                Else
                    dbl���� = dbl���� + .���� * .����       '* IIf(InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ, .Detail.סԺ��װ, 1)
                End If
            End If
        End With
    Next
     zlGetBillOtherRowNumToTal = dbl����
End Function

Private Function CheckBillNegative() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���أ��Ϸ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-05-05 17:02:57
    '˵����29412
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, intR As Long
    Dim strItems As String, str���� As String, strValues(0 To 10) As String
    Dim str��λ As String, dbl���� As Double, dbl�ѽ����� As Double, dbl���κϼ� As Double
    Dim strSubTable As String
     
    '����:26951
    If InStr(1, mstrPrivsOpt, ";�������ʲ���鷢����Ŀ;") > 0 Then
        '���ڸ�������ʱ����鱾��סԺ��������Ŀ����,�д�Ȩ��,����¼�벡��δ�������ķ�����Ŀ���г���,�����鱾��סԺ��������Ŀ�������ܳ���
        CheckBillNegative = True: Exit Function
    End If
    
    strItems = ""
    strSubTable = ""
    intR = 0
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And .ִ�в���ID <> 0 Then
                If Len(strItems) > 2000 Then
                    If intR <= 10 Then
                        strValues(intR) = Mid(strItems, 2)
                        strSubTable = strSubTable & " Union ALL " & _
                        "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As ����id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As ��ҳid, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As �շ�ϸĿid, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As ִ�в���id, 0 As ����,0 as ��������" & _
                        " From Table(Cast(f_str2list([" & intR + 2 & "]) As ZLTOOLS.t_strlist))"
                    Else
                        strSubTable = strSubTable & " Union ALL " & _
                        "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As ����id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As ��ҳid, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As �շ�ϸĿid, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As ִ�в���id, 0 As ����,0 as ��������" & _
                        " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
                    End If
                    strItems = "": intR = intR + 1
                End If
                strItems = strItems & "," & .����ID & ";" & .��ҳID & ":" & .�շ�ϸĿID & "_" & .ִ�в���ID & ""
            End If
        End With
    Next
    
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
            "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As ����id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As ��ҳid, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As �շ�ϸĿid, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As ִ�в���id, 0 As ����,0 as ��������" & _
            " From Table(Cast(f_str2list([" & intR + 2 & "]) As ZLTOOLS.t_strlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
            "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As ����id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As ��ҳid, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As �շ�ϸĿid, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As ִ�в���id, 0 As ����,0 as ��������" & _
            " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
        End If
    End If
    CheckBillNegative = True
    If strSubTable = "" Then Exit Function
    strSubTable = Mid(strSubTable, 11)
    
    strSQL = " " & _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */ A.����ID,A.��ҳID, A.�շ�ϸĿID,A.ִ�в���ID,  " & _
    "             Nvl(Sum(Decode(A.��¼����, 2, 1, 3, 1, 0) * Nvl(A.����, 1) * A.����), 0) As ����, " & _
     "            Sum(Decode(nvL(Mod(M.��¼״̬ , 3),1),  0, 1, 1, 1, -1) * Decode(A.����id, Null, 0, 1) * Nvl(����, 1) * ����) As �������� " & _
     "     From סԺ���ü�¼ A, ���˽��ʼ�¼ M " & _
     "     Where  A.����id = M.ID(+)  And A.���ʷ���=1 And A.�۸񸸺� Is Null   " & IIf(gbytBilling = 0, " And A.��¼״̬<>0", "") & _
     "             And (A.����ID,A.��ҳID,A.�շ�ϸĿID,ִ�в���ID,0,0) in (select * From C1) " & _
                    IIf(mstrInNO <> "", " And NO<>[1]", "") & _
     "     Group By A.����ID,A.��ҳID,A.�շ�ϸĿID,A.ִ�в���ID" & _
     "     Union ALL Select * From C1 "
    'strSQL = _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */ A.����ID,A.��ҳID, A.�շ�ϸĿID,A.ִ�в���ID,Sum(Nvl(A.����,1)*A.����) as ����, " & _
    "           Sum(decode(����ID,NULL,0,1)* Nvl(A.����,1)*A.����) as ��������  " & _
    " From  סԺ���ü�¼ A " & _
    " Where ���ʷ���=1 And �۸񸸺� is NULL   " & _
                IIf(gbytBilling = 0, " And ��¼״̬<>0", "") & _
                IIf(mstrInNO <> "", " And NO<>[1]", "") & _
    "           And (A.����ID,A.��ҳID,A.�շ�ϸĿID,ִ�в���ID,0,0) in (select * From C1) " & _
    " Group by A.����ID,A.��ҳID,A.�շ�ϸĿID,A.ִ�в���ID" & _
    " Union ALL Select * From C1"
    strSQL = "" & _
    "   Select ����ID,��ҳID,�շ�ϸĿID,ִ�в���ID,Sum(����) as ����,sum(nvl(��������,0)) as �������� " & _
    "   From (" & strSQL & ") " & _
    "   Group by ����ID,��ҳID,�շ�ϸĿID,ִ�в���ID"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And .ִ�в���ID <> 0 Then
                rsTmp.Filter = " ����ID=" & .����ID & " And ��ҳID = " & .��ҳID & " And �շ�ϸĿID = " & .�շ�ϸĿID & " And ִ�в���ID = " & .ִ�в���ID
                
                If Not rsTmp.EOF Then
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                        str��λ = .Detail.סԺ��λ
                        dbl���� = Nvl(rsTmp!����, 0) / .Detail.סԺ��װ
                        dbl�ѽ����� = Val(Nvl(rsTmp!��������)) / .Detail.סԺ��װ
                    Else
                        str��λ = .Detail.���㵥λ
                        dbl���� = Nvl(rsTmp!����, 0)
                        dbl�ѽ����� = Val(Nvl(rsTmp!��������))
                    End If
                    '���ܴ���������ͬ�ļ�¼
                    '����:29412
                    dbl���κϼ� = Abs(.����) * .����
                    For j = i + 1 To mobjBill.Details.Count
                         If .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And mobjBill.Details(j).����ID = .����ID And mobjBill.Details(j).��ҳID = .��ҳID _
                            And mobjBill.Details(j).���� < 0 And mobjBill.Details(j).ִ�в���ID = .ִ�в���ID Then
                                dbl���κϼ� = dbl���κϼ� + Abs(.����) * .����
                         End If
                    Next
                    '����:32106
                    If dbl���κϼ� > dbl���� - dbl�ѽ����� Then
                        Select Case gbytBillOpt '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
                        Case 0  '����
                            If dbl���κϼ� > dbl���� Then
                                    str���� = GET��������(.ִ�в���ID, mrsUnit)
                                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                        " ���ڿ��������� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                    CheckBillNegative = False: Exit Function
                            End If
                        Case 1   '����
                            str���� = GET��������(.ִ�в���ID, mrsUnit)
                            If dbl���κϼ� > dbl���� Then
                                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                        " ���ڿ��������� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                    CheckBillNegative = False: Exit Function
                            End If
                            
                            If MsgBox("�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                " �а������ѽᲿ��(δ��:" & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "; �ѽ�:" & FormatEx(dbl�ѽ�����, 5) & str��λ & ") ��" & vbCrLf & _
                                " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                CheckBillNegative = False: Exit Function
                            End If
                        Case 2   '��ֹ
                            str���� = GET��������(.ִ�в���ID, mrsUnit)
                            MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                " ���ڿ��������� " & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "��", vbInformation, gstrSysName
                                CheckBillNegative = False: Exit Function
                        End Select
                    End If
                Else
                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]����������Ϊ�㣬�����������", vbInformation, gstrSysName
                    CheckBillNegative = False: Exit Function
                End If
            End If
        End With
    Next
    CheckBillNegative = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub SetStatuPatiInfor(ByVal str���� As String, ByVal dblԤ�� As Double, dblFee As Double, dblʣ�� As Double, Optional dblӦ�� As Double = 0)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����״̬����Ϣ
    '���ƣ����˺�
    '���ڣ�2010-06-23 11:28:31
    '˵����30604
    '------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    picStatuPancl.Visible = False
    '78082:���ϴ�,2014/10/10,Ԥ�������ʾ
    strTemp = str���� & "Ԥ��:" & Format(Val(dblԤ��), "0.00")
    strTemp = strTemp & "/����:" & Format(dblFee, gstrDec)
    strTemp = strTemp & "/ʣ��:" & Format(dblʣ��, "0.00")
    If dblӦ�� <> 0 Then
        strTemp = strTemp & "/Ӧ�տ�:" & Format(dblӦ��, "0.00")
    End If
    
    sta.Panels(3).Text = strTemp
    Call MoveStatuPatiInfor
    If dblʣ�� <= 0 Then
        lblStatuPati.Caption = strTemp
        lblStatuPati.AutoSize = True
        picStatuPancl.Visible = True
    End If
    Err = 0
End Sub
Private Sub MoveStatuPatiInfor()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��ƶ�״̬���Ĳ���Ƿ����Ϣ
    '��Σ�
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-23 13:51:45
    '˵����30604
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With picStatuPancl
        .Left = sta.Panels(3).Left + 50
        .Width = sta.Panels(3).Width - 10
        .Top = Me.ScaleHeight - .Height - 10
    End With
End Sub

Private Sub SetDetailtStock(ByVal lngִ�п���ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ϸ�Ŀ������
    '���ƣ����˺�
    '���ڣ�2010-07-12 14:27:51
    '˵����
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim strҩ��IDs As String, dblStock As Double
    
    '��ȡ���
    '�������ҩƷ������
    If InStr(1, "5,6,7,4", objDetail.���) = 0 Then Exit Sub
    If objDetail.��� = "4" And objDetail.�������� = False Then Exit Sub
    If objDetail.��� = "4" Then
        '����
        dblStock = GetStock(objDetail.ID, lngִ�п���ID)
        objDetail.��� = dblStock
        Exit Sub
    End If
    
    If Not gbln���뷢ҩ Then
        dblStock = GetStock(objDetail.ID, lngִ�п���ID)
        If gblnסԺ��λ Then
            dblStock = dblStock / objDetail.סԺ��װ
        End If
        objDetail.��� = dblStock  '��¼��ǰ��ҩƷ���
        Exit Sub
    End If
    strҩ��IDs = Decode(mobjDetail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
    If strҩ��IDs <> "" Then
        dblStock = GetMultiStock(mobjDetail.ID, strҩ��IDs)
        If gblnסԺ��λ Then
            dblStock = dblStock / mobjDetail.סԺ��װ
        End If
        mobjDetail.��� = dblStock
    End If
End Sub
Private Function ReadDrugAndStuffStock(ByVal lng�ⷿID As Long, ByRef objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ�������ϵĿ����Ϣ
    '���:lng�ⷿID-�ⷿID
    '����:objDetail-Detail����
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-10 09:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblStock As Double, strҩ��IDs As String
    
    On Error GoTo errHandle
    If objDetail Is Nothing Then Exit Function
    '��ҩƷ���������ϵģ�ֱ�ӷ���True
    If InStr(",5,6,7,4,", objDetail.���) = 0 Then ReadDrugAndStuffStock = True: Exit Function
    If objDetail.��� = "4" And objDetail.�������� = False Then ReadDrugAndStuffStock = True: Exit Function
   
    If objDetail.��� = "4" And objDetail.�������� Then
        dblStock = GetStock(objDetail.ID, lng�ⷿID)
        objDetail.��� = dblStock
        Call ShowStock(objDetail.����, objDetail.���)
        ReadDrugAndStuffStock = True: Exit Function
    End If
    If InStr(",5,6,7,", objDetail.���) > 0 Then
        '��ǰ��ҩƷ���
        If Not gbln���뷢ҩ Then
            dblStock = GetStock(objDetail.ID, lng�ⷿID)
            If gblnסԺ��λ Then
                dblStock = dblStock / objDetail.סԺ��װ
            End If
            objDetail.��� = dblStock
            Call ShowStock(objDetail.����, objDetail.���)
        Else
            strҩ��IDs = Decode(objDetail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
            If strҩ��IDs <> "" Then
                dblStock = GetMultiStock(objDetail.ID, strҩ��IDs)
                
                If dblStock = 0 And gblnStock Then
                    MsgBox "[" & objDetail.���� & "]�Ŀ��ÿ��Ϊ��!", vbInformation, gstrSysName
                    Exit Function
                End If
                If gblnסԺ��λ Then
                    dblStock = dblStock / objDetail.סԺ��װ
                End If
                objDetail.��� = dblStock
                Call ShowStock(objDetail.����, objDetail.���)
            End If
        End If
    End If
    ReadDrugAndStuffStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
