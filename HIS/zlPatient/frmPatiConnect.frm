VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmPatiConnect 
   Caption         =   "������ݹ���"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatiConnect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12465
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   3735
      TabIndex        =   37
      Top             =   1440
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   660
         Left            =   720
         TabIndex        =   38
         Top             =   480
         Width           =   1320
         _Version        =   589884
         _ExtentX        =   2328
         _ExtentY        =   1164
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   3615
         TabIndex        =   41
         Top             =   3720
         Width           =   3615
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ȡ������"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   2880
            TabIndex        =   44
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "δ����"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   1800
            TabIndex        =   43
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ѹ���"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   600
            TabIndex        =   42
            Top             =   120
            Width           =   630
         End
         Begin VB.Image img 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   2
            Left            =   2400
            Picture         =   "frmPatiConnect.frx":6852
            Top             =   0
            Width           =   480
         End
         Begin VB.Image img 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   1
            Left            =   1200
            Picture         =   "frmPatiConnect.frx":711C
            Top             =   0
            Width           =   480
         End
         Begin VB.Image img 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   0
            Picture         =   "frmPatiConnect.frx":79E6
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   3960
      ScaleHeight     =   6135
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   360
      Width           =   7815
      Begin VB.Frame fraInfo 
         BackColor       =   &H8000000E&
         Caption         =   "������Ϣ"
         Height          =   6045
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7575
         Begin VB.TextBox txtAddTime 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2025
         End
         Begin VB.TextBox txtסԺ���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   4320
            Width           =   2025
         End
         Begin VB.TextBox txt��λ 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   3960
            Width           =   2025
         End
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   3600
            Width           =   2025
         End
         Begin VB.TextBox txt��ͥ��ַ 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   3240
            Width           =   2025
         End
         Begin VB.TextBox txt�����ص� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2880
            Width           =   2025
         End
         Begin VB.TextBox txt���֤�� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2520
            Width           =   2025
         End
         Begin VB.TextBox txt��� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2025
         End
         Begin VB.TextBox txtְҵ 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2025
         End
         Begin VB.TextBox txt����״�� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2025
         End
         Begin VB.TextBox txtѧ�� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2025
         End
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2025
         End
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2025
         End
         Begin VB.TextBox txt�������� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2025
         End
         Begin VB.TextBox txt�Ա� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   2025
         End
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   720
            Width           =   2025
         End
         Begin VB.TextBox txtסԺ�� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   2025
         End
         Begin VB.TextBox txt״̬ 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblAddTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ǽ�ʱ��"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4035
            TabIndex        =   40
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label lblסԺ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   4245
            TabIndex        =   36
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lbl״̬ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "״̬"
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   660
            TabIndex        =   35
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lblסԺ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   240
            TabIndex        =   34
            Top             =   4320
            Width           =   840
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   660
            TabIndex        =   33
            Top             =   3960
            Width           =   420
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   660
            TabIndex        =   32
            Top             =   3600
            Width           =   420
         End
         Begin VB.Label lbl��ͥ��ַ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��סַ"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   450
            TabIndex        =   31
            Top             =   3240
            Width           =   630
         End
         Begin VB.Label lbl�����ص� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ص�"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   30
            Top             =   2880
            Width           =   840
         End
         Begin VB.Label lbl���֤�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���֤��"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   2520
            Width           =   840
         End
         Begin VB.Label lblְҵ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ְҵ"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   615
            TabIndex        =   28
            Top             =   1800
            Width           =   420
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   27
            Top             =   1800
            Width           =   420
         End
         Begin VB.Label lbl����״�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����״��"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   26
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label lblѧ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ѧ��"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   660
            TabIndex        =   25
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   24
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   23
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lbl�Ա� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   4455
            TabIndex        =   21
            Top             =   720
            Width           =   420
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H00333333&
            Height          =   210
            Left            =   660
            TabIndex        =   20
            Top             =   720
            Width           =   420
         End
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1588
            MinWidth        =   1587
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18830
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   988
            MinWidth        =   988
            Text            =   "�༭"
            TextSave        =   "�༭"
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
   Begin MSComctlLib.ImageList ilsPati 
      Left            =   1560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":82B0
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":EB12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":15374
            Key             =   "link"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":1BBD6
            Key             =   "linkAdd"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":22438
            Key             =   "linkNew"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":28C9A
            Key             =   "linkdelete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":2F4FC
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":35D5E
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":362F8
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiConnect.frx":36892
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   840
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatiConnect.frx":3D0F4
      Left            =   3240
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPatiConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object
Private mrsPati    As ADODB.Recordset   '"����ID", "����ID", "����", "�Ա�", "����", "��������", "���֤��", "��ͥ��ַ"

Private mlngPatiId As Long
Private mblnUndo   As Boolean
Private mbytEditState As Byte    '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�༭״̬
Private mlngLinkID    As Long       '��¼ѡ���еĹ���ID
Private mstrPrivs As String
Private mstrFilter As String
Private mbytFunc As Byte            '=1 ���õ��ã����벡��ID�Զ�������ͬ��ݲ��ˣ����û������Ƿ��Զ�������

Private Const M_BGK_CORLOR As Long = &HEBFFFF

Private Enum PATI_COLUMN
    COL_ͼ�� = 0
    COL_����
    COL_�Ա�
    COL_����
    COL_��������
    COL_���֤��
    COL_��ͥ��ַ
    COL_�Ǽ�ʱ��
    '������
    COL_����ID
    COL_����Id
    COL_����
    COL_����
    COL_EDIT            '0-ԭʼ;1-�Զ�����;2-���ӹ���;3-ȡ������
End Enum

Private Enum E_EDIT
    E_LINKLOAD = 0
    E_LINKAUTO
    E_LINKADD
    E_LINKCANCEL
End Enum

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Manage_RelatingPatiet * 10# + 1 '�Զ�����
        If UpdatePati(E_EDIT.E_LINKAUTO) Then mbytEditState = 1
    Case conMenu_Manage_RelatingPatiet * 10# + 2 '���ӹ���
        If UpdatePati(E_EDIT.E_LINKADD) Then mbytEditState = 1
    Case conMenu_Manage_RelatingPatiet * 10# + 3    'ȡ������
        If UpdatePati(E_EDIT.E_LINKCANCEL) Then mbytEditState = 1
    Case conMenu_Edit_Save
        Call SaveData
        mbytEditState = 0
        If mbytFunc = 1 Then Unload Me: Exit Sub
    Case conMenu_Edit_Untread
        Call LoadPati(E_LINKLOAD)
        mbytEditState = 0
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    picLeft.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    staThis.Width = lngRight - lngLeft
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
    Case conMenu_Manage_RelatingPatiet * 10# + 1, conMenu_Manage_RelatingPatiet * 10# + 2
        Control.Enabled = mbytEditState = 0
    Case conMenu_Manage_RelatingPatiet * 10# + 3
        Control.Enabled = mlngLinkID <> 0 And mbytEditState = 0
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = mbytEditState = 1 Or (Control.ID = conMenu_Edit_Save And mbytFunc = 1)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As XtremeDockingPane.Pane
    Call InitCommandBar
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    
    Set objPane = Me.dkpMain.CreatePane(1, 320, 400, DockLeftOf, Nothing)
    objPane.Title = "���������б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Call InitReportColumn
    If mbytFunc = 0 Then
        Call LoadPati(E_LINKLOAD)
    Else
        Call LoadPati(E_LINKAUTO)
    End If
    img(0).Picture = ilsPati.ListImages("link").Picture
    img(1).Picture = ilsPati.ListImages("linkAdd").Picture
    img(2).Picture = ilsPati.ListImages("linkdelete").Picture
    lblNote(0).Caption = "�ѹ���"
    lblNote(1).Caption = "������"
    lblNote(2).Caption = "��ȡ��"
    
    Call RestoreWinState(Me, App.ProductName, , True)
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
End Sub


Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    If mbytFunc = 0 Then
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet * 10# + 1, "�Զ�����")
            objControl.IconId = conMenu_Kss_Adjustment
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet * 10# + 2, "���ӹ���")
            objControl.IconId = conMenu_Kss_Grant
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet * 10# + 3, "ȡ������")
            objControl.IconId = conMenu_Kss_Cancellation
            
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "ȷ��"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
            Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True
        End With
        
        objBar.EnableDocking xtpFlagHideWrap
        objBar.ContextMenuPresent = False
        For Each objControl In objBar.Controls
            If objControl.type <> xtpControlCustom And objControl.type <> xtpControlLabel Then
                objControl.Style = xtpButtonIconAndCaption
            End If
        Next
        
        With cbsMain.KeyBindings
            .Add FALT, vbKeyQ, conMenu_File_Exit
            .Add FALT, vbKeyS, conMenu_Edit_Save
        End With
    Else
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "ȷ��")
            Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True
        End With
        
        objBar.EnableDocking xtpFlagHideWrap
        objBar.ContextMenuPresent = False
        For Each objControl In objBar.Controls
            If objControl.type <> xtpControlCustom And objControl.type <> xtpControlLabel Then
                objControl.Style = xtpButtonIconAndCaption
            End If
        Next
        
        With cbsMain.KeyBindings
            .Add FALT, vbKeyQ, conMenu_File_Exit
        End With
    End If
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        .Columns.DeleteAll
        Set objCol = .Columns.Add(COL_ͼ��, "", 20, False)
        Set objCol = .Columns.Add(COL_����, "����", 80, True)
        Set objCol = .Columns.Add(COL_�Ա�, "�Ա�", 45, True)
        Set objCol = .Columns.Add(COL_����, "����", 45, True)
        Set objCol = .Columns.Add(COL_��������, "��������", 80, True)
        Set objCol = .Columns.Add(COL_���֤��, "���֤��", 150, True)
        Set objCol = .Columns.Add(COL_��ͥ��ַ, "��סַ", 180, True)
        Set objCol = .Columns.Add(COL_�Ǽ�ʱ��, "�Ǽ�ʱ��", 150, True)
        
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_����Id, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_EDIT, "�༭", 0, False): objCol.Visible = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĺ�������..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        mblnUndo = True
        .MultipleSelection = False '������SelectionChanged�¼�
         mblnUndo = False
        .ShowItemsInGroups = False
        .SetImageList Me.ilsPati
    End With
End Sub

Public Function ShowMe(ByRef frmParent As Object, ByVal strPrivs As String, ByVal lngPatiID As Long, Optional ByVal bytFunc As Byte = 0) As Boolean
'����:��ݹ���
    Dim rsPati As ADODB.Recordset
    
    If lngPatiID = 0 Then Exit Function
    Set mfrmParent = frmParent
    mstrPrivs = strPrivs
    mstrFilter = ""
    mbytFunc = bytFunc
    mlngPatiId = lngPatiID
    If mbytFunc = 1 Then
        Set mrsPati = GetPatiLinked(mlngPatiId)
        Set mrsPati = zlDatabase.CopyNewRec(mrsPati, , , Array("EDIT", adInteger, 2, Empty))
        mrsPati.Filter = "����ID =" & mlngPatiId
        If Not mrsPati.EOF Then
            mstrFilter = mrsPati!���� & "|" & mrsPati!���� & "|" & mrsPati!�Ա� & _
            "|" & mrsPati!���� & "|" & Format(mrsPati!�������� & "", "YYYY-MM-DD") & "|" & mrsPati!���֤��
        End If
        Set rsPati = GetPatiSimilar(mstrFilter)
        If Not AppendPatiSimilar(rsPati, E_LINKAUTO) Then Exit Function
    End If
    Me.Show 1, frmParent
    ShowMe = True
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SaveWinState(Me, App.ProductName)
    Set mrsPati = Nothing
End Sub

Private Sub picLeft_Resize()
    
    On Error Resume Next
    fraInfo.Move 120, 120, picLeft.ScaleWidth - 240, picLeft.ScaleHeight - 240
End Sub

Private Sub picNote_Resize()
    On Error Resume Next
    img(0).Move 120, 120, 480, 480
    img(1).Move img(0).Left + 1380, 120, 480, 480
    img(2).Move img(1).Left + 1380, 120, 480, 480
    lblNote(0).Move img(0).Left + 360, 120
    lblNote(1).Move img(1).Left + 360, 120
    lblNote(2).Move img(2).Left + 360, 120
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    rptPati.Move 0, 0, picPati.ScaleWidth, picPati.ScaleHeight - picNote.Height
    picNote.Move 0, picPati.Height - picNote.Height, picPati.ScaleWidth
End Sub

Private Function LoadPati(ByVal bytFunc As Byte, Optional ByVal lngPatiID As Long) As Boolean
'����:���ز�����ݹ����б�
'����:
'bytFunc=0-ԭʼ����;1-�Զ�����;2-���ӹ���;3-ȡ������
'lngPatiID-����ID
'   strSimilar '����|����|�Ա�|����|��������(To_Date('2015/4/30', 'YYYY-MM-DD'))|���֤��
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim lngSelID As Long
    
    On Error GoTo errH
    If mbytFunc = 0 Then
        If bytFunc = E_LINKLOAD Then
            Set rsPati = GetPatiLinked(mlngPatiId)
        ElseIf bytFunc = E_LINKADD Then
            Set rsPati = GetPatiLinked(lngPatiID)
        ElseIf bytFunc = E_LINKAUTO Then
            Set rsPati = GetPatiSimilar(mstrFilter)
        End If
        
        If bytFunc <> E_LINKCANCEL Then
            If rsPati.EOF Then
                If bytFunc = E_LINKAUTO Then
                    MsgBox "δ����������ƵĲ�����Ϣ��", vbInformation, gstrSysName
                Else
                    MsgBox "δ���ָò��˵������Ϣ��", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
        
        If bytFunc = E_LINKLOAD Then
            Set mrsPati = zlDatabase.CopyNewRec(rsPati, , , Array("EDIT", adInteger, 2, Empty))      'COPY���ں�����ɾ��
        ElseIf bytFunc = E_LINKADD Or bytFunc = E_LINKAUTO Then
            If Not AppendPatiSimilar(rsPati, bytFunc) Then Exit Function
        ElseIf bytFunc = E_LINKCANCEL Then
            mrsPati.Filter = ""
        End If
    Else
        mrsPati.Filter = ""
    End If
    mrsPati.Sort = "�Ǽ�ʱ�� ASC"
    '���ز����б�
    Call ClearPatiInfo
    rptPati.Tag = ""
    rptPati.Records.DeleteAll
    rptPati.SortOrder.DeleteAll
    For i = 1 To mrsPati.RecordCount
        Set objRecord = rptPati.Records.Add()
        Set objItem = objRecord.AddItem("")
        If CLng(mrsPati!����ID) > 0 And Val(mrsPati!EDIT & "") = E_LINKCANCEL Then
            objItem.Icon = ilsPati.ListImages("linkdelete").Index - 1
        ElseIf CLng(mrsPati!����ID) > 0 Then
            objItem.Icon = ilsPati.ListImages("link").Index - 1
        ElseIf CLng(mrsPati!����ID) = 0 Then
            objItem.Icon = ilsPati.ListImages("linkAdd").Index - 1
        End If
        objRecord.AddItem mrsPati!���� & ""
        objRecord.AddItem mrsPati!�Ա� & ""
        objRecord.AddItem mrsPati!���� & ""
        
        objRecord.AddItem Format(mrsPati!��������, "YYYY-MM-DD")
        objRecord.AddItem mrsPati!���֤�� & ""
        objRecord.AddItem Nvl(mrsPati!��ͥ��ַ, "δ�Ǽ�")
        objRecord.AddItem Format(mrsPati!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS")
        '������
        objRecord.AddItem CLng(mrsPati!����ID)
        objRecord.AddItem CLng(mrsPati!����ID)
        objRecord.AddItem mrsPati!���� & ""
        objRecord.AddItem mrsPati!���� & ""
        objRecord.AddItem Nvl(mrsPati!EDIT, "0")
        If CLng(mrsPati!����ID) = mlngPatiId And bytFunc = E_LINKLOAD And mstrFilter = "" Then
            mstrFilter = mrsPati!���� & "|" & mrsPati!���� & "|" & mrsPati!�Ա� & _
            "|" & mrsPati!���� & "|" & Format(mrsPati!�������� & "", "YYYY-MM-DD") & "|" & mrsPati!���֤��
        End If
        mrsPati.MoveNext
    Next
    
    If bytFunc = E_LINKAUTO Then
        For i = 0 To rptPati.Records.Count - 1
            If Val(rptPati.Records.Record(i).Item(COL_EDIT).Value) = E_LINKAUTO Then
                For j = COL_���� To COL_�Ǽ�ʱ��
                    rptPati.Records.Record(i).Item(j).BackColor = M_BGK_CORLOR
                Next
            End If
        Next
    End If
    '��ǰ��������Ӵ�
    For i = 0 To rptPati.Records.Count - 1
        If Val(rptPati.Records.Record(i).Item(COL_����Id).Value) = mlngPatiId Then
            For j = COL_���� To COL_�Ǽ�ʱ��
                rptPati.Records.Record(i).Item(j).Bold = True
            Next
            Exit For
        End If
    Next
    rptPati.Populate
    
    If bytFunc = E_LINKLOAD Or bytFunc = E_LINKAUTO Then
        lngSelID = mlngPatiId
    ElseIf bytFunc = E_LINKADD Or bytFunc = E_LINKCANCEL Then
        lngSelID = lngPatiID
    End If
    For i = 0 To rptPati.Records.Count - 1
        If Val(rptPati.Records.Record(i).Item(COL_����Id).Value) = lngSelID Then
            Set rptPati.FocusedRow = rptPati.Rows(i)
            Exit For
        End If
    Next
    LoadPati = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPatiInfo(lngID As Long) As Boolean
'���ܣ���ʾһ��������Ϣ
'������lngID=����ID
    Dim rsTmp As New ADODB.Recordset, rsPati As ADODB.Recordset
    Dim strSQL As String
    Dim strסԺ�� As String, str����� As String
    Dim strJsonAsk As String, strJsonOut As String
    Dim colReturn As Collection
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "Select ����ID,�����,סԺ��,���￨��,����,�Ա�,����,��������,�����ص�,���֤��,���,ְҵ,����,����,����,ѧ��,����״��,��ͥ��ַ,��ͥ�绰,�Ǽ�ʱ��" & _
             "  From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        
    If rsTmp.EOF Then
        MsgBox "δ���ָò��˵������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If

    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt�Ա�.Text = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
    txt��������.Text = Format(IIf(IsNull(rsTmp!��������), "", rsTmp!��������), "yyyy��MM��dd��")
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txtѧ��.Text = IIf(IsNull(rsTmp!ѧ��), "", rsTmp!ѧ��)
    txt���.Text = IIf(IsNull(rsTmp!���), "", rsTmp!���)
    txtְҵ.Text = IIf(IsNull(rsTmp!ְҵ), "", rsTmp!ְҵ)
    txt���֤��.Text = IIf(IsNull(rsTmp!���֤��), "", rsTmp!���֤��)
    txt�����ص�.Text = IIf(IsNull(rsTmp!�����ص�), "", rsTmp!�����ص�)
    txt��ͥ��ַ.Text = IIf(IsNull(rsTmp!��ͥ��ַ), "", rsTmp!��ͥ��ַ)
    txt����״��.Text = IIf(IsNull(rsTmp!����״��), "", rsTmp!����״��)
    txtAddTime.Text = Format(Nvl(rsTmp!�Ǽ�ʱ��), "YYYY-MM-DD HH:MM:SS")
    str����� = IIf(IsNull(rsTmp!�����), "", rsTmp!�����)
    strסԺ�� = IIf(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
    strSQL = "Select a.����id, a.��ҳid,a.סԺ���� " & vbNewLine & _
            "From ������Ϣ a " & _
            "Where a.����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    Set rsPati = InitRecordset("��Ժ����||30,סԺ��||18,��Ժ����||50,����id|adInteger|18,��ҳid|adInteger|18,סԺ����|adInteger|3,����||100")
    
    strJsonAsk = "{""input"":{""query_type"":1,""pati_pageids"":""" & lngID & """}}"
    If CallService("Zl_Cissvr_Getpatipageinfo", strJsonAsk, strJsonOut) Then
        Set colReturn = gobjService.GetJsonListValue("output.page_list")
    End If
    If Not rsTmp.EOF And Not colReturn Is Nothing Then
        For i = 1 To colReturn.Count
            If zval(rsTmp!����ID & "") = zval(gobjService.GetCollValue(colReturn, i, "_pati_id")) And zval(rsTmp!��ҳID & "") = zval(gobjService.GetCollValue(colReturn, i, "_pati_pageid")) Then
                With rsPati
                    .AddNew Array("��Ժ����", "סԺ��", "��Ժ����", "����id", "��ҳid", "סԺ����", "����"), Array(Format(gobjService.GetCollValue(colReturn, i, "_adtd_time") & "", "YYYY-MM-DD HH:MM:SS"), gobjService.GetCollValue(colReturn, i, "_inpatient_num"), _
                    gobjService.GetCollValue(colReturn, i, "_pati_bed"), gobjService.GetCollValue(colReturn, i, "_pati_id"), gobjService.GetCollValue(colReturn, i, "_pati_pageid"), _
                    Val(rsTmp!סԺ���� & ""), gobjService.GetCollValue(colReturn, i, "_pati_dept_name"))
                    .Update
                End With
            End If
        Next
        If Not rsPati.EOF Then rsPati.MoveFirst
    End If
    Set rsTmp = rsPati.Clone
    If rsTmp.EOF Then
        If glngSys Like "8??" Then
            txt״̬.Text = "����"
        Else
            txt״̬.Text = "����"
        End If
        lblסԺ��.Caption = "�����"
        txtסԺ��.Text = IIf(str����� = "", "", str�����)
        txt����.Text = ""
        txt��λ.Text = ""
        txtסԺ����.Text = ""
    Else
        txt״̬.Text = IIf(IsNull(rsTmp!��Ժ����), "��Ժ", "��Ժ")
        lblסԺ��.Caption = "סԺ��"
        txtסԺ��.Text = IIf(strסԺ�� = "", "", strסԺ��)
        txt����.Text = rsTmp!����
        txt��λ.Text = IIf(IsNull(rsTmp!��Ժ����), "��ͥ", rsTmp!��Ժ����)
        txtסԺ����.Text = Nvl(rsTmp!סԺ����)
    End If
    
    ShowPatiInfo = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'���ܣ����һ��������Ϣ
'������x=�ؼ�����,0=Դ����,1=Ŀ�겡��
    txt����.Text = ""
    txt�Ա�.Text = ""
    txt��������.Text = ""
    txt����.Text = ""
    txt����.Text = ""
    txtѧ��.Text = ""
    txt���.Text = ""
    txtְҵ.Text = ""
    txt���֤��.Text = ""
    txt�����ص�.Text = ""
    txt��ͥ��ַ.Text = ""
    txt����״��.Text = ""
    txt״̬.Text = ""
    lblסԺ��.Caption = "סԺ��"
    txtסԺ��.Text = ""
    txt����.Text = ""
    txt��λ.Text = ""
    txtסԺ����.Text = ""
    txtAddTime.Text = ""
End Sub

Private Sub rptPati_SelectionChanged()
'����:
    If mblnUndo Then Exit Sub
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '���������
    Me.staThis.Panels(2).Text = ""
    With rptPati.SelectedRows(0)
        If .GroupRow Then
            Call ClearPatiInfo
        Else
            If rptPati.Tag = .Record(COL_����Id).Value Then Exit Sub
            mlngLinkID = Val(.Record(COL_����ID).Value & "")
            Call ShowPatiInfo(Val(.Record(COL_����Id).Value & ""))
            rptPati.Tag = .Record(COL_����Id).Value
        End If
        If Val(.Record(COL_EDIT).Value) = E_LINKAUTO Then
            Me.staThis.Panels(2).Text = "���֤����ͬ������|�Ա�|����|��������|��������ͬ�Ĳ��ˡ�"
        End If
    End With
End Sub


Private Function UpdatePati(ByVal bytFunc As Byte) As Boolean
'����:���ӹ���
'����:bytFunc 1-�Զ�����;2-���ӹ���;3-ȡ������
    Dim objFrm As New frmPatiSel
    If bytFunc = E_EDIT.E_LINKADD Then
        objFrm.mstrPrivs = mstrPrivs
        objFrm.Show 1, Me
        If objFrm.mlng����ID <> 0 Then
            mrsPati.Filter = "����ID=" & objFrm.mlng����ID
            If mrsPati.RecordCount > 0 Then
                MsgBox "�ò����Ѿ��ڡ����������б��У����������ӣ�", vbInformation + vbOKOnly, gstrSysName: Exit Function
            End If
            UpdatePati = LoadPati(bytFunc, objFrm.mlng����ID)
        End If
    ElseIf bytFunc = E_EDIT.E_LINKAUTO Then
        UpdatePati = LoadPati(bytFunc)
    ElseIf bytFunc = E_EDIT.E_LINKCANCEL Then
        With rptPati.SelectedRows(0)
            mrsPati.Filter = "����ID=" & .Record(COL_����Id).Value
            mrsPati!EDIT = E_EDIT.E_LINKCANCEL
            UpdatePati = LoadPati(bytFunc, Val(.Record(COL_����Id).Value))
        End With
    End If
End Function

Private Function SaveData() As Boolean
'����:�������
    Dim strTime As String
    Dim strPatiID As String
    Dim lngLinKID As Long
    Dim arrSQL As Variant
    Dim blnTrans As Boolean
    Dim i As Long
    
    On Error GoTo errH
    arrSQL = Array()
    mrsPati.Filter = "EDIT=" & E_LINKCANCEL
    If mrsPati.RecordCount > 0 Then
        If Val(mrsPati!����ID & "") > 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������ݹ���_Update(1," & mrsPati!����ID & ",'" & mrsPati!����ID & "')"
        End If
    Else
        mrsPati.Filter = ""
        mrsPati.Sort = "����ID ASC"
        For i = 1 To mrsPati.RecordCount
            If Val(mrsPati!����ID & "") <> 0 Then
                If lngLinKID = 0 Then lngLinKID = Val(mrsPati!����ID & "")
                If Val(mrsPati!����ID & "") <> lngLinKID Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_������ݹ���_Update(2," & lngLinKID & ",'" & mrsPati!����ID & "')"
                End If
            Else
                strPatiID = strPatiID & "," & mrsPati!����ID
            End If
            mrsPati.MoveNext
        Next
        If strPatiID <> "" Then
            strPatiID = Mid(strPatiID, 2)
            strTime = "TO_DATE('" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������ݹ���_Update(0," & lngLinKID & ",'" & strPatiID & "','" & UserInfo.���� & "'," & strTime & ")"
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    SaveData = LoadPati(E_LINKLOAD)
    Exit Function
errH:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiSimilar(ByVal strSimilar As String) As ADODB.Recordset
'����: strSimilar '����|����|�Ա�|����|��������(To_Date('2015/4/30', 'YYYY-MM-DD'))|���֤��
    Dim arrTmp As Variant
    Dim strSQL As String
    
    arrTmp = Split(strSimilar, "|")
    
    strSQL = "Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.���֤��, a.��ͥ��ַ, a.����, a.����, 0 As ����id, a.�Ǽ�ʱ��" & vbNewLine & _
            "From ������Ϣ A" & vbNewLine & _
            "Where ((a.���� = [1] And a.���� = [2] And a.�Ա� = [3] And a.���� = [4] And a.�������� = To_Date([5], 'YYYY-MM-DD')) Or" & vbNewLine & _
            "      a.���֤�� = [6]) And Not Exists (Select * From ������ݹ��� B Where b.����id = a.����id)" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.���֤��, a.��ͥ��ַ, a.����, a.����, b.����id, a.�Ǽ�ʱ��" & vbNewLine & _
            "From ������Ϣ A, ������ݹ��� B, ������ݹ��� C, ������Ϣ D" & vbNewLine & _
            "Where a.����id = b.����id And b.����id = c.����id And c.����id = d.����id And" & vbNewLine & _
            "      ((d.���� = [1] And d.���� = [2] And d.�Ա� = [3] And d.���� = [4] And d.�������� = To_Date([5], 'YYYY-MM-DD')) Or" & vbNewLine & _
            "      d.���֤�� = [6])"
    On Error GoTo errH
    Set GetPatiSimilar = zlDatabase.OpenSQLRecord(strSQL, "GetPatiSimilar", (arrTmp(0)), (arrTmp(1)), (arrTmp(2)), (arrTmp(3)), (arrTmp(4)), (arrTmp(5)))
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiLinked(ByVal lngPatiID As Long) As ADODB.Recordset
'����:ͨ������ID�����ѹ����Ĳ���
    Dim strSQL As String
    
    strSQL = "Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.���֤��, a.��ͥ��ַ, a.����, a.����, 0 As ����id, a.�Ǽ�ʱ��" & vbNewLine & _
            "From ������Ϣ A" & vbNewLine & _
            "Where ����id = [1] And Not Exists (Select * From ������ݹ��� Where ����id = [1])" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.���֤��, a.��ͥ��ַ, a.����, a.����, b.����id, a.�Ǽ�ʱ��" & vbNewLine & _
            "From ������Ϣ A, ������ݹ��� B, ������ݹ��� C" & vbNewLine & _
            "Where a.����id = b.����id And b.����id = c.����id And c.����id = [1]"
    On Error GoTo errH
    Set GetPatiLinked = zlDatabase.OpenSQLRecord(strSQL, "GetPatiLinked", lngPatiID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AppendPatiSimilar(ByVal rsPati As ADODB.Recordset, ByVal bytFunc As Byte) As Boolean
'����:׷�����Ʋ���
    Dim blnFind As Boolean
    Dim i As Long
    
    For i = 1 To rsPati.RecordCount
        mrsPati.Filter = "����ID=" & rsPati!����ID
        If mrsPati.RecordCount = 0 Then
            mrsPati.AddNew Array("����ID", "����ID", "����", "�Ա�", "����", "��������", "���֤��", "��ͥ��ַ", "����", "����", "�Ǽ�ʱ��", "EDIT"), _
            Array(rsPati!����ID, rsPati!����ID, rsPati!����, rsPati!�Ա�, rsPati!����, rsPati!��������, rsPati!���֤��, rsPati!��ͥ��ַ, _
            rsPati!����, rsPati!����, rsPati!�Ǽ�ʱ��, bytFunc)
            blnFind = True
        End If
        rsPati.MoveNext
    Next
    mrsPati.Filter = ""
    If Not blnFind And bytFunc = E_LINKAUTO Then
        MsgBox "δ�����������ͬʱ��δ������ݵĲ�����Ϣ��", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    AppendPatiSimilar = True
End Function
                
